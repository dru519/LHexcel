#!/usr/bin/env python3
"""Assemble release evidence only when every source is identity-bound.

The builder does not decide whether a release passes.  It prevents a caller
from composing unrelated PASS labels by requiring the frozen source, actual
artifact bytes, Windows inspection, independence verdict, static gates and
three final reviews to name the same commit, source tree and artifact.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import os
import re
import tempfile
from copy import deepcopy
from pathlib import Path
from typing import Any, Mapping

try:
    from .compare_independence import validate_approved_lineage_policy, validate_approved_lineage_source
except ImportError:  # direct script execution
    from compare_independence import validate_approved_lineage_policy, validate_approved_lineage_source


SHA40 = re.compile(r"^[0-9a-f]{40}$")
SHA64 = re.compile(r"^[0-9a-f]{64}$")
STATIC_GATES = (
    "spec_lock",
    "clean_build",
    "source_package",
    "tests",
    "feature_policy",
    "provenance",
    "detector",
)
PRODUCT_EVIDENCE = (
    "ProductRibbon.Package.json",
    "ProductRibbon.RibbonXml.json",
    "ProductRibbon.OfficeIdentity.json",
    "ProductRibbon.Ui.Current200.json",
    "ProductRibbon.Fast.json",
    "ProductRibbon.Guarded.json",
    "ProductRibbon.Planned.json",
    "ProductRibbon.UserForms.json",
)
PRODUCT_UI_EVIDENCE = ("ProductUi.Current200.json", "ProductUi.UserForms.json")


class EvidenceError(RuntimeError):
    """Raised when evidence is missing, malformed or cross-bound."""


def _read(path: Path, label: str) -> dict[str, Any]:
    try:
        value = json.loads(path.read_text(encoding="utf-8-sig"))
    except (OSError, UnicodeDecodeError, json.JSONDecodeError) as exc:
        raise EvidenceError(f"{label} is not readable JSON: {exc}") from exc
    if not isinstance(value, dict):
        raise EvidenceError(f"{label} must contain an object")
    return value


def _sha(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def _bundle_digest(records: list[Mapping[str, Any]]) -> str:
    payload = "".join(f"{row['path']}\n{row['sha256']}\n" for row in sorted(records, key=lambda row: str(row.get("path")).encode("utf-8"))).encode()
    return hashlib.sha256(payload).hexdigest()


def _verify_bundle(freeze: Mapping[str, Any], artifact: Path, bundle_root: Path | None = None) -> tuple[list[dict[str, str]], str]:
    records = freeze.get("product_files")
    if not isinstance(records, list) or not records:
        raise EvidenceError("freeze manifest product_files are missing")
    base = (bundle_root or artifact.parent).resolve()
    normalized: list[dict[str, str]] = []
    for item in records:
        if not isinstance(item, Mapping) or set(item) != {"path", "sha256"} or not isinstance(item.get("path"), str) or not isinstance(item.get("sha256"), str) or not SHA64.fullmatch(item["sha256"]):
            raise EvidenceError("product file record is malformed")
        relative = item["path"]
        rel = Path(relative)
        if "\\" in relative or re.search(r'[:*?"<>|]', relative) or rel.is_absolute() or any(part in {"", ".", ".."} for part in relative.split("/")):
            raise EvidenceError("product file path is unsafe")
        path = base / rel
        if not path.is_file() or _sha(path) != item["sha256"]:
            raise EvidenceError(f"product file missing or tampered: {item['path']}")
        normalized.append({"path": rel.as_posix(), "sha256": item["sha256"]})
    paths = [record["path"] for record in normalized]
    if paths != sorted(paths, key=lambda value: value.encode("utf-8")) or len(set(paths)) != len(paths):
        raise EvidenceError("product_files must be unique and UTF-8 ordinal sorted")
    expected = [{"path": "Product.xlam", "sha256": _sha(artifact)}]
    if normalized != expected:
        raise EvidenceError("product_files must contain exactly the single Product XLAM")
    digest = _bundle_digest(normalized)
    expected_digest = freeze.get("product_bundle_sha256")
    if expected_digest != digest:
        raise EvidenceError("bundle digest does not match freeze manifest")
    return normalized, digest


def _identity(document: Mapping[str, Any], label: str) -> tuple[str, str, str]:
    nested = document.get("candidate_identity")
    source = nested if isinstance(nested, Mapping) else document
    commit = source.get("frozen_source_commit", source.get("packaging_commit", source.get("source_commit")))
    tree = source.get("source_tree_sha256", source.get("source_tree_digest"))
    artifact = source.get("artifact_sha256")
    if not isinstance(commit, str) or not SHA40.fullmatch(commit):
        raise EvidenceError(f"{label} frozen source commit is missing or malformed")
    if not isinstance(tree, str) or not SHA64.fullmatch(tree):
        raise EvidenceError(f"{label} source tree SHA-256 is missing or malformed")
    if not isinstance(artifact, str) or not SHA64.fullmatch(artifact):
        raise EvidenceError(f"{label} artifact SHA-256 is missing or malformed")
    return commit, tree, artifact


def _require_identity(document: Mapping[str, Any], expected: tuple[str, str, str], label: str) -> None:
    if _identity(document, label) != expected:
        raise EvidenceError(f"{label} identity does not match frozen candidate")


def _atomic_json(path: Path, value: Mapping[str, Any]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with tempfile.NamedTemporaryFile("w", encoding="utf-8", dir=path.parent, delete=False) as handle:
        json.dump(value, handle, ensure_ascii=False, indent=2, sort_keys=True)
        handle.write("\n")
        temporary = Path(handle.name)
    os.replace(temporary, path)


def _verify_lineage_binding(
    independence: Mapping[str, Any],
    lineage_policy: Path | None,
) -> tuple[dict[str, Any], dict[str, str]]:
    has_digest = "approved_lineage_policy_sha256" in independence
    has_source = "approved_lineage_source" in independence
    if not has_digest and not has_source:
        if lineage_policy is not None:
            raise EvidenceError("lineage policy was supplied without an approved lineage source")
        return {}, {}
    if not has_digest or not has_source:
        raise EvidenceError("approved lineage source requires a complete lineage policy binding")
    if lineage_policy is None:
        raise EvidenceError("approved lineage source requires the bound lineage policy file")

    policy = _read(lineage_policy, "lineage policy")
    policy_errors = validate_approved_lineage_policy(policy)
    if policy_errors:
        raise EvidenceError("lineage policy is not the closed approved contract: " + "; ".join(policy_errors))
    policy_sha = _sha(lineage_policy)
    if independence.get("approved_lineage_policy_sha256") != policy_sha:
        raise EvidenceError("lineage policy SHA-256 does not match the independence verdict")
    source_errors = validate_approved_lineage_source(independence.get("approved_lineage_source"), policy)
    if source_errors:
        raise EvidenceError("approved lineage source evidence is invalid: " + "; ".join(source_errors))
    source = independence["approved_lineage_source"]
    bindings = independence.get("reference_bindings")
    approved_bindings = [
        row for row in bindings or []
        if isinstance(row, Mapping) and row.get("identity") == policy["approved_reference_identity"]
    ]
    if len(approved_bindings) != 1 or approved_bindings[0].get("inventory_sha256") != source["inventory_sha256"]:
        raise EvidenceError("approved lineage source inventory does not match the comparator binding")
    compared_references = independence.get("references")
    if not isinstance(compared_references, Mapping) or policy["approved_reference_identity"] in compared_references:
        raise EvidenceError("approved v3.4 lineage must not be scored as an independence reference")
    return (
        {
            "approved_lineage_policy_sha256": policy_sha,
            "approved_lineage_source": deepcopy(independence["approved_lineage_source"]),
        },
        {"lineage_policy_sha256": policy_sha},
    )


def assemble(
    *,
    freeze_manifest: Path,
    artifact: Path,
    windows_inspection: Path,
    product_ui_inspection: Path,
    independence: Path,
    static_gates: Path,
    ai_slop: Path,
    code_review: Path,
    architecture: Path,
    compatibility: Path,
    lineage_policy: Path | None = None,
    bundle_root: Path | None = None,
) -> dict[str, Any]:
    freeze = _read(freeze_manifest, "freeze manifest")
    expected = _identity(freeze, "freeze manifest")
    commit, tree, artifact_sha = expected
    if not artifact.is_file() or _sha(artifact) != artifact_sha:
        raise EvidenceError("actual frozen artifact does not match freeze manifest")
    bundle_records, bundle_sha = _verify_bundle(freeze, artifact, bundle_root)

    windows = _read(windows_inspection, "Windows inspection")
    _require_identity(windows, expected, "Windows inspection")
    if windows.get("product_bundle_sha256") != bundle_sha or windows.get("product_files") != bundle_records:
        raise EvidenceError("Windows inspection product bundle identity does not match frozen set")
    if windows.get("verdict") != "PASS" or windows.get("suite") != "ProductRibbon" or windows.get("return_zip") != "PASS.zip" or windows.get("errors") != []:
        raise EvidenceError("Windows inspection is not a strict ProductRibbon PASS")
    if windows.get("packaging_commit") != commit:
        raise EvidenceError("Windows inspection packaging commit does not match frozen source")
    snapshot = windows.get("source_snapshot_sha256", windows.get("snapshot_sha256"))
    if not isinstance(snapshot, str) or not SHA64.fullmatch(snapshot):
        raise EvidenceError("Windows inspection snapshot SHA-256 is missing")
    evidence_files = windows.get("evidence_files")
    if not isinstance(evidence_files, list) or tuple(sorted(evidence_files)) != tuple(sorted(PRODUCT_EVIDENCE)):
        raise EvidenceError("Windows inspection does not name the exact ProductRibbon evidence set")
    evidence_hashes = windows.get("evidence_sha256")
    if not isinstance(evidence_hashes, Mapping) or set(evidence_hashes) != set(PRODUCT_EVIDENCE) or any(not isinstance(value, str) or not SHA64.fullmatch(value) for value in evidence_hashes.values()):
        raise EvidenceError("Windows inspection does not bind the exact ProductRibbon evidence hashes")
    if windows.get("office_version") not in {"Office 2024", "Microsoft 365"} or windows.get("compatibility_baseline") != "Office 2024" or windows.get("architecture") != "x64":
        raise EvidenceError("Windows inspection must prove Office 2024 compatibility on Office 2024 or Microsoft 365 x64")

    product_ui = _read(product_ui_inspection, "ProductUi inspection")
    _require_identity(product_ui, expected, "ProductUi inspection")
    if product_ui.get("product_bundle_sha256") != bundle_sha or product_ui.get("product_files") != bundle_records:
        raise EvidenceError("ProductUi inspection product bundle identity does not match frozen set")
    if product_ui.get("verdict") != "PASS" or product_ui.get("suite") != "ProductUi" or product_ui.get("return_zip") != "PASS.zip" or product_ui.get("errors") != []:
        raise EvidenceError("ProductUi inspection is not a strict PASS")
    if product_ui.get("packaging_commit") != commit:
        raise EvidenceError("ProductUi inspection packaging commit does not match frozen source")
    product_ui_snapshot = product_ui.get("source_snapshot_sha256", product_ui.get("snapshot_sha256"))
    if not isinstance(product_ui_snapshot, str) or not SHA64.fullmatch(product_ui_snapshot):
        raise EvidenceError("ProductUi inspection snapshot SHA-256 is missing")
    product_ui_files = product_ui.get("evidence_files")
    if not isinstance(product_ui_files, list) or tuple(product_ui_files) != PRODUCT_UI_EVIDENCE:
        raise EvidenceError("ProductUi inspection does not name the exact current-200 evidence set")
    product_ui_hashes = product_ui.get("evidence_sha256")
    if not isinstance(product_ui_hashes, Mapping) or set(product_ui_hashes) != set(PRODUCT_UI_EVIDENCE) or any(not isinstance(value, str) or not SHA64.fullmatch(value) for value in product_ui_hashes.values()):
        raise EvidenceError("ProductUi inspection does not bind the exact current-200 evidence hash")

    independence_doc = _read(independence, "independence verdict")
    _require_identity(independence_doc, expected, "independence verdict")
    if independence_doc.get("status") != "PASS":
        raise EvidenceError("independence verdict is not PASS")
    axes = independence_doc.get("axes")
    if not isinstance(axes, Mapping) or any(not isinstance(axes.get(axis), Mapping) or axes[axis].get("status") != "PASS" for axis in ("code", "workbook", "ui_wording", "assets", "provenance")):
        raise EvidenceError("all five independence axes must be PASS")
    lineage_fields, lineage_sources = _verify_lineage_binding(independence_doc, lineage_policy)

    static = _read(static_gates, "static gates")
    _require_identity(static, expected, "static gates")
    if any(static.get(name) != "PASS" for name in STATIC_GATES):
        raise EvidenceError("one or more static gates are not PASS")

    reviews = {
        "ai_slop": (_read(ai_slop, "ai-slop review"), "status", "PASS"),
        "code_review": (_read(code_review, "code review"), "verdict", "APPROVE"),
        "architecture": (_read(architecture, "architecture review"), "verdict", "CLEAR"),
    }
    for label, (document, field, accepted) in reviews.items():
        _require_identity(document, expected, label)
        if document.get(field) != accepted:
            raise EvidenceError(f"{label} is not {accepted}")

    compatibility_doc = _read(compatibility, "compatibility evidence")
    _require_identity(compatibility_doc, expected, "compatibility evidence")
    supported = {str(item) for item in compatibility_doc.get("supported_bitness", [])}
    native = {str(item) for item in compatibility_doc.get("native_validated_bitness", [])}
    if compatibility_doc.get("status") != "PASS" or compatibility_doc.get("baseline") != "Office 2024" or supported != {"32", "64"} or "64" not in native:
        raise EvidenceError("compatibility evidence must prove the Office 2024 baseline, 32/64 support, and native x64 validation")

    evidence_sources = {
        "freeze_manifest_sha256": _sha(freeze_manifest),
        "windows_inspection_sha256": _sha(windows_inspection),
        "product_ui_inspection_sha256": _sha(product_ui_inspection),
        "independence_sha256": _sha(independence),
        "static_gates_sha256": _sha(static_gates),
        "ai_slop_sha256": _sha(ai_slop),
        "code_review_sha256": _sha(code_review),
        "architecture_sha256": _sha(architecture),
        "compatibility_sha256": _sha(compatibility),
        **lineage_sources,
    }

    return {
        "schema_version": 1,
        **{name: static[name] for name in STATIC_GATES},
        "ai_slop": "PASS",
        "code_review": "APPROVE",
        "architecture": "CLEAR",
        "frozen_source_commit": commit,
        "source_tree_sha256": tree,
        "artifact_sha256": artifact_sha,
        **lineage_fields,
        "artifact_count": len(bundle_records),
        "product_bundle_sha256": bundle_sha,
        "product_files": bundle_records,
        "candidate_identity": {"frozen_source_commit": commit, "source_tree_sha256": tree, "artifact_sha256": artifact_sha, "product_bundle_sha256": bundle_sha},
        "compatibility": {
            "status": "PASS",
            "baseline": "Office 2024",
            "supported_bitness": ["32", "64"],
            "native_validated_bitness": sorted(native),
        },
        "windows": {
            "status": "PASS",
            "frozen_source_commit": commit,
            "source_tree_sha256": tree,
            "artifact_sha256": artifact_sha,
            "snapshot_sha256": snapshot,
            "office_version": windows.get("office_version"),
            "compatibility_baseline": windows.get("compatibility_baseline"),
            "architecture": windows.get("architecture"),
            "product_bundle_sha256": bundle_sha,
            "product_files": bundle_records,
            "compatibility": {
                "status": "PASS",
                "baseline": "Office 2024",
                "supported_bitness": ["32", "64"],
                "native_validated_bitness": sorted(native),
            },
            "evidence_files": list(PRODUCT_EVIDENCE),
            "evidence_sha256": dict(evidence_hashes),
        },
        "product_ui": {
            "status": "PASS",
            "frozen_source_commit": commit,
            "source_tree_sha256": tree,
            "artifact_sha256": artifact_sha,
            "snapshot_sha256": product_ui_snapshot,
            "product_bundle_sha256": bundle_sha,
            "product_files": bundle_records,
            "evidence_files": list(PRODUCT_UI_EVIDENCE),
            "evidence_sha256": dict(product_ui_hashes),
        },
        "independence": {axis: dict(axes[axis]) for axis in ("code", "workbook", "ui_wording", "assets", "provenance")},
        "evidence_sources": evidence_sources,
    }


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--freeze-manifest", required=True, type=Path)
    parser.add_argument("--artifact", required=True, type=Path)
    parser.add_argument("--windows-inspection", required=True, type=Path)
    parser.add_argument("--product-ui-inspection", required=True, type=Path)
    parser.add_argument("--independence", required=True, type=Path)
    parser.add_argument("--static-gates", required=True, type=Path)
    parser.add_argument("--ai-slop", required=True, type=Path)
    parser.add_argument("--code-review", required=True, type=Path)
    parser.add_argument("--architecture", required=True, type=Path)
    parser.add_argument("--compatibility", required=True, type=Path)
    parser.add_argument("--lineage-policy", type=Path)
    parser.add_argument("--output", required=True, type=Path)
    args = parser.parse_args(argv)
    try:
        result = assemble(**{key: getattr(args, key) for key in ("freeze_manifest", "artifact", "windows_inspection", "product_ui_inspection", "independence", "static_gates", "ai_slop", "code_review", "architecture", "compatibility", "lineage_policy")})
        _atomic_json(args.output, result)
    except EvidenceError as exc:
        print(f"FAIL release evidence: {exc}")
        return 1
    print(json.dumps(result, ensure_ascii=False, sort_keys=True))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
