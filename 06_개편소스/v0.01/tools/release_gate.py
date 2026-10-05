"""Fail-closed release evidence evaluator and promotion gate."""
from __future__ import annotations

import argparse
import hashlib
import json
import os
import re
import shutil
import tempfile
from pathlib import Path
from typing import Any, Mapping

try:
    from .build_release_evidence import EvidenceError, assemble
    from .compare_independence import validate_approved_lineage_source
except ImportError:  # direct script execution
    from build_release_evidence import EvidenceError, assemble
    from compare_independence import validate_approved_lineage_source

AXES = ("code", "workbook", "ui_wording", "assets", "provenance")
REQUIRED = ("spec_lock", "clean_build", "source_package", "tests", "feature_policy", "provenance", "detector", "ai_slop", "code_review", "architecture")
RIBBON_EVIDENCE = (
    "ProductRibbon.Package.json", "ProductRibbon.RibbonXml.json", "ProductRibbon.OfficeIdentity.json",
    "ProductRibbon.Ui.Current200.json",
    "ProductRibbon.Fast.json", "ProductRibbon.Guarded.json", "ProductRibbon.Planned.json",
    "ProductRibbon.UserForms.json",
)
PRODUCT_UI_EVIDENCE = ("ProductUi.Current200.json", "ProductUi.UserForms.json")
SHA40 = re.compile(r"^[0-9a-f]{40}$")
SHA64 = re.compile(r"^[0-9a-f]{64}$")
EVIDENCE_SOURCE_KEYS = (
    "freeze_manifest_sha256", "windows_inspection_sha256", "product_ui_inspection_sha256", "independence_sha256",
    "static_gates_sha256", "ai_slop_sha256", "code_review_sha256",
    "architecture_sha256", "compatibility_sha256",
)


def _ok(value: Any, accepted: set[str] = {"PASS"}) -> bool:
    return isinstance(value, str) and value in accepted


def _walk_names(value: Any) -> set[str]:
    found: set[str] = set()
    if isinstance(value, Mapping):
        for key, item in value.items():
            if isinstance(key, str) and key in RIBBON_EVIDENCE:
                found.add(key)
            found.update(_walk_names(item))
    elif isinstance(value, list):
        for item in value:
            if isinstance(item, str) and item in RIBBON_EVIDENCE:
                found.add(item)
            found.update(_walk_names(item))
    return found


def evaluate(evidence: Mapping[str, Any]) -> dict[str, Any]:
    """Return PASS only when every gate, identity, and Windows proof is closed."""
    errors: list[str] = []
    for field in REQUIRED:
        accepted = {"APPROVE"} if field == "code_review" else ({"CLEAR"} if field == "architecture" else {"PASS"})
        if not _ok(evidence.get(field), accepted):
            errors.append(f"{field} is not {next(iter(accepted))}")
    for field, pattern in (("frozen_source_commit", SHA40), ("source_tree_sha256", SHA64), ("artifact_sha256", SHA64)):
        if not isinstance(evidence.get(field), str) or not pattern.fullmatch(evidence[field]):
            errors.append(f"invalid {field}")
    axes = evidence.get("independence")
    if not isinstance(axes, Mapping):
        errors.append("independence axes missing")
    else:
        for axis in AXES:
            if not isinstance(axes.get(axis), Mapping) or axes[axis].get("status") != "PASS":
                errors.append(f"independence.{axis} is not PASS")
    windows = evidence.get("windows")
    if not isinstance(windows, Mapping) or windows.get("status") != "PASS":
        errors.append("windows evidence is not PASS")
    else:
        for field, pattern in (("frozen_source_commit", SHA40), ("source_tree_sha256", SHA64), ("artifact_sha256", SHA64), ("snapshot_sha256", SHA64)):
            if not isinstance(windows.get(field), str) or not pattern.fullmatch(windows[field]):
                errors.append(f"Windows snapshot missing/invalid {field}")
        for field, label in (("frozen_source_commit", "source commit"), ("artifact_sha256", "artifact digest"), ("source_tree_sha256", "source tree digest")):
            if windows.get(field) != evidence.get(field):
                errors.append(f"Windows/{label} mismatch")
        if "rows" in windows and windows.get("rows") != 52:
            errors.append("Windows matrix must contain 52 rows")
        office = str(windows.get("office_version", windows.get("office", ""))).casefold()
        compatibility_baseline = str(windows.get("compatibility_baseline", "")).casefold()
        architecture = str(windows.get("architecture", windows.get("arch", ""))).casefold()
        if office not in {"office 2024", "microsoft 365"} or compatibility_baseline != "office 2024":
            errors.append("Office 2024 compatibility baseline with Office 2024 or Microsoft 365 native evidence required")
        if architecture not in {"x64", "amd64"}:
            errors.append("Office 2024 or Microsoft 365 x64 evidence required")
        if windows.get("target_only") is True or windows.get("release_target") == "target_only" or windows.get("stale") is True:
            errors.append("stale/target_only evidence cannot release")
        compatibility = windows.get("compatibility", evidence.get("compatibility"))
        if not isinstance(compatibility, Mapping) or compatibility.get("status") != "PASS":
            errors.append("compatibility status is not PASS")
        else:
            if compatibility.get("baseline") != "Office 2024":
                errors.append("compatibility baseline must be Office 2024")
            supported = {str(v) for v in compatibility.get("supported_bitness", [])}
            native = {str(v) for v in compatibility.get("native_validated_bitness", [])}
            if not {"32", "64"}.issubset(supported):
                errors.append("compatibility must support 32 and 64 bitness")
            if "64" not in native:
                errors.append("compatibility native 64 bitness validation required")
        names = _walk_names(windows) | _walk_names(evidence.get("product_ribbon_evidence"))
        missing = [name for name in RIBBON_EVIDENCE if name not in names]
        if missing:
            errors.append("ProductRibbon evidence missing: " + ", ".join(missing))
        evidence_hashes = windows.get("evidence_sha256")
        if not isinstance(evidence_hashes, Mapping) or set(evidence_hashes) != set(RIBBON_EVIDENCE) or any(not isinstance(value, str) or not SHA64.fullmatch(value) for value in evidence_hashes.values()):
            errors.append("ProductRibbon evidence hashes missing or malformed")
    product_ui = evidence.get("product_ui")
    if not isinstance(product_ui, Mapping) or product_ui.get("status") != "PASS":
        errors.append("ProductUi evidence is not PASS")
    else:
        for field, pattern in (("frozen_source_commit", SHA40), ("source_tree_sha256", SHA64), ("artifact_sha256", SHA64), ("snapshot_sha256", SHA64)):
            if not isinstance(product_ui.get(field), str) or not pattern.fullmatch(product_ui[field]):
                errors.append(f"ProductUi snapshot missing/invalid {field}")
        for field, label in (("frozen_source_commit", "source commit"), ("artifact_sha256", "artifact digest"), ("source_tree_sha256", "source tree digest")):
            if product_ui.get(field) != evidence.get(field):
                errors.append(f"ProductUi/{label} mismatch")
        if product_ui.get("evidence_files") != list(PRODUCT_UI_EVIDENCE):
            errors.append("ProductUi current-200 evidence missing")
        product_ui_hashes = product_ui.get("evidence_sha256")
        if not isinstance(product_ui_hashes, Mapping) or set(product_ui_hashes) != set(PRODUCT_UI_EVIDENCE) or any(not isinstance(value, str) or not SHA64.fullmatch(value) for value in product_ui_hashes.values()):
            errors.append("ProductUi evidence hashes missing or malformed")
    records = evidence.get("product_files")
    if not isinstance(records, list) or not records:
        errors.append("complete release bundle records required")
    else:
        valid_records = all(isinstance(record, Mapping) and set(record) == {"path", "sha256"} and isinstance(record.get("path"), str) and isinstance(record.get("sha256"), str) and SHA64.fullmatch(record["sha256"]) for record in records)
        paths = [record["path"] for record in records if isinstance(record, Mapping) and isinstance(record.get("path"), str)]
        if not valid_records or len(paths) != len(records) or paths != sorted(paths, key=lambda value: value.encode("utf-8")) or len(set(paths)) != len(paths) or any("\\" in path or re.search(r'[:*?"<>|]', path) or Path(path).is_absolute() or any(part in {"", ".", ".."} for part in path.split("/")) for path in paths):
            errors.append("product_files are malformed, unsafe, duplicated, or unsorted")
        expected = [{"path": "Product.xlam", "sha256": evidence.get("artifact_sha256")}]
        if records != expected:
            errors.append("single Product XLAM file set required")
        if evidence.get("artifact_count") != len(records):
            errors.append("bundle artifact count mismatch")
        if valid_records and evidence.get("product_bundle_sha256") != _bundle_digest(records):
            errors.append("bundle digest mismatch")
        if isinstance(windows, Mapping) and (windows.get("product_files") != records or windows.get("product_bundle_sha256") != evidence.get("product_bundle_sha256")):
            errors.append("Windows product bundle identity mismatch")
        if isinstance(product_ui, Mapping) and (product_ui.get("product_files") != records or product_ui.get("product_bundle_sha256") != evidence.get("product_bundle_sha256")):
            errors.append("ProductUi product bundle identity mismatch")
    sources = evidence.get("evidence_sources")
    if not isinstance(sources, Mapping):
        errors.append("identity-bound evidence source hashes missing")
    else:
        for name in EVIDENCE_SOURCE_KEYS:
            if not isinstance(sources.get(name), str) or not SHA64.fullmatch(sources[name]):
                errors.append(f"invalid evidence source hash: {name}")
    lineage_digest_present = "approved_lineage_policy_sha256" in evidence
    lineage_approved_source_present = "approved_lineage_source" in evidence
    lineage_source_present = isinstance(sources, Mapping) and "lineage_policy_sha256" in sources
    if any((lineage_digest_present, lineage_approved_source_present, lineage_source_present)):
        if not all((lineage_digest_present, lineage_approved_source_present, lineage_source_present)):
            errors.append("approved lineage source binding is incomplete")
        else:
            policy_sha = evidence.get("approved_lineage_policy_sha256")
            if not isinstance(policy_sha, str) or not SHA64.fullmatch(policy_sha):
                errors.append("approved lineage policy SHA-256 is malformed")
            elif sources.get("lineage_policy_sha256") != policy_sha:
                errors.append("approved lineage policy/source SHA-256 mismatch")
            source_errors = validate_approved_lineage_source(evidence.get("approved_lineage_source"))
            errors.extend(source_errors)
    return {"status": "PASS" if not errors else "FAIL", "errors": errors}


def _load(path: Path) -> Any:
    return json.loads(path.read_text(encoding="utf-8"))


def _bundle_digest(records: list[Mapping[str, Any]]) -> str:
    payload = "".join(f"{row['path']}\n{row['sha256']}\n" for row in sorted(records, key=lambda row: str(row.get("path")).encode("utf-8"))).encode()
    return hashlib.sha256(payload).hexdigest()


def release(root: str | Path) -> dict[str, Any]:
    """Promote the SHA-bound frozen artifact only after all evidence passes."""
    root = Path(root).resolve()
    out, release_root = root / "out", root / "out" / "release"
    errors: list[str] = []
    try:
        evidence_path = out / "evidence" / "release-evidence.json"
        manifest_path = out / "frozen" / "current" / "freeze-manifest.json"
        evidence = _load(evidence_path)
        manifest = _load(manifest_path)
        artifact = out / "frozen" / "current" / str(manifest.get("artifact", ""))
        independence_path = out / "evidence" / "independence-verdict.json"
        lineage_policy = root / "provenance" / "approved-lineage-policy.json" if any(
            key in evidence for key in ("approved_lineage_policy_sha256", "approved_lineage_source")
        ) else None
        rebuilt = assemble(
            freeze_manifest=manifest_path,
            artifact=artifact,
            windows_inspection=out / "evidence" / "windows-inspection.json",
            product_ui_inspection=out / "evidence" / "product-ui-inspection.json",
            independence=independence_path,
            static_gates=out / "evidence" / "static-gates.json",
            ai_slop=out / "evidence" / "ai-slop.json",
            code_review=out / "evidence" / "code-review.json",
            architecture=out / "evidence" / "architecture.json",
            compatibility=out / "evidence" / "compatibility.json",
            lineage_policy=lineage_policy,
        )
        if rebuilt != evidence:
            errors.append("release evidence does not match identity-bound source evidence")
        result = evaluate(evidence)
        errors.extend(result["errors"])
        for field in ("frozen_source_commit", "source_tree_digest", "artifact_sha256", "artifact"):
            if not isinstance(manifest.get(field), str) or not manifest[field]:
                errors.append(f"freeze manifest missing {field}")
        if not artifact.is_file():
            errors.append("frozen artifact missing")
        else:
            digest = hashlib.sha256(artifact.read_bytes()).hexdigest()
            if digest != manifest.get("artifact_sha256") or digest != evidence.get("artifact_sha256"):
                errors.append("artifact SHA-256 mismatch")
        if manifest.get("frozen_source_commit") != evidence.get("frozen_source_commit") or manifest.get("source_tree_digest") != evidence.get("source_tree_sha256"):
            errors.append("freeze/evidence identity mismatch")
        records = manifest.get("product_files")
        evidence_records = evidence.get("product_files")
        if not isinstance(records, list) or records != evidence_records:
            errors.append("freeze/evidence product file set mismatch")
        elif manifest.get("product_bundle_sha256") != evidence.get("product_bundle_sha256") or manifest.get("product_bundle_sha256") != _bundle_digest(records):
            errors.append("freeze/evidence product bundle digest mismatch")
        else:
            for record in records:
                relative = Path(str(record.get("path", "")))
                member = out / "frozen" / "current" / relative
                if relative.is_absolute() or ".." in relative.parts or not member.is_file() or hashlib.sha256(member.read_bytes()).hexdigest() != record.get("sha256"):
                    errors.append(f"product file missing or tampered: {relative.as_posix()}")
        for name in RIBBON_EVIDENCE:
            path = out / "evidence" / name
            if not path.is_file():
                errors.append(f"missing {name}")
            else:
                item = _load(path)
                if item.get("status") != "PASS":
                    errors.append(f"{name} is not PASS")
                for field in ("artifact_sha256", "source_tree_sha256"):
                    if item.get(field) != evidence.get(field):
                        errors.append(f"{name} {field} mismatch")
                if item.get("source_snapshot_sha256") != evidence.get("windows", {}).get("snapshot_sha256"):
                    errors.append(f"{name} source snapshot mismatch")
                expected_evidence_sha = evidence.get("windows", {}).get("evidence_sha256", {}).get(name)
                if expected_evidence_sha != hashlib.sha256(path.read_bytes()).hexdigest():
                    errors.append(f"{name} evidence SHA-256 mismatch")
        for name in PRODUCT_UI_EVIDENCE:
            path = out / "evidence" / name
            if not path.is_file():
                errors.append(f"missing {name}")
            else:
                item = _load(path)
                if item.get("status") != "PASS":
                    errors.append(f"{name} is not PASS")
                for field in ("artifact_sha256", "source_tree_sha256"):
                    if item.get(field) != evidence.get(field):
                        errors.append(f"{name} {field} mismatch")
                if item.get("source_snapshot_sha256") != evidence.get("product_ui", {}).get("snapshot_sha256"):
                    errors.append(f"{name} source snapshot mismatch")
                expected_evidence_sha = evidence.get("product_ui", {}).get("evidence_sha256", {}).get(name)
                if expected_evidence_sha != hashlib.sha256(path.read_bytes()).hexdigest():
                    errors.append(f"{name} evidence SHA-256 mismatch")
    except (OSError, ValueError, json.JSONDecodeError, EvidenceError) as exc:
        errors.append(str(exc))
    verdict = {"status": "PASS" if not errors else "FAIL", "errors": errors}
    if not errors:
        staging = Path(tempfile.mkdtemp(prefix=".release-", dir=out))
        try:
            for record in manifest["product_files"]:
                relative = Path(record["path"])
                destination = staging / relative
                destination.parent.mkdir(parents=True, exist_ok=True)
                shutil.copy2(out / "frozen" / "current" / relative, destination)
            (staging / "release-verdict.json").write_text(json.dumps(verdict, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
            shutil.rmtree(release_root, ignore_errors=True)
            os.replace(staging, release_root)
        except Exception as exc:
            shutil.rmtree(staging, ignore_errors=True)
            errors.append(f"release bundle promotion failed: {exc}")
            verdict = {"status": "FAIL", "errors": errors}
    if errors:
        shutil.rmtree(release_root, ignore_errors=True)
        release_root.mkdir(parents=True, exist_ok=True)
        (release_root / "release-verdict.json").write_text(json.dumps(verdict, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    return verdict


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--evidence", type=Path)
    parser.add_argument("--root", type=Path)
    args = parser.parse_args(argv)
    if args.root:
        result = release(args.root)
    elif args.evidence:
        result = evaluate(_load(args.evidence))
    else:
        parser.error("one of --evidence or --root is required")
    print(json.dumps(result, ensure_ascii=False, indent=2))
    return 0 if result["status"] == "PASS" else 1


if __name__ == "__main__":
    raise SystemExit(main())
