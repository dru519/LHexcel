#!/usr/bin/env python3
"""Fail-closed inspection of a ProductRibbon Windows return directory.

This tool is deliberately read-only.  It accepts either ``return/`` itself or
an archived lane and reports ``PASS`` only when the one returned ZIP contains
the complete, SHA-bound ProductRibbon evidence set.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import math
import re
import os
import shutil
import sys
import tempfile
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path
from typing import Any

ROOT = Path(__file__).resolve().parents[1]
FEATURE_CONTRACT = json.loads(
    (ROOT / "contracts/feature-contract.json").read_text(encoding="utf-8")
)
EXPECTED_FEATURE_COUNT = len(FEATURE_CONTRACT["lineage"]["release_feature_ids"])
RIBBON_ROOT = ET.parse(ROOT / "src/ribbon/customUI14.xml").getroot()
EXPECTED_RIBBON_BUTTON_COUNT = sum(
    element.tag.rsplit("}", 1)[-1] == "button" for element in RIBBON_ROOT.iter()
)

EVIDENCE_JSON = (
    "ProductRibbon.Package.json",
    "ProductRibbon.RibbonXml.json",
    "ProductRibbon.OfficeIdentity.json",
    "ProductRibbon.Ui.Current200.json",
    "ProductRibbon.Fast.json",
    "ProductRibbon.Guarded.json",
    "ProductRibbon.Planned.json",
    "ProductRibbon.UserForms.json",
)
SUMMARY_JSON = "ProductRibbon.json"
PRODUCT_UI_EVIDENCE = ("ProductUi.Current200.json", "ProductUi.UserForms.json")
PRODUCT_UI_SUMMARY = "ProductUi.json"
EVIDENCE_LABELS = {
    "ProductRibbon.Package.json": "Package",
    "ProductRibbon.RibbonXml.json": "RibbonXml",
    "ProductRibbon.OfficeIdentity.json": "OfficeIdentity",
    "ProductRibbon.Ui.Current200.json": "Ui.Current200",
    "ProductRibbon.Fast.json": "Fast",
    "ProductRibbon.Guarded.json": "Guarded",
    "ProductRibbon.Planned.json": "Planned",
    "ProductRibbon.UserForms.json": "UserForms",
}
REQUIRED = EVIDENCE_JSON + (
    SUMMARY_JSON,
    "Product.xlam",
    "manifest.json",
    "SHA256SUMS",
)
SHA64 = re.compile(r"^[0-9a-f]{64}$")
SHA40 = re.compile(r"^[0-9a-f]{40}$")
UUID_RE = re.compile(r"^[0-9a-f]{8}-[0-9a-f]{4}-[1-5][0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$", re.I)
PRODUCT_RIBBON_CALLBACKS = (
    "NxRibbonExecute",
    "NxRibbonGetAllFunctionsContent",
    "NxRibbonGetFavoritesContent",
    "NxRibbonGetFocusContent",
    "NxRibbonGetManagementContent",
    "NxRibbonOnLoad",
)
COMMON_FIELDS = {"schema_version", "suite", "status", "evidence", "method", "run_id", "artifact_sha256", "source_tree_sha256", "source_snapshot_sha256", "measured"}
MEASURED_FIELDS = {
    "ProductRibbon.Package.json": {"method", "custom_ui_parts", "relationship_type"},
    "ProductRibbon.RibbonXml.json": {
        "method", "callbacks", "compile_control_id", "compile_saved",
        "active_project_artifact_name", "reference_count",
        "broken_reference_count", "group_count", "button_count", "feature_count",
        "embedded_xml_sha256", "source_xml_sha256", "macro_execution",
        "grade_policy_contract_attested", "grade_policy_pattern_count",
    },
    "ProductRibbon.OfficeIdentity.json": {"method", "office_version", "compatibility_baseline", "architecture", "excel_version", "excel_build", "product_release_ids", "macro_bitness"},
    "ProductRibbon.Ui.Current200.json": {"method", "windows_scale_percent", "host_dpi", "scale_changed_by_test", "foreground_verified", "keytip_accessible", "keytips", "excel_window_bounds", "excel_client_bounds", "home_controls", "product_groups"},
    "ProductRibbon.Fast.json": {
        "method", "execution_grade", "resolved", "route_feature_id",
        "route_verified", "form_opened", "route_proof",
        "policy_attested_from_compiled_product",
    },
    "ProductRibbon.Guarded.json": {
        "method", "execution_grade", "resolved", "route_feature_id",
        "route_verified", "form_opened", "route_proof",
        "policy_attested_from_compiled_product",
    },
    "ProductRibbon.Planned.json": {
        "method", "execution_grade", "resolved", "route_feature_id",
        "route_verified", "form_opened", "route_proof",
        "policy_attested_from_compiled_product",
    },
    "ProductRibbon.UserForms.json": {"method", "form_bindings", "forms", "total_forms", "total_controls", "textboxes", "listboxes"},
}
DIAGNOSTIC_REQUIRED = {
    "manifest.json",
    "diagnostic.json",
    "SHA256SUMS",
    "RUN_PRODUCT_RIBBON_VERIFY.bat",
    "검증_안내.md",
}
DIAGNOSTIC_EXIT_CODES = {10, 20, 21, 22, 23, 24, 25, 26}
WINDOWS_ABSOLUTE_PATH = re.compile(
    r'(?i)(?<![A-Za-z0-9_.-])(?:[A-Z]:\\|\\\\[^\\/:*?"<>|\r\n]+\\[^\\/:*?"<>|\r\n]+)'
)
UNREDACTED_BEARER = re.compile(r"(?i)\bBearer\s+(?!\[REDACTED\])\S+")
UNREDACTED_SECRET = re.compile(
    r"(?i)\b(?:password|secret|token|apikey|key)\s*[=:]\s*(?!\[REDACTED\])\S+"
)


def _sha(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def _zip_member_sha(path: Path, name: str) -> str:
    with zipfile.ZipFile(path) as archive:
        return _sha(archive.read(name))


def _json(data: bytes, name: str) -> dict[str, Any]:
    try:
        value = json.loads(data.decode("utf-8-sig"))
    except (UnicodeDecodeError, json.JSONDecodeError) as exc:
        raise ValueError(f"{name} is not valid JSON") from exc
    if not isinstance(value, dict):
        raise ValueError(f"{name} must contain an object")
    return value


def _find(doc: Any, *keys: str) -> list[Any]:
    """Collect values recursively, allowing evidence producers to nest fields."""
    found: list[Any] = []
    if isinstance(doc, dict):
        for key, value in doc.items():
            if key in keys:
                found.append(value)
            found.extend(_find(value, *keys))
    elif isinstance(doc, list):
        for value in doc:
            found.extend(_find(value, *keys))
    return found


def _one(doc: dict[str, Any], keys: tuple[str, ...]) -> Any:
    values = [v for v in _find(doc, *keys) if v is not None]
    return values[0] if values else None


def _exact(doc: dict[str, Any], allowed: set[str], name: str, errors: list[str]) -> None:
    if set(doc) != allowed:
        errors.append(f"{name} schema keys must be exact")


def _is_int(value: Any) -> bool:
    return isinstance(value, int) and not isinstance(value, bool)


def _finite_positive_number(value: Any) -> bool:
    try:
        return type(value) in (int, float) and math.isfinite(value) and value > 0
    except OverflowError:
        return False


def _scalar_strings(value: Any):
    if isinstance(value, str):
        yield value
    elif isinstance(value, dict):
        for item in value.values():
            yield from _scalar_strings(item)
    elif isinstance(value, list):
        for item in value:
            yield from _scalar_strings(item)


def _unredacted(value: Any) -> bool:
    return any(
        WINDOWS_ABSOLUTE_PATH.search(text)
        or UNREDACTED_BEARER.search(text)
        or UNREDACTED_SECRET.search(text)
        for text in _scalar_strings(value)
    )


def _inspect_product_diagnostic(
    archive: zipfile.ZipFile,
    members: list[str],
    archive_path: Path,
    result: dict[str, Any],
    errors: list[str],
) -> None:
    by_path = set(members)
    missing = sorted(DIAGNOSTIC_REQUIRED - by_path)
    if missing:
        errors.append("ProductRibbon DIAGNOSTIC missing control files: " + ", ".join(missing))
    extras = sorted(
        name
        for name in by_path - DIAGNOSTIC_REQUIRED
        if not (
            re.fullmatch(r"results/(?:build/)?[^/]+\.json", name)
            or re.fullmatch(r"logs/[^/]+\.redacted\.txt", name)
        )
    )
    if extras:
        errors.append("ProductRibbon DIAGNOSTIC contains unsupported members: " + ", ".join(extras))
    if missing:
        return

    manifest = _json(archive.read("manifest.json"), "manifest.json")
    diagnostic = _json(archive.read("diagnostic.json"), "diagnostic.json")
    run_id = manifest.get("run_id")
    if manifest.get("schema_version") != 1 or manifest.get("suite") != "ProductRibbon":
        errors.append("ProductRibbon DIAGNOSTIC manifest identity is invalid")
    if not isinstance(run_id, str) or not UUID_RE.fullmatch(run_id):
        errors.append("ProductRibbon DIAGNOSTIC manifest run_id missing or malformed")

    _exact(
        diagnostic,
        {"schema_version", "run_id", "exit_code", "phase", "failure"},
        "diagnostic.json",
        errors,
    )
    diagnostic_run_id = diagnostic.get("run_id")
    exit_code = diagnostic.get("exit_code")
    phase = diagnostic.get("phase")
    if diagnostic.get("schema_version") != 1:
        errors.append("diagnostic.json schema_version must be exactly 1")
    if diagnostic_run_id != run_id:
        errors.append("manifest and diagnostic run_id disagree")
    if not _is_int(exit_code) or exit_code not in DIAGNOSTIC_EXIT_CODES:
        errors.append("diagnostic exit_code must be a supported nonzero code")
    if not isinstance(phase, str) or not re.fullmatch(r"[A-Za-z][A-Za-z0-9_-]*", phase):
        errors.append("diagnostic phase is missing or malformed")
    if not isinstance(diagnostic.get("failure"), str) or not diagnostic["failure"].strip():
        errors.append("diagnostic failure must be a meaningful string")
    if _unredacted(diagnostic):
        errors.append("diagnostic.json contains an unredacted path or secret scalar")

    product_result_name = "results/ProductRibbon.json"
    product_result = None
    if product_result_name in by_path:
        product_result = _json(archive.read(product_result_name), product_result_name)
        if product_result.get("run_id") != run_id:
            errors.append("manifest and ProductRibbon diagnostic result run_id disagree")
        has_exit = "exit_code" in product_result
        has_phase = "failure_phase" in product_result
        if has_exit != has_phase:
            errors.append("ProductRibbon diagnostic result must carry exit_code and failure_phase together")
        if has_exit and product_result.get("exit_code") != exit_code:
            errors.append("diagnostic and ProductRibbon result exit_code disagree")
        if has_phase and product_result.get("failure_phase") != phase:
            errors.append("diagnostic and ProductRibbon result phase disagree")
        if _unredacted(product_result):
            errors.append("ProductRibbon diagnostic result contains an unredacted path or secret scalar")

    for name in sorted(n for n in by_path if n.startswith("logs/")):
        try:
            log_text = archive.read(name).decode("utf-8-sig")
        except UnicodeDecodeError:
            errors.append(f"{name} is not valid UTF-8")
            continue
        if _unredacted(log_text):
            errors.append(f"{name} contains unredacted path or secret text")

    result.update(
        {
            "run_id": run_id,
            "exit_code": exit_code,
            "phase": phase,
            "return_zip_sha256": _sha(archive_path.read_bytes()),
            "diagnostic_sha256": _sha(archive.read("diagnostic.json")),
            "result_sha256": (
                _sha(archive.read(product_result_name)) if product_result is not None else None
            ),
        }
    )


def _sum_records(text: str) -> dict[str, str]:
    records: dict[str, str] = {}
    previous = ""
    for line in text.splitlines():
        if not line:
            continue
        match = re.fullmatch(r"([0-9a-f]{64})  ([^\\/].*)", line)
        if not match:
            raise ValueError("malformed SHA256SUMS line")
        digest, name = match.groups()
        if name in records or (previous and name.encode("utf-8") <= previous.encode("utf-8")):
            raise ValueError("duplicate or unsorted SHA256SUMS entry")
        parts = name.split("/")
        if "\\" in name or re.search(r'[:*?"<>|]', name) or name in {"manifest.json", "SHA256SUMS"} or any(part in {"", ".", ".."} for part in parts):
            raise ValueError("SHA256SUMS cannot self-reference control files")
        records[name] = digest
        previous = name
    return records


def _valid_bounds(value: Any) -> bool:
    return (
        isinstance(value, dict)
        and set(value) == {"left", "top", "width", "height"}
        and all(isinstance(value.get(key), (int, float)) and not isinstance(value.get(key), bool) for key in value)
        and value["width"] > 0
        and value["height"] > 0
    )


def _inspect_product_ui_return(input_dir: str | Path) -> dict[str, Any]:
    root = Path(input_dir)
    result: dict[str, Any] = {"suite": "ProductUi", "verdict": "DIAGNOSTIC", "errors": []}
    errors: list[str] = result["errors"]
    if not root.is_dir():
        errors.append("return directory is missing")
        return result
    candidates = [path for path in (root / "PASS.zip", root / "DIAGNOSTIC.zip") if path.is_file()]
    if len(candidates) != 1:
        errors.append("exactly one PASS.zip or DIAGNOSTIC.zip is required")
        return result
    archive_path = candidates[0]
    result["return_zip"] = archive_path.name
    try:
        with zipfile.ZipFile(archive_path) as archive:
            members = archive.namelist()
            if len(members) != len(set(members)) or any(
                "\\" in name
                or name.startswith("/")
                or any(part in {"", ".", ".."} for part in name.split("/"))
                for name in members
            ):
                errors.append("return ZIP member names are unsafe or duplicated")
                return result
            if archive_path.name == "DIAGNOSTIC.zip":
                required = {"manifest.json", "diagnostic.json", "SHA256SUMS", "RUN_PRODUCT_UI_VERIFY.bat", "검증_안내.md"}
                missing = sorted(required - set(members))
                extras = sorted(
                    name for name in set(members) - required
                    if not re.fullmatch(r"(?:results/(?:build/)?[^/]+\.json|logs/[^/]+\.redacted\.txt)", name)
                )
                if missing:
                    errors.append("ProductUi DIAGNOSTIC missing control files: " + ", ".join(missing))
                if extras:
                    errors.append("ProductUi DIAGNOSTIC contains unsupported members: " + ", ".join(extras))
                if not missing:
                    manifest = _json(archive.read("manifest.json"), "manifest.json")
                    diagnostic = _json(archive.read("diagnostic.json"), "diagnostic.json")
                    run_id = manifest.get("run_id")
                    if manifest.get("schema_version") != 1 or manifest.get("suite") != "ProductUi":
                        errors.append("ProductUi DIAGNOSTIC manifest identity is invalid")
                    if not isinstance(run_id, str) or not UUID_RE.fullmatch(run_id):
                        errors.append("ProductUi DIAGNOSTIC run_id is invalid")
                    if diagnostic.get("schema_version") != 1 or diagnostic.get("run_id") != run_id:
                        errors.append("ProductUi diagnostic identity is invalid")
                    if not _is_int(diagnostic.get("exit_code")) or diagnostic.get("exit_code") not in DIAGNOSTIC_EXIT_CODES:
                        errors.append("ProductUi diagnostic exit code is invalid")
                    if _unredacted(diagnostic):
                        errors.append("ProductUi diagnostic contains unredacted content")
                    result.update({"run_id": run_id, "exit_code": diagnostic.get("exit_code"), "phase": diagnostic.get("phase"), "diagnostic_valid": not errors})
                return result

            minimal_required = {
                PRODUCT_UI_SUMMARY, *PRODUCT_UI_EVIDENCE, "Product.xlam",
                "manifest.json", "SHA256SUMS",
            }
            composite_required = minimal_required | set(EVIDENCE_JSON) | {SUMMARY_JSON}
            required = set(members)
            if required not in (minimal_required, composite_required):
                errors.append("ProductUi PASS member set must be exact")
                return result
            sums = _sum_records(archive.read("SHA256SUMS").decode("utf-8-sig"))
            expected_sum_names = required - {"manifest.json", "SHA256SUMS"}
            if set(sums) != expected_sum_names:
                errors.append("ProductUi SHA256SUMS member set is invalid")
            for name, digest in sums.items():
                if _sha(archive.read(name)) != digest:
                    errors.append(f"ProductUi checksum mismatch: {name}")

            manifest = _json(archive.read("manifest.json"), "manifest.json")
            summary = _json(archive.read("ProductUi.json"), "ProductUi.json")
            ui_bytes = archive.read("ProductUi.Current200.json")
            ui = _json(ui_bytes, "ProductUi.Current200.json")
            attestation_bytes = archive.read("ProductUi.UserForms.json")
            attestation = _json(attestation_bytes, "ProductUi.UserForms.json")
            attestation_sha = _sha(attestation_bytes)
            manifest_fields = {"schema_version", "suite", "run_id", "packaging_commit", "artifact_sha256", "source_tree_sha256", "source_snapshot_sha256", "product_bundle_sha256", "product_files"}
            _exact(manifest, manifest_fields, "ProductUi manifest", errors)
            run_id = manifest.get("run_id")
            artifact = manifest.get("artifact_sha256")
            source_tree = manifest.get("source_tree_sha256")
            source_snapshot = manifest.get("source_snapshot_sha256")
            if manifest.get("schema_version") != 1 or manifest.get("suite") != "ProductUi":
                errors.append("ProductUi manifest identity is invalid")
            if not isinstance(manifest.get("packaging_commit"), str) or not SHA40.fullmatch(manifest["packaging_commit"]):
                errors.append("ProductUi packaging commit is invalid")
            if not isinstance(run_id, str) or not UUID_RE.fullmatch(run_id):
                errors.append("ProductUi run_id is invalid")
            for label, value in (("artifact", artifact), ("source tree", source_tree), ("source snapshot", source_snapshot)):
                if not isinstance(value, str) or not SHA64.fullmatch(value):
                    errors.append(f"ProductUi {label} digest is invalid")
            artifact_bytes = archive.read("Product.xlam")
            if isinstance(artifact, str) and _sha(artifact_bytes) != artifact:
                errors.append("ProductUi artifact digest mismatch")
            product_files = manifest.get("product_files")
            if product_files != [{"path": "Product.xlam", "sha256": artifact}]:
                errors.append("ProductUi product file inventory is invalid")
            elif manifest.get("product_bundle_sha256") != _sha(f"Product.xlam\n{artifact}\n".encode()):
                errors.append("ProductUi product bundle digest is invalid")

            summary_fields = {"schema_version", "suite", "status", "mode", "run_id", "source_digest", "snapshot_digest", "artifact_sha256", "environment", "cleanup", "run"}
            _exact(summary, summary_fields, "ProductUi.json", errors)
            cleanup = summary.get("cleanup")
            cleanup_valid = (
                isinstance(cleanup, dict)
                and set(cleanup) == {"status", "exit_mode", "failures"}
                and cleanup.get("status") == "PASS"
                and isinstance(cleanup.get("exit_mode"), str)
                and bool(cleanup["exit_mode"])
                and cleanup.get("failures") == []
            )
            if (
                summary.get("schema_version") != 1
                or summary.get("suite") != "ProductUi"
                or summary.get("status") != "PASS"
                or summary.get("mode") != "Green"
                or summary.get("run_id") != run_id
                or summary.get("source_digest") != source_tree
                or summary.get("snapshot_digest") != source_snapshot
                or summary.get("artifact_sha256") != artifact
                or summary.get("environment") != {"excel_bitness": "64-bit", "windows_scale_percent": 200}
                or not cleanup_valid
                or summary.get("run") != {"total": 1, "passed": 1, "failed": 0, "skipped": 0}
            ):
                errors.append("ProductUi summary identity or run evidence is invalid")

            if required == composite_required:
                prerequisite_summary = _json(archive.read(SUMMARY_JSON), SUMMARY_JSON)
                if (
                    prerequisite_summary.get("schema_version") != 1
                    or prerequisite_summary.get("suite") != "ProductRibbon"
                    or prerequisite_summary.get("status") != "PASS"
                    or prerequisite_summary.get("mode") != "Green"
                    or prerequisite_summary.get("run_id") != run_id
                    or prerequisite_summary.get("source_digest") != source_tree
                    or prerequisite_summary.get("snapshot_digest") != source_snapshot
                    or prerequisite_summary.get("artifact_sha256") != artifact
                    or prerequisite_summary.get("run") != {"total": 3, "passed": 3, "failed": 0, "skipped": 0}
                ):
                    errors.append("ProductRibbon prerequisite summary binding is invalid")
                for name in EVIDENCE_JSON:
                    prerequisite = _json(archive.read(name), name)
                    if (
                        prerequisite.get("schema_version") != 1
                        or prerequisite.get("suite") != "ProductRibbon"
                        or prerequisite.get("status") != "PASS"
                        or prerequisite.get("run_id") != run_id
                        or prerequisite.get("artifact_sha256") != artifact
                        or prerequisite.get("source_tree_sha256") != source_tree
                        or prerequisite.get("source_snapshot_sha256") != source_snapshot
                    ):
                        errors.append(f"ProductRibbon prerequisite evidence binding is invalid: {name}")

            ui_fields = {
                "schema_version", "suite", "status", "run_id", "artifact_sha256",
                "source_tree_sha256", "source_snapshot_sha256", "windows_scale_percent",
                "scale_changed_by_test", "excel_client_bounds", "menu_cases", "direct_case",
                "form_cases", "forbidden_text_matches",
            }
            _exact(ui, ui_fields, "ProductUi.Current200.json", errors)
            if (
                ui.get("schema_version") != 1
                or ui.get("suite") != "ProductUi"
                or ui.get("status") != "PASS"
                or ui.get("run_id") != run_id
                or ui.get("artifact_sha256") != artifact
                or ui.get("source_tree_sha256") != source_tree
                or ui.get("source_snapshot_sha256") != source_snapshot
                or ui.get("windows_scale_percent") != 200
                or ui.get("scale_changed_by_test") is not False
                or ui.get("forbidden_text_matches") != []
                or not _valid_bounds(ui.get("excel_client_bounds"))
            ):
                errors.append("ProductUi current 200 percent identity is invalid")

            _exact(attestation, COMMON_FIELDS, "ProductUi.UserForms.json", errors)
            attestation_measured = attestation.get("measured")
            if isinstance(attestation_measured, dict):
                _exact(
                    attestation_measured,
                    MEASURED_FIELDS["ProductRibbon.UserForms.json"],
                    "ProductUi.UserForms.json measured",
                    errors,
                )
            else:
                errors.append("ProductRibbon UserForms measured evidence is invalid")
                attestation_measured = {}
            if (
                attestation.get("schema_version") != 1
                or attestation.get("suite") != "ProductRibbon"
                or attestation.get("status") != "PASS"
                or attestation.get("evidence") != "UserForms"
                or attestation.get("method") != "built Product.xlam close/reopen + VBProject Designer.Controls readback"
                or attestation.get("run_id") != run_id
                or attestation.get("artifact_sha256") != artifact
                or attestation.get("source_tree_sha256") != source_tree
                or attestation.get("source_snapshot_sha256") != source_snapshot
                or attestation_measured.get("method") != attestation.get("method")
            ):
                errors.append("ProductRibbon UserForms attestation binding is invalid")
            source_root = Path(__file__).resolve().parents[1]
            contract_path = source_root / "contracts" / "ui-surface-contract.json"
            contract = json.loads(contract_path.read_text(encoding="utf-8"))
            product_manifest_path = source_root / "build" / "manifests" / "Product.json"
            product_manifest = json.loads(product_manifest_path.read_text(encoding="utf-8"))
            product_form_bindings = product_manifest.get("form_bindings")
            expected_attested_forms = {
                Path(binding["layout"]).name.removesuffix(".form.json")
                for binding in product_form_bindings
                if isinstance(binding, dict) and isinstance(binding.get("layout"), str)
            } if isinstance(product_form_bindings, list) else set()
            if (
                not isinstance(product_form_bindings, list)
                or len(expected_attested_forms) != len(product_form_bindings)
                or not expected_attested_forms
            ):
                errors.append("Product manifest UserForms inventory is invalid")
            attested_form_rows = attestation_measured.get("forms")
            if not isinstance(attested_form_rows, list) or len(attested_form_rows) != len(expected_attested_forms):
                errors.append("ProductRibbon UserForms attestation inventory is invalid")
                attested_form_rows = []
            attested_forms: dict[str, dict[str, Any]] = {}
            for attested_form in attested_form_rows:
                if (
                    not isinstance(attested_form, dict)
                    or set(attested_form) != {
                        "name", "layout", "code", "layout_sha256", "code_sha256",
                        "default_control", "cancel_control", "controls",
                    }
                ):
                    errors.append("ProductRibbon UserForms form record is invalid")
                    continue
                attested_name = attested_form.get("name")
                controls = attested_form.get("controls")
                if (
                    not isinstance(attested_name, str)
                    or not attested_name
                    or attested_name in attested_forms
                    or not isinstance(controls, list)
                ):
                    errors.append("ProductRibbon UserForms form identity is invalid")
                    continue
                attested_forms[attested_name] = attested_form
            if set(attested_forms) != expected_attested_forms:
                errors.append("ProductRibbon UserForms active form inventory is invalid")

            catalog_path = Path(__file__).resolve().parents[1] / "src" / "vba" / "ui" / "NxRibbonMenuCatalog.bas"
            catalog_bytes = catalog_path.read_bytes()
            catalog_source = catalog_bytes.decode("utf-8-sig")
            management_labels = list(dict.fromkeys(
                re.findall(r'NxManagementButton\("[^"]+", "([^"]+)"', catalog_source)
            ))
            navigation_path = Path(__file__).resolve().parents[1] / "src" / "resources" / "navigation.ko-KR.json"
            navigation_bytes = navigation_path.read_bytes()
            navigation = _json(navigation_bytes, "navigation.ko-KR.json")
            category_labels = [row.get("label") for row in navigation.get("categories", []) if isinstance(row, dict)]
            favorite_empty_match = re.search(r'id=""fav_empty"" label=""([^"]+)""', catalog_source)
            favorite_editor_match = re.search(r'id=""fav_editor"" label=""([^"]+)""', catalog_source)
            favorite_add_match = re.search(r'id=""fav_add"" label=""([^"]+)""', catalog_source)
            favorite_remove_match = re.search(r'id=""fav_remove"" label=""([^"]+)""', catalog_source)
            feature_contract_path = Path(__file__).resolve().parents[1] / "contracts" / "feature-contract.json"
            feature_contract = _json(feature_contract_path.read_bytes(), "feature-contract.json")
            allowed_favorite_labels = {
                row.get("ribbon", {}).get("label")
                for row in feature_contract.get("features", [])
                if isinstance(row, dict) and row.get("owner") != "ai" and isinstance(row.get("ribbon"), dict)
            }
            allowed_favorite_labels.discard(None)
            if len(management_labels) != 6 or len(category_labels) != 13 or not allowed_favorite_labels or favorite_editor_match is None or favorite_empty_match is None or favorite_add_match is None or favorite_remove_match is None:
                errors.append("ProductUi local dynamic menu catalog is invalid")
            favorite_editor = "" if favorite_editor_match is None else favorite_editor_match.group(1)
            favorite_empty = "" if favorite_empty_match is None else favorite_empty_match.group(1)
            favorite_add = "" if favorite_add_match is None else favorite_add_match.group(1)
            favorite_remove = "" if favorite_remove_match is None else favorite_remove_match.group(1)
            catalog_sha = _sha(catalog_bytes)
            navigation_sha = _sha(navigation_bytes)
            ribbon_path = source_root / "src" / "ribbon" / "customUI14.xml"
            ribbon_root = ET.fromstring(ribbon_path.read_bytes())
            ribbon_labels = {
                element.attrib.get("id"): element.attrib.get("label")
                for element in ribbon_root.iter()
                if element.attrib.get("id")
            }
            menu_control_ids = ("NX-PROD-MANAGEMENT", "NX-PROD-FAVORITES", "NX-PROD-ALL")
            if any(not ribbon_labels.get(control_id) for control_id in menu_control_ids):
                errors.append("ProductUi local Ribbon menu identity is invalid")
            expected_menu_details = [
                ("Product", label, labels, prefixes, source_sha)
                for label, labels, prefixes, source_sha in (
                    (ribbon_labels.get("NX-PROD-MANAGEMENT"), management_labels, ["nx1|entry|"], catalog_sha),
                    (ribbon_labels.get("NX-PROD-FAVORITES"), None, None, catalog_sha),
                    (ribbon_labels.get("NX-PROD-ALL"), category_labels, ["nx1|feature|", "nx1|command|"], navigation_sha),
                )
            ]
            menu_cases = ui.get("menu_cases")
            expected_menu_pairs = [(surface, label) for surface, label, _, _, _ in expected_menu_details]
            actual_menu_pairs: list[tuple[Any, Any]] = []
            if not isinstance(menu_cases, list) or len(menu_cases) != 3:
                errors.append("ProductUi menu case count must be exactly 3")
                menu_cases = []
            for index, case in enumerate(menu_cases):
                fields = {"surface", "menu_label", "bounds", "child_count", "children", "typed_route_prefixes", "typed_routes_verified", "typed_route_source_sha256"}
                if not isinstance(case, dict) or set(case) != fields:
                    errors.append("ProductUi menu case schema is invalid")
                    continue
                actual_menu_pairs.append((case.get("surface"), case.get("menu_label")))
                children = case.get("children")
                prefixes = case.get("typed_route_prefixes")
                expected_children = expected_menu_details[index][2] if index < len(expected_menu_details) else []
                expected_prefixes = expected_menu_details[index][3] if index < len(expected_menu_details) else []
                expected_source_sha = expected_menu_details[index][4] if index < len(expected_menu_details) else ""
                child_labels = [child.get("label") for child in children if isinstance(child, dict)] if isinstance(children, list) else []
                if index < len(expected_menu_details) and expected_menu_details[index][1] == "즐겨찾기":
                    empty_favorites = child_labels == [favorite_editor, favorite_empty, favorite_add]
                    populated_favorite_labels = child_labels[1:-2] if len(child_labels) >= 4 else []
                    populated_favorites = (
                        len(child_labels) >= 4
                        and child_labels[0] == favorite_editor
                        and child_labels[-2:] == [favorite_add, favorite_remove]
                        and len(populated_favorite_labels) == len(set(populated_favorite_labels))
                        and all(label in allowed_favorite_labels for label in populated_favorite_labels)
                    )
                    expected_children_valid = empty_favorites or populated_favorites
                    expected_prefixes = ["nx1|entry|", "fav2|add|"] if empty_favorites else ["nx1|entry|", "fav2|add|", "fav2|remove|"]
                else:
                    expected_children_valid = child_labels == expected_children
                if (
                    not _valid_bounds(case.get("bounds"))
                    or not isinstance(children, list)
                    or len(children) < 2
                    or case.get("child_count") != len(children)
                    or not expected_children_valid
                    or any(
                        not isinstance(child, dict)
                        or set(child) != {"label", "control_type", "left", "top", "width", "height"}
                        or not _valid_bounds({key: child.get(key) for key in ("left", "top", "width", "height")})
                        or not isinstance(child.get("label"), str)
                        or not child.get("label")
                        or not isinstance(child.get("control_type"), str)
                        or not child.get("control_type")
                        for child in children
                    )
                    or prefixes != expected_prefixes
                    or case.get("typed_routes_verified") is not True
                    or case.get("typed_route_source_sha256") != expected_source_sha
                ):
                    errors.append("ProductUi menu case evidence is invalid")
            if actual_menu_pairs != expected_menu_pairs:
                errors.append("ProductUi menu case order or identity is invalid")

            direct = ui.get("direct_case")
            if (
                not isinstance(direct, dict)
                or set(direct) != {"feature_id", "popup_count", "workbook_before_sha256", "workbook_after_sha256"}
                or direct.get("feature_id") != "NX-DATA-COPY-VISIBLE"
                or direct.get("popup_count") != 0
                or direct.get("workbook_before_sha256") != direct.get("workbook_after_sha256")
                or not isinstance(direct.get("workbook_before_sha256"), str)
                or not SHA64.fullmatch(direct["workbook_before_sha256"])
            ):
                errors.append("ProductUi direct no-popup evidence is invalid")

            expected_forms = contract["native_200_sample_forms"]
            expected_features = {
                "FNxAi": "",
                "FNxDataAnalyze": "NX-DATA-UNIQUE-COUNT",
                "FNxDataNormalize": "NX-DATA-NORMALIZE",
                "FNxFocusSettings": "",
                "FNxDrawTable": "NX-DRAW-BUSINESS-TABLE",
                "FNxRoleStyle": "NX-DRAW-ROLE-STYLE",
                "FNxFolderCreate": "NX-FILE-FOLDER-CREATE",
                "FNxFileConsolidate": "NX-FILE-CONSOLIDATE",
                "FNxCalculator": "NX-UTIL-CALCULATOR",
                "FNxSymbols": "NX-UTIL-SYMBOLS",
            }
            active_surfaces = contract["active_form_surfaces"]
            classes = contract["classes"]
            client = ui.get("excel_client_bounds") if _valid_bounds(ui.get("excel_client_bounds")) else {"width": 0, "height": 0}
            form_cases = ui.get("form_cases")
            if not isinstance(form_cases, list) or len(form_cases) != 10:
                errors.append("ProductUi form case count must be exactly 10")
                form_cases = []
            if [case.get("form") for case in form_cases if isinstance(case, dict)] != expected_forms:
                errors.append("ProductUi form sample inventory is invalid")
            for case in form_cases:
                fields = {
                    "form", "feature_id", "ui_surface", "caption_sha256", "native_handle",
                    "bounds", "width_ratio", "height_ratio", "inside_work_area", "default_control",
                    "cancel_control", "close_method", "control_attestation", "control_attestation_sha256", "workbook_before_sha256",
                    "workbook_after_sha256", "forbidden_text_matches",
                }
                if not isinstance(case, dict) or set(case) != fields:
                    errors.append("ProductUi form case schema is invalid")
                    continue
                form_name = case.get("form")
                attested_form = attested_forms.get(form_name)
                attested_controls = attested_form.get("controls", []) if isinstance(attested_form, dict) else []
                attested_default = [
                    control for control in attested_controls
                    if isinstance(control, dict)
                    and control.get("name") == case.get("default_control")
                    and control.get("type") == "CommandButton"
                ]
                cancel_name = case.get("cancel_control")
                attested_cancel = [
                    control for control in attested_controls
                    if isinstance(control, dict)
                    and control.get("name") == cancel_name
                    and control.get("type") == "CommandButton"
                ] if cancel_name is not None else []
                surface = active_surfaces.get(form_name)
                bounds = case.get("bounds")
                width_ratio = case.get("width_ratio")
                height_ratio = case.get("height_ratio")
                ratios_ok = (
                    _valid_bounds(bounds)
                    and client["width"] > 0
                    and client["height"] > 0
                    and isinstance(width_ratio, (int, float))
                    and not isinstance(width_ratio, bool)
                    and isinstance(height_ratio, (int, float))
                    and not isinstance(height_ratio, bool)
                    and abs(width_ratio - bounds["width"] / client["width"]) <= 0.000001
                    and abs(height_ratio - bounds["height"] / client["height"]) <= 0.000001
                    and surface in classes
                    and width_ratio <= classes[surface]["max_width_ratio"]
                    and height_ratio <= classes[surface]["max_height_ratio"]
                )
                if (
                    case.get("ui_surface") != surface
                    or case.get("feature_id") != expected_features.get(form_name)
                    or not ratios_ok
                    or case.get("inside_work_area") is not True
                    or case.get("control_attestation") != "ProductUi.UserForms.json"
                    or case.get("control_attestation_sha256") != attestation_sha
                    or not isinstance(case.get("native_handle"), int)
                    or isinstance(case.get("native_handle"), bool)
                    or case.get("native_handle", 0) <= 0
                    or not isinstance(case.get("caption_sha256"), str)
                    or not SHA64.fullmatch(case["caption_sha256"])
                    or not isinstance(case.get("default_control"), str)
                    or not case.get("default_control")
                    or not (
                        case.get("cancel_control") is None
                        or isinstance(case.get("cancel_control"), str) and bool(case.get("cancel_control"))
                    )
                    or case.get("close_method") not in {"escape-cancel", "alt-f4"}
                    or not isinstance(attested_form, dict)
                    or attested_form.get("default_control") != case.get("default_control")
                    or attested_form.get("cancel_control") != cancel_name
                    or len(attested_default) != 1
                    or (cancel_name is not None and len(attested_cancel) != 1)
                    or case.get("workbook_before_sha256") != case.get("workbook_after_sha256")
                    or not isinstance(case.get("workbook_before_sha256"), str)
                    or not SHA64.fullmatch(case["workbook_before_sha256"])
                    or case.get("forbidden_text_matches") != []
                ):
                    errors.append(f"ProductUi form evidence is invalid: {form_name}")
    except (OSError, UnicodeDecodeError, zipfile.BadZipFile, ValueError, KeyError, TypeError) as exc:
        errors.append(str(exc))
    if not errors:
        result.update({
            "verdict": "PASS", "run_id": run_id, "packaging_commit": manifest.get("packaging_commit"),
            "artifact_sha256": artifact, "source_tree_sha256": source_tree,
            "source_snapshot_sha256": source_snapshot, "product_bundle_sha256": manifest.get("product_bundle_sha256"),
            "product_files": product_files,
            "evidence_files": ["ProductUi.Current200.json", "ProductUi.UserForms.json"],
            "evidence_sha256": {
                "ProductUi.Current200.json": _sha(ui_bytes),
                "ProductUi.UserForms.json": attestation_sha,
            },
            "return_zip_sha256": _sha(archive_path.read_bytes()),
        })
    return result


def inspect_return(input_dir: str | Path, suite: str = "ProductRibbon") -> dict[str, Any]:
    if suite == "ProductUi":
        return _inspect_product_ui_return(input_dir)
    root = Path(input_dir)
    result: dict[str, Any] = {"suite": suite, "verdict": "DIAGNOSTIC", "errors": []}
    errors: list[str] = result["errors"]
    if suite != "ProductRibbon":
        errors.append("unsupported suite")
        return result
    if not root.is_dir():
        errors.append("return directory is missing")
        return result
    candidates = [p for p in (root / "PASS.zip", root / "DIAGNOSTIC.zip") if p.is_file()]
    if len(candidates) != 1:
        errors.append("return must contain exactly one PASS.zip or DIAGNOSTIC.zip")
        return result
    archive_path = candidates[0]
    result["return_zip"] = archive_path.name
    evidence_hashes: dict[str, str] = {}
    try:
        with zipfile.ZipFile(archive_path) as archive:
            members = [n for n in archive.namelist() if not n.endswith("/")]
            if len(set(members)) != len(members) or any("\\" in n or n.startswith("/") or re.search(r'[:*?"<>|]', n) or any(p in {"", ".", ".."} for p in n.split("/")) for n in members):
                errors.append("unsafe or duplicate archive member path")
                return result
            if archive_path.name == "DIAGNOSTIC.zip":
                _inspect_product_diagnostic(archive, members, archive_path, result, errors)
                result["diagnostic_valid"] = not errors
                return result
            by_path = set(members)
            missing = [name for name in REQUIRED if name not in by_path]
            if missing:
                errors.append("required evidence missing: " + ", ".join(missing))
                return result
            if any(name != Path(name).name for name in EVIDENCE_JSON + (SUMMARY_JSON, "Product.xlam", "manifest.json", "SHA256SUMS") if name in by_path):
                errors.append("root evidence members must not be nested")
                return result
            manifest = _json(archive.read("manifest.json"), "manifest.json")
            summary = _json(archive.read(SUMMARY_JSON), SUMMARY_JSON)
            evidence = {name: _json(archive.read(name), name) for name in EVIDENCE_JSON}
            evidence_hashes = {name: _sha(archive.read(name)) for name in EVIDENCE_JSON}
            allowed_manifest = {"schema_version", "suite", "run_id", "packaging_commit", "artifact_sha256", "source_tree_sha256", "source_snapshot_sha256", "product_bundle_sha256", "product_files"}
            _exact(manifest, allowed_manifest, "manifest.json", errors)
            if manifest.get("schema_version") != 1:
                errors.append("manifest schema_version must be exactly 1")
            manifest_suite = manifest.get("suite")
            if manifest_suite not in {"ProductRibbon", "ProductUi"}:
                errors.append("manifest suite must be ProductRibbon or composite ProductUi")
            run_id = manifest.get("run_id")
            if not isinstance(run_id, str) or not UUID_RE.fullmatch(run_id):
                errors.append("manifest run_id missing or malformed")
            if not isinstance(manifest.get("product_bundle_sha256"), str) or not SHA64.fullmatch(manifest["product_bundle_sha256"]):
                errors.append("manifest product_bundle_sha256 missing or malformed")
            product_files = manifest.get("product_files")
            if not isinstance(product_files, list) or any(not isinstance(item, dict) or set(item) != {"path", "sha256"} or not isinstance(item.get("path"), str) or not SHA64.fullmatch(str(item.get("sha256"))) for item in product_files):
                errors.append("manifest product_files schema is invalid")
                product_files = []
            product_paths = [item["path"] for item in product_files]
            if product_paths != sorted(product_paths, key=lambda path: path.encode("utf-8")) or len(set(product_paths)) != len(product_paths):
                errors.append("manifest product_files must be unique and UTF-8 ordinal sorted")
            if any("\\" in path or path.startswith("/") or re.search(r'[:*?"<>|]', path) or any(p in {"", ".", ".."} for p in path.split("/")) for path in product_paths):
                errors.append("manifest product_files contains unsafe path")
            packaging_commit = manifest.get("packaging_commit")
            if not isinstance(packaging_commit, str) or not SHA40.fullmatch(packaging_commit):
                errors.append("manifest packaging_commit missing or malformed")
            summary_fields = {"schema_version", "suite", "status", "mode", "run_id", "source_digest", "snapshot_digest", "artifact_sha256", "environment", "cleanup", "run"}
            _exact(summary, summary_fields, SUMMARY_JSON, errors)
            if summary.get("schema_version") != 1 or summary.get("suite") != "ProductRibbon" or summary.get("status") != "PASS" or summary.get("mode") != "Green":
                errors.append("ProductRibbon.json identity or status is invalid")
            if summary.get("run_id") != run_id:
                errors.append("ProductRibbon.json run_id is not bound to manifest")
            environment = summary.get("environment")
            if (
                not isinstance(environment, dict)
                or set(environment) != {"excel_bitness", "office", "compatibility_baseline"}
                or environment.get("excel_bitness") != "64-bit"
                or environment.get("office") not in {"Office 2024", "Microsoft 365"}
                or environment.get("compatibility_baseline") != "Office 2024"
            ):
                errors.append("ProductRibbon.json environment is invalid")
            cleanup = summary.get("cleanup")
            if not isinstance(cleanup, dict) or set(cleanup) != {"status", "exit_mode", "failures"} or cleanup.get("status") != "PASS" or cleanup.get("exit_mode") not in {"NATURAL", "FORCED_CONTAINED"} or cleanup.get("failures") != []:
                errors.append("ProductRibbon.json cleanup proof is invalid")
            run = summary.get("run")
            if not isinstance(run, dict) or set(run) != {"total", "passed", "failed", "skipped"} or run != {"total": 3, "passed": 3, "failed": 0, "skipped": 0}:
                errors.append("ProductRibbon.json run summary is invalid")
            for name, doc in evidence.items():
                _exact(doc, COMMON_FIELDS, name, errors)
                measured = doc.get("measured")
                if not isinstance(measured, dict) or set(measured) != MEASURED_FIELDS[name]:
                    errors.append(f"{name} measured schema keys must be exact")
                if doc.get("schema_version") != 1 or doc.get("suite") != "ProductRibbon" or doc.get("evidence") != EVIDENCE_LABELS[name]:
                    errors.append(f"{name} producer identity is invalid")
                if doc.get("run_id") != run_id:
                    errors.append(f"{name} run_id is not bound to manifest")
                if not isinstance(doc.get("method"), str) or not doc["method"] or not isinstance(measured, dict) or measured.get("method") != doc.get("method"):
                    errors.append(f"{name} method binding is invalid")
            expected_files_set = set(EVIDENCE_JSON) | {SUMMARY_JSON, "Product.xlam"}
            if manifest_suite == "ProductUi":
                expected_files_set |= set(PRODUCT_UI_EVIDENCE) | {PRODUCT_UI_SUMMARY}
            expected_files = sorted(expected_files_set)
            declared_product = {item["path"]: item["sha256"] for item in product_files}
            expected_product = {"Product.xlam": artifact if isinstance((artifact := manifest.get("artifact_sha256")), str) else ""}
            if declared_product != expected_product:
                errors.append("manifest product_files must contain exactly Product.xlam")
            bundle_payload = "".join(f"{path}\n{declared_product[path]}\n" for path in sorted(declared_product, key=lambda path: path.encode("utf-8")))
            if isinstance(manifest.get("product_bundle_sha256"), str) and _sha(bundle_payload.encode()) != manifest.get("product_bundle_sha256"):
                errors.append("manifest product_bundle_sha256 mismatch")
            records = _sum_records(archive.read("SHA256SUMS").decode("utf-8-sig"))
            if sorted(records) != expected_files:
                errors.append("SHA256SUMS does not cover the exact evidence set")
            for name, digest in records.items():
                if name not in by_path or _sha(archive.read(name)) != digest:
                    errors.append(f"SHA256SUMS digest mismatch: {name}")
            if set(members) != {"manifest.json", "SHA256SUMS"} | set(expected_files):
                errors.append("archive contains extra or missing members")
            # Every evidence document must bind to manifest identity.
            artifact = manifest.get("artifact_sha256")
            source_tree = manifest.get("source_tree_sha256")
            source_snapshot = manifest.get("source_snapshot_sha256")
            if not isinstance(artifact, str) or not SHA64.fullmatch(artifact):
                errors.append("manifest artifact SHA-256 missing or malformed")
            elif _sha(archive.read("Product.xlam")) != artifact:
                errors.append("Product.xlam artifact SHA-256 mismatch")
            if not isinstance(source_tree, str) or not SHA64.fullmatch(source_tree):
                errors.append("manifest source tree SHA-256 missing or malformed")
            if not isinstance(source_snapshot, str) or not SHA64.fullmatch(source_snapshot):
                errors.append("manifest source snapshot SHA-256 missing or malformed")
            for field, expected in (("artifact_sha256", artifact), ("source_tree_sha256", source_tree), ("source_snapshot_sha256", source_snapshot)):
                values = [manifest.get(field)]
                if not isinstance(expected, str) or values[0] != expected:
                    errors.append(f"manifest {field} values disagree")
            for name, doc in evidence.items():
                for field, expected in (("artifact_sha256", artifact), ("source_tree_sha256", source_tree), ("source_snapshot_sha256", source_snapshot)):
                    values = [doc.get(field)]
                    if not isinstance(expected, str) or values[0] != expected:
                        errors.append(f"{name} {field} is not bound to manifest")
            if summary.get("artifact_sha256") != artifact or summary.get("source_digest") != source_tree or summary.get("snapshot_digest") != source_snapshot:
                errors.append("ProductRibbon.json digest identity is not bound to manifest")
            for name, doc in evidence.items():
                if doc.get("status") != "PASS":
                    errors.append(f"{name} status must be PASS")
            package = evidence["ProductRibbon.Package.json"]["measured"]
            if package.get("custom_ui_parts") != 1 or package.get("relationship_type") != "http://schemas.microsoft.com/office/2007/relationships/ui/extensibility":
                errors.append("Package evidence actual OOXML checks failed")
            ribbon = evidence["ProductRibbon.RibbonXml.json"]["measured"]
            callbacks = ribbon.get("callbacks")
            ribbon_hashes = (ribbon.get("embedded_xml_sha256"), ribbon.get("source_xml_sha256"))
            ribbon_ok = (
                ribbon.get("compile_control_id") == 578
                and ribbon.get("compile_saved") is True
                and ribbon.get("active_project_artifact_name") == "Product.xlam"
                and _is_int(ribbon.get("reference_count"))
                and ribbon["reference_count"] > 0
                and ribbon.get("broken_reference_count") == 0
                and ribbon.get("group_count") == 12
                and ribbon.get("button_count") == EXPECTED_RIBBON_BUTTON_COUNT
                and ribbon.get("feature_count") == EXPECTED_FEATURE_COUNT
                and callbacks == list(PRODUCT_RIBBON_CALLBACKS)
                and all(isinstance(value, str) and SHA64.fullmatch(value) for value in ribbon_hashes)
                and ribbon_hashes[0] == ribbon_hashes[1]
                and ribbon.get("macro_execution") == "actual-ribbon-ui"
                and ribbon.get("grade_policy_contract_attested") is True
                and ribbon.get("grade_policy_pattern_count") == 8
            )
            if not ribbon_ok:
                errors.append("RibbonXml evidence actual checks failed")
            identity = evidence["ProductRibbon.OfficeIdentity.json"]
            office = identity["measured"].get("office_version")
            compatibility_baseline = identity["measured"].get("compatibility_baseline")
            arch = identity["measured"].get("architecture")
            product_release_ids = identity["measured"].get("product_release_ids")
            office_ids_valid = (
                isinstance(product_release_ids, list)
                and bool(product_release_ids)
                and all(isinstance(value, str) and value.strip() for value in product_release_ids)
                and (
                    (office == "Office 2024" and any("2024" in value for value in product_release_ids))
                    or (office == "Microsoft 365" and any(re.match(r"^(?:O365|M365)", value, re.I) for value in product_release_ids))
                )
            )
            if office not in {"Office 2024", "Microsoft 365"} or compatibility_baseline != "Office 2024" or arch != "x64" or not office_ids_valid:
                errors.append("Office 2024 compatibility baseline with Office 2024 or Microsoft 365 x64 native evidence required")
            current_ui = evidence["ProductRibbon.Ui.Current200.json"]["measured"]
            keytips = current_ui.get("keytips")
            home_controls = current_ui.get("home_controls")
            product_groups = current_ui.get("product_groups")
            window_bounds = current_ui.get("excel_window_bounds")
            client_bounds = current_ui.get("excel_client_bounds")
            bounds_ok = all(
                isinstance(row, dict)
                and all(isinstance(row.get(key), (int, float)) and not isinstance(row.get(key), bool) for key in ("left", "top", "width", "height"))
                and row.get("width", 0) > 0
                and row.get("height", 0) > 0
                for row in (window_bounds, client_bounds)
            )
            home_labels = [row.get("label") for row in home_controls] if isinstance(home_controls, list) else []
            product_labels = [row.get("label") for row in product_groups] if isinstance(product_groups, list) else []
            control_rows = (home_controls if isinstance(home_controls, list) else []) + (product_groups if isinstance(product_groups, list) else [])
            control_bounds_ok = all(
                isinstance(row, dict)
                and all(isinstance(row.get(key), (int, float)) and not isinstance(row.get(key), bool) and row.get(key) > 0 for key in ("width", "height"))
                for row in control_rows
            )
            keytips_ok = isinstance(keytips, list) and bool(keytips) and all(isinstance(item, str) and item.strip() for item in keytips)
            current_ui_ok = (
                current_ui.get("windows_scale_percent") == 200
                and current_ui.get("host_dpi") == 192
                and current_ui.get("scale_changed_by_test") is False
                and current_ui.get("foreground_verified") is True
                and current_ui.get("keytip_accessible") is True
                and keytips_ok
                and bounds_ok
                and control_bounds_ok
                and home_labels == ["내엑셀", "즐겨찾기", "포커스셀", "전체기능"]
                and product_labels == ["내엑셀", "저장", "인쇄", "복붙", "삽입", "파일관리", "템플릿", "데이터", " ", "스타일", "추가기능", "정보진단"]
            )
            if not current_ui_ok:
                errors.append("ProductRibbon.Ui.Current200.json lacks passing current 200 percent UI evidence")
            route_features = {
                "Fast": "NX-DATA-UNIQUE-COUNT",
                "Guarded": "NX-DATA-DUPLICATE-LIST",
                "Planned": "NX-FILE-FOLDER-CREATE",
            }
            for grade in ("Fast", "Guarded", "Planned"):
                doc = evidence[f"ProductRibbon.{grade}.json"]
                expected_resolved = {"Fast": 0, "Guarded": 1, "Planned": 2}[grade]
                measured = doc["measured"]
                if (
                    measured.get("execution_grade") != grade
                    or measured.get("resolved") != expected_resolved
                    or measured.get("route_feature_id") != route_features[grade]
                    or measured.get("route_verified") is not True
                    or measured.get("form_opened") is not True
                    or measured.get("route_proof") != "actual-ribbon-control-to-owned-form"
                    or measured.get("policy_attested_from_compiled_product") is not True
                ):
                    errors.append(f"{grade} execution evidence is not PASS")
            # UserForm attestation is deliberately independent of the three grade probes.
            # It must describe the saved XLAM after a close/reopen VBProject Designer readback.
            forms_doc = evidence["ProductRibbon.UserForms.json"]
            forms = forms_doc["measured"].get("forms")
            bindings = forms_doc["measured"].get("form_bindings")
            # Read only trusted local manifest layouts, never receipt-supplied paths.
            # Older non-fixed ListBox receipts remain valid; explicit contracts
            # require actual persisted IntegralHeight and dimension readback.
            source_root = Path(__file__).resolve().parents[1]
            product_manifest = _json((source_root / "build/manifests/Product.json").read_bytes(), "Product.json")
            fixed_listboxes: dict[tuple[str, str, str, str], dict[str, Any]] = {}
            for binding in product_manifest["form_bindings"]:
                layout = _json((source_root / binding["layout"]).read_bytes(), binding["layout"])
                for definition in layout["controls"]:
                    if "integral_height" not in definition:
                        continue
                    if (
                        definition.get("type") != "ListBox"
                        or type(definition["integral_height"]) is not bool
                        or not all(_finite_positive_number(definition.get(axis)) for axis in ("width", "height"))
                    ):
                        raise ValueError("UserForms fixed ListBox layout contract is invalid")
                    key = (layout["name"], binding["layout"], binding["code"], definition["name"])
                    fixed_listboxes[key] = definition
            seen_fixed_listboxes: set[tuple[str, str, str, str]] = set()
            retired_forms = {
                "FNxStart", "FNxAllFunctions", "FNxData", "FNxDraw", "FNxFile",
                "FNxTemplate", "FNxFocusPalette",
            }
            if forms_doc["method"] != "built Product.xlam close/reopen + VBProject Designer.Controls readback":
                errors.append("UserForms evidence method is not authoritative save/reopen readback")
            if not isinstance(bindings, list) or not bindings:
                errors.append("UserForms form_bindings must exactly match the Product manifest")
                bindings = []
            binding_pairs: set[tuple[str, str]] = set()
            for binding in bindings:
                if not isinstance(binding, dict) or set(binding) != {"layout", "code"}:
                    errors.append("UserForms form_binding schema is invalid")
                    continue
                layout, code = binding.get("layout"), binding.get("code")
                if not isinstance(layout, str) or not isinstance(code, str):
                    errors.append("UserForms form_binding paths are invalid")
                    continue
                form_name = Path(layout).name.removesuffix(".form.json")
                if (
                    not layout.startswith("src/vba/")
                    or not layout.endswith(".form.json")
                    or not code.startswith("src/vba/")
                    or not code.endswith(".vba")
                    or Path(code).stem != form_name
                    or form_name in retired_forms
                ):
                    errors.append("UserForms form_binding path is not an active Product form")
                binding_pairs.add((layout, code))
            if len(binding_pairs) != len(bindings) or not isinstance(forms, list) or len(forms) != len(bindings):
                errors.append("UserForms form_bindings must exactly match the Product manifest")
            if isinstance(forms, list):
                seen_forms: set[str] = set(); total = textboxes = listboxes = 0
                for form in forms:
                    if not isinstance(form, dict) or set(form) != {"name", "layout", "code", "layout_sha256", "code_sha256", "default_control", "cancel_control", "controls"}:
                        errors.append("UserForms form record schema is invalid"); continue
                    name = form.get("name")
                    if not isinstance(name, str) or name in seen_forms: errors.append("UserForms form names must be unique")
                    seen_forms.add(name)
                    if (form.get("layout"), form.get("code")) not in binding_pairs:
                        errors.append("UserForms readback is not bound to a Product form_binding")
                    controls = form.get("controls")
                    if not isinstance(controls, list) or len({c.get("name") for c in controls if isinstance(c, dict)}) != len(controls):
                        errors.append("UserForms control names must be unique")
                        continue
                    tabs = [c.get("tab_index") for c in controls if isinstance(c, dict) and c.get("tab_stop") is True]
                    if len(tabs) != len(set(tabs)): errors.append("UserForms interactive tab order must be unique")
                    total += len(controls)
                    for control in controls:
                        if not isinstance(control, dict) or not {"name", "type", "tab_index", "tab_stop", "enabled"}.issubset(control):
                            errors.append("UserForms control readback schema is invalid"); continue
                        if control["type"] == "TextBox":
                            textboxes += 1
                            if set(control) != {"name", "type", "value", "locked", "multi_line", "tab_index", "tab_stop", "enabled"}: errors.append("UserForms TextBox readback schema is invalid")
                        if control["type"] == "ListBox":
                            listboxes += 1
                            base_fields = {
                                "name", "type", "column_count", "column_widths",
                                "multi_select", "items_sha256", "list_index",
                                "tab_index", "tab_stop", "enabled",
                            }
                            static_item_fields = {
                                "item_count", "selected_index", "list_tag_sha256",
                            }
                            fixed_key = (name, form["layout"], form["code"], control["name"])
                            fixed_contract = fixed_listboxes.get(fixed_key)
                            if fixed_contract is not None:
                                seen_fixed_listboxes.add(fixed_key)
                                base_fields |= {"integral_height", "width", "height"}
                                allowed_fields = (base_fields, base_fields | static_item_fields)
                            else:
                                allowed_fields = (
                                    base_fields, base_fields | static_item_fields,
                                    base_fields | {"integral_height"},
                                    base_fields | {"integral_height"} | static_item_fields,
                                )
                            fields = set(control)
                            if fields not in allowed_fields:
                                errors.append("UserForms ListBox readback schema is invalid")
                                continue
                            if "integral_height" in control and (
                                type(control["integral_height"]) is not bool
                                or (fixed_contract is not None and control["integral_height"] is not fixed_contract["integral_height"])
                            ):
                                errors.append("UserForms ListBox integral_height is invalid")
                            if fixed_contract is not None and not all(
                                _finite_positive_number(control.get(axis))
                                and abs(control[axis] - fixed_contract[axis]) <= 0.051
                                for axis in ("width", "height")
                            ):
                                errors.append("UserForms fixed ListBox dimensions are invalid")
                            if not isinstance(control.get("items_sha256"), str) or not SHA64.fullmatch(control["items_sha256"]):
                                errors.append("UserForms ListBox items hash is invalid")
                            elif static_item_fields.issubset(fields):
                                item_count = control.get("item_count")
                                selected_index = control.get("selected_index")
                                list_tag_sha256 = control.get("list_tag_sha256")
                                if (
                                    type(item_count) is not int
                                    or item_count <= 0
                                    or type(selected_index) is not int
                                    or selected_index < -1
                                    or selected_index >= item_count
                                    or not isinstance(list_tag_sha256, str)
                                    or not SHA64.fullmatch(list_tag_sha256)
                                ):
                                    errors.append("UserForms static ListBox metadata is invalid")
                if seen_fixed_listboxes != set(fixed_listboxes):
                    errors.append("UserForms fixed ListBox readback is missing")
                measured = forms_doc["measured"]
                if (
                    measured.get("total_forms") != len(forms)
                    or measured.get("total_controls") != total
                    or measured.get("textboxes") != textboxes
                    or measured.get("listboxes") != listboxes
                ):
                    errors.append("UserForms readback control totals are invalid")
    except (OSError, UnicodeDecodeError, zipfile.BadZipFile, ValueError, KeyError) as exc:
        errors.append(str(exc))
    if archive_path.name == "DIAGNOSTIC.zip":
        result["diagnostic_valid"] = False
        return result
    result["verdict"] = "PASS" if not errors and archive_path.name == "PASS.zip" else "DIAGNOSTIC"
    if result["verdict"] == "PASS":
        result.update({"run_id": run_id, "packaging_commit": packaging_commit, "artifact_sha256": artifact, "source_tree_sha256": source_tree, "source_snapshot_sha256": source_snapshot, "product_bundle_sha256": manifest.get("product_bundle_sha256"), "product_files": product_files, "office_version": office, "compatibility_baseline": compatibility_baseline, "architecture": arch, "evidence_files": list(EVIDENCE_JSON), "evidence_sha256": evidence_hashes, "return_zip_sha256": _sha(archive_path.read_bytes()), "summary_sha256": _zip_member_sha(archive_path, SUMMARY_JSON), "result_manifest_sha256": _zip_member_sha(archive_path, "manifest.json")})
    return result


inspect_windows_return = inspect_return


def archive_pass(input_dir: str | Path, destination: str | Path, suite: str = "ProductRibbon") -> dict[str, Any]:
    """Atomically materialize one strictly inspected PASS return."""
    root, destination = Path(input_dir), Path(destination)
    result = inspect_return(root, suite)
    if result.get("verdict") != "PASS":
        raise ValueError(f"{suite} return is not strict PASS")
    if destination.exists():
        raise ValueError("archive destination already exists")
    destination.parent.mkdir(parents=True, exist_ok=True)
    temporary = Path(tempfile.mkdtemp(prefix=f".{destination.name}.", dir=destination.parent))
    try:
        with zipfile.ZipFile(root / "PASS.zip") as archive:
            members = [name for name in archive.namelist() if not name.endswith("/")]
            for name in members:
                target = temporary / name
                target.parent.mkdir(parents=True, exist_ok=True)
                target.write_bytes(archive.read(name))
        (temporary / "windows-inspection.json").write_text(json.dumps(result, ensure_ascii=False, indent=2, sort_keys=True) + "\n", encoding="utf-8")
        os.replace(temporary, destination)
    except Exception:
        shutil.rmtree(temporary, ignore_errors=True)
        raise
    return result


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--suite", default="ProductRibbon")
    parser.add_argument("--input", required=True, type=Path)
    parser.add_argument("--output", type=Path)
    parser.add_argument("--archive", type=Path)
    args = parser.parse_args(argv)
    try:
        result = archive_pass(args.input, args.archive, args.suite) if args.archive else inspect_return(args.input, args.suite)
    except (OSError, ValueError) as exc:
        result = {"suite": args.suite, "verdict": "DIAGNOSTIC", "errors": [str(exc)]}
    if args.output:
        args.output.parent.mkdir(parents=True, exist_ok=True)
        with tempfile.NamedTemporaryFile("w", encoding="utf-8", dir=args.output.parent, delete=False) as handle:
            json.dump(result, handle, ensure_ascii=False, indent=2, sort_keys=True)
            handle.write("\n")
            temporary = Path(handle.name)
        os.replace(temporary, args.output)
    print(json.dumps(result, ensure_ascii=False, sort_keys=True))
    return 0 if result["verdict"] == "PASS" else 1


if __name__ == "__main__":
    raise SystemExit(main())
