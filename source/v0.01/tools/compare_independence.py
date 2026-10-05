"""Deterministic adapter and five-axis clean-room independence comparator."""
from __future__ import annotations

import argparse, base64, binascii, hashlib, json, re
from collections import Counter
from copy import deepcopy
from pathlib import Path
from typing import Any, Mapping

AXES = ("code", "workbook", "ui_wording", "assets", "provenance")
LINEAGE_AXES = AXES[:4]
STATUSES = {"PASS", "WARN", "FAIL", "INCOMPLETE"}
_RANK = {"PASS": 0, "WARN": 1, "INCOMPLETE": 2, "FAIL": 3}
SEMANTIC_MIN_COUNTS = {"workbook": 1, "ui_wording": 1, "assets": 1}
_HEX64 = re.compile(r"^[0-9a-fA-F]{64}$")
_VBA_EXPORT_BOILERPLATE = re.compile(
    r"^(?:VERSION\s+\d+(?:\.\d+)?\s+CLASS|BEGIN|END|MULTIUSE\s*=.*|OPTION\s+EXPLICIT)$",
    re.I,
)
_VBA_PROCEDURE_START = re.compile(
    r"^\s*(?:public\s+|private\s+|friend\s+)?(?:sub|function|property\s+(?:get|let|set))\s+(\w+)",
    re.I,
)
_VBA_PROCEDURE_END = re.compile(r"^\s*end\s+(?:sub|function|property)\s*$", re.I)

APPROVED_LINEAGE_POLICY = {
    "version": 2,
    "policy": "approved-first-party-v3.4-lineage",
    "approved_reference_identity": "v3x",
    "scope": "all-functions",
    "style_authority": "내엑셀 v0.01",
    "integration_requirements": [
        "v0.01-ribbon-and-registry",
        "v0.01-risk-and-recovery",
        "v0.01-native-verification",
    ],
    "forbidden_reference_identities": ["original", "abtools"],
}
APPROVED_LINEAGE_SOURCE_KEYS = {
    "status",
    "reference_identity",
    "inventory_sha256",
    "scope",
    "style_authority",
    "integration_requirements",
}

def approved_lineage_policy_contract() -> dict[str, Any]:
    """Return the closed whole-lineage v3.4 adoption policy."""
    return deepcopy(APPROVED_LINEAGE_POLICY)


def validate_approved_lineage_policy(policy: Mapping[str, Any]) -> list[str]:
    """Validate that only first-party v3.4 is approved and v0.01 remains authoritative."""
    if not isinstance(policy, Mapping):
        return ["approved lineage policy must be an object"]
    errors: list[str] = []
    expected = APPROVED_LINEAGE_POLICY
    if set(policy) != set(expected):
        errors.append("approved lineage policy keys are not closed")
    for key, expected_value in expected.items():
        actual = policy.get(key)
        if actual != expected_value or type(actual) is not type(expected_value):
            errors.append(f"approved lineage policy {key} is not the approved value")
    approved = policy.get("approved_reference_identity")
    forbidden = policy.get("forbidden_reference_identities")
    if approved != "v3x":
        errors.append("only the first-party v3.4 identity may be approved")
    if not isinstance(forbidden, list) or "original" not in forbidden or "abtools" not in forbidden:
        errors.append("external originals and abTools must remain forbidden")
    if isinstance(forbidden, list) and approved in forbidden:
        errors.append("approved and forbidden lineage identities overlap")
    return errors


def approved_lineage_source_contract(
    inventory_sha256: str,
    policy: Mapping[str, Any] | None = None,
) -> dict[str, Any]:
    """Bind the approved v3.4 source inventory without treating it as an independence reference."""
    current = policy if policy is not None else APPROVED_LINEAGE_POLICY
    return {
        "status": "PASS",
        "reference_identity": current.get("approved_reference_identity"),
        "inventory_sha256": inventory_sha256,
        "scope": current.get("scope"),
        "style_authority": current.get("style_authority"),
        "integration_requirements": deepcopy(current.get("integration_requirements")),
    }


def validate_approved_lineage_source(
    source: Any,
    policy: Mapping[str, Any] | None = None,
) -> list[str]:
    """Validate the exact release-evidence projection for the approved v3.4 source."""
    expected_policy = policy if policy is not None else APPROVED_LINEAGE_POLICY
    policy_errors = validate_approved_lineage_policy(expected_policy)
    if policy_errors:
        return [f"approved lineage source policy: {error}" for error in policy_errors]
    if not isinstance(source, Mapping):
        return ["approved lineage source must be an object"]
    errors: list[str] = []
    if set(source) != APPROVED_LINEAGE_SOURCE_KEYS:
        errors.append("approved lineage source keys are not closed")
    expected = approved_lineage_source_contract(str(source.get("inventory_sha256", "")), expected_policy)
    for key, expected_value in expected.items():
        actual = source.get(key)
        if actual != expected_value or type(actual) is not type(expected_value):
            errors.append(f"approved lineage source {key} is not the approved value")
    digest = source.get("inventory_sha256")
    if not isinstance(digest, str) or not _HEX64.fullmatch(digest) or digest != digest.lower():
        errors.append("approved lineage source inventory SHA-256 is invalid")
    return errors

def _sha(value: str) -> str: return hashlib.sha256(value.encode("utf-8")).hexdigest()
def _inventory_digest(records: Any) -> str:
    if not isinstance(records, list):
        return ""
    canonical = "".join(f"{row['path']}\n{row['sha256']}\n{row['size']}\n" for row in records)
    return _sha(canonical)
def _norm(text: str) -> str:
    text = re.sub(r"(?im)^\s*attribute\s+vb_[^\n]*\n?", "", text)
    text = re.sub(r"\r\n?", "\n", text)
    return "\n".join(re.sub(r"[ \t]+", " ", x).strip() for x in text.splitlines()).strip()
def _tokens(text: str) -> set[str]: return set(re.findall(r"[A-Za-z_][A-Za-z0-9_]*", text.lower()))
def _ngram(text: str, n: int = 5) -> list[str]:
    words = re.findall(r"\w+", text.lower(), re.UNICODE)
    return [" ".join(words[i:i+n]) for i in range(max(0, len(words)-n+1))]
def _values(subject: Mapping[str, Any], axis: str) -> list[str]:
    value = subject.get(axis)
    if isinstance(value, Mapping): value = value.get("items", value.get("values", value.get("hashes", [])))
    if isinstance(value, str): return [value]
    if isinstance(value, (list, tuple, set)): return [str(x) for x in value]
    return []
def _axis(subject: Mapping[str, Any], axis: str) -> dict[str, Any]:
    value = subject.get(axis)
    return dict(value) if isinstance(value, Mapping) else {"items": _values(subject, axis), "complete": bool(_values(subject, axis))}

def _module_subject(vba_dir: Path) -> dict[str, Any]:
    modules = []
    for path in sorted(vba_dir.rglob("*")) if vba_dir.is_dir() else []:
        if not path.is_file(): continue
        try: raw = path.read_text(encoding="utf-8", errors="replace")
        except OSError: continue
        normalized = _norm(raw)
        rows: list[str] = []
        row_procedures: list[str] = []
        current_procedure = ""
        for row in normalized.splitlines():
            if not row or row.startswith("'") or _VBA_EXPORT_BOILERPLATE.fullmatch(row):
                continue
            declaration = _VBA_PROCEDURE_START.match(row)
            if declaration:
                current_procedure = declaration.group(1)
            rows.append(row)
            row_procedures.append(current_procedure)
            if _VBA_PROCEDURE_END.match(row):
                current_procedure = ""
        procedures = []
        for m in re.finditer(r"(?im)^\s*(?:public\s+|private\s+|friend\s+)?(?:sub|function|property\s+(?:get|let|set))\s+(\w+)[^\n]*\n(.*?)(?=^\s*end\s+(?:sub|function|property)|\Z)", normalized, re.S):
            body = m.group(0); procedures.append({"name": m.group(1), "tokens": sorted(_tokens(body)), "token_count": len(re.findall(r"[A-Za-z_][A-Za-z0-9_]*", body)), "structure": _sha(re.sub(r"\w+", "X", body.lower()))})
        semantic_text = "\n".join(rows)
        modules.append({"path": str(path.relative_to(vba_dir)), "hash": _sha(semantic_text), "rows": rows, "row_procedures": row_procedures, "procedures": procedures, "tokens": sorted(_tokens(semantic_text)), "semantic": bool(rows)})
    semantic_modules = [module for module in modules if module["semantic"]]
    return {"items": [m["hash"] for m in semantic_modules], "modules": semantic_modules, "complete": bool(semantic_modules)}

def _workbook_subject(report: Mapping[str, Any]) -> dict[str, Any]:
    entries = report.get("cell_inventory", {})
    non_boiler = []
    if isinstance(entries, Mapping):
        for sheet, data in entries.items():
            for cell in (data.get("occupied_cells", []) if isinstance(data, Mapping) else []):
                val = str(cell.get("value", "") if isinstance(cell, Mapping) else cell).strip()
                if val and val.lower() not in {"true", "false", "0", "1"}: non_boiler.append(f"{sheet}!{val}")
    topo = report.get("sheet_topology", []); rel = report.get("relationships", []); names = report.get("defined_names", []); styles = report.get("styles", {})
    sections = [x for x in (non_boiler, topo, rel, names, styles) if x]
    items = [_sha(json.dumps(x, sort_keys=True, ensure_ascii=False)) for x in sections]
    complete = (all(k in report for k in ("sheet_topology", "relationships", "defined_names", "styles", "cell_inventory"))
                and report.get("status") != "INCOMPLETE" and len(items) >= SEMANTIC_MIN_COUNTS["workbook"])
    return {"items": items, "entry_hashes": [_sha(x) for x in non_boiler], "topology": topo, "relationships": rel, "defined_names": names, "styles": styles, "complete": complete}

def _ui_subject(report: Mapping[str, Any]) -> dict[str, Any]:
    ribbon = report.get("ribbon", {}) if isinstance(report.get("ribbon", {}), Mapping) else {}
    labels = [str(x) for x in ribbon.get("labels", [])]; callbacks = [str(x.get("value", x)) if isinstance(x, Mapping) else str(x) for x in ribbon.get("callbacks", [])]
    wording = labels + callbacks
    items = sorted(set(labels + callbacks + _ngram(" ".join(wording))))
    return {"items": items, "labels": labels, "callbacks": callbacks, "ngrams": _ngram(" ".join(wording)), "complete": "ribbon" in report and len(labels) + len(callbacks) >= SEMANTIC_MIN_COUNTS["ui_wording"]}

def _asset_subject(report: Mapping[str, Any]) -> dict[str, Any]:
    assets = report.get("assets", []); hashes = [str(x.get("sha256", x)) if isinstance(x, Mapping) else str(x) for x in assets]
    complete = "assets" in report and len(hashes) >= SEMANTIC_MIN_COUNTS["assets"] and all(_HEX64.fullmatch(x) for x in hashes)
    return {"items": sorted(hashes), "hashes": sorted(hashes), "complete": complete}

def make_subject(report: Mapping[str, Any] | None = None, vba_dir: str | Path | None = None, provenance: Mapping[str, Any] | None = None) -> dict[str, Any]:
    report = report or {}; code = _module_subject(Path(vba_dir)) if vba_dir else {"items": [], "modules": [], "complete": False}
    prov = dict(provenance or report.get("provenance", {}))
    return {"schema_version": 2, "code": code, "workbook": _workbook_subject(report), "ui_wording": _ui_subject(report), "assets": _asset_subject(report), "provenance": prov}

def subject_from_path(path: str | Path, role: str | None = None) -> dict[str, Any]:
    p = Path(path); report = {}
    if role not in {None, "candidate", "reference"}: raise ValueError("role must be candidate or reference")
    is_reference = role == "reference" or (role is None and (p.name.startswith("reference-") or "reference" in p.name.lower()))
    if p.is_file(): report = json.loads(p.read_text(encoding="utf-8"))
    elif p.is_dir():
        package_names = ("reference-package.json", "package.json", "candidate-package.json") if is_reference else ("candidate-package.json", "package.json", "reference-package.json")
        for name in package_names:
            if (p / name).is_file(): report = json.loads((p / name).read_text(encoding="utf-8")); break
    if isinstance(report, Mapping) and all(k in report for k in LINEAGE_AXES): return report
    base = p if p.is_dir() else p.parent
    preferred_vba = base / ("reference-vba" if is_reference else "candidate-vba")
    vba = preferred_vba if preferred_vba.is_dir() else None
    provenance = {}
    manifest = base / "access-manifest.json"
    if manifest.is_file():
        try:
            m = json.loads(manifest.read_text(encoding="utf-8-sig"))
            if is_reference:
                inventory = m.get("reference_root_inventory", {})
                reference_tree = inventory.get("inventory_sha256", "") if isinstance(inventory, Mapping) else ""
                provenance = {"frozen_source_commit": m.get("source_commit", ""), "artifact_sha256": m.get("reference_candidate_sha256", report.get("package_sha256", "")), "source_tree_sha256": reference_tree, "origin": "reference", "license": "Reference", "spec_ids": [], "reference_identity": m.get("reference_identity", base.name), "reference_root_inventory": inventory}
            else:
                provenance = {"frozen_source_commit": m.get("source_commit", ""), "artifact_sha256": m.get("candidate_sha256", ""), "source_tree_sha256": m.get("source_tree", ""), "origin": "project-original", "license": "Proprietary", "spec_ids": [m.get("source_commit", "")]}
        except (OSError, ValueError): pass
    if not provenance and report:
        provenance = {"artifact_sha256": report.get("package_sha256", ""), "source_tree_sha256": "", "origin": "reference", "license": "Reference", "spec_ids": []}
    return make_subject(report, vba, provenance)

def _longest_run(a: list[str], b: list[str]) -> int:
    best = 0
    for i, x in enumerate(a):
        for j, y in enumerate(b):
            k = 0
            while i+k < len(a) and j+k < len(b) and a[i+k] == b[j+k]: k += 1
            best = max(best, k)
    return best


def _matching_runs(left: Mapping[str, Any], right: Mapping[str, Any], minimum: int = 5) -> list[dict[str, Any]]:
    """Return deterministic maximal equal-row runs between two VBA modules."""
    left_rows = [str(row) for row in left.get("rows", [])]
    right_rows = [str(row) for row in right.get("rows", [])]
    left_procedures = [str(value) for value in left.get("row_procedures", [])]
    right_procedures = [str(value) for value in right.get("row_procedures", [])]
    if len(left_procedures) != len(left_rows):
        left_procedures = [""] * len(left_rows)
    if len(right_procedures) != len(right_rows):
        right_procedures = [""] * len(right_rows)
    positions: dict[str, list[int]] = {}
    for index, row in enumerate(right_rows):
        positions.setdefault(row, []).append(index)
    runs: list[dict[str, Any]] = []
    for left_index, row in enumerate(left_rows):
        for right_index in positions.get(row, []):
            if left_index and right_index and left_rows[left_index - 1] == right_rows[right_index - 1]:
                continue
            count = 0
            while (
                left_index + count < len(left_rows)
                and right_index + count < len(right_rows)
                and left_rows[left_index + count] == right_rows[right_index + count]
            ):
                count += 1
            if count < minimum:
                continue
            matching_rows = left_rows[left_index:left_index + count]
            left_owner = left_procedures[left_index:left_index + count]
            right_owner = right_procedures[right_index:right_index + count]
            candidate_procedure = left_owner[0] if left_owner and left_owner[0] and all(value == left_owner[0] for value in left_owner) else ""
            reference_procedure = right_owner[0] if right_owner and right_owner[0] and all(value == right_owner[0] for value in right_owner) else ""
            runs.append({
                "candidate_start": left_index,
                "reference_start": right_index,
                "count": count,
                "rows_sha256": _sha("\n".join(matching_rows)),
                "candidate_procedure": candidate_procedure,
                "reference_procedure": reference_procedure,
            })
    return sorted(runs, key=lambda item: (item["candidate_start"], item["reference_start"], -item["count"]))


def _module_from_rows(path: str, rows: list[str], row_procedures: list[str]) -> dict[str, Any]:
    procedures = []
    seen: set[str] = set()
    for procedure in row_procedures:
        if not procedure or procedure in seen:
            continue
        seen.add(procedure)
        procedure_rows = [row for row, owner in zip(rows, row_procedures) if owner == procedure]
        body = "\n".join(procedure_rows)
        procedures.append({
            "name": procedure,
            "tokens": sorted(_tokens(body)),
            "token_count": len(re.findall(r"[A-Za-z_][A-Za-z0-9_]*", body)),
            "structure": _sha(re.sub(r"\w+", "X", body.lower())),
        })
    semantic_text = "\n".join(rows)
    return {
        "path": path,
        "hash": _sha(semantic_text),
        "rows": rows,
        "row_procedures": row_procedures,
        "procedures": procedures,
        "tokens": sorted(_tokens(semantic_text)),
        "semantic": bool(rows),
    }


def _replace_approved_run(
    code: Mapping[str, Any], module_path: str, start: int, count: int, boundary: str
) -> dict[str, Any]:
    rebuilt = []
    replaced = False
    for module in code.get("modules", []):
        if str(module.get("path", "")) != module_path:
            rebuilt.append(dict(module))
            continue
        if replaced:
            raise ValueError("approved lineage module path is not unique")
        rows = [str(row) for row in module.get("rows", [])]
        owners = [str(value) for value in module.get("row_procedures", [])]
        if len(owners) != len(rows) or start < 0 or count < 1 or start + count > len(rows):
            raise ValueError("approved lineage run boundaries are invalid")
        owner = owners[start]
        rows = rows[:start] + [boundary] + rows[start + count:]
        owners = owners[:start] + [owner] + owners[start + count:]
        rebuilt.append(_module_from_rows(module_path, rows, owners))
        replaced = True
    if not replaced:
        raise ValueError("approved lineage module was not found")
    semantic_modules = [module for module in rebuilt if module.get("semantic")]
    return {
        "items": [module["hash"] for module in semantic_modules],
        "modules": semantic_modules,
        "complete": bool(semantic_modules),
    }


def _select_module(code: Mapping[str, Any], exact_name: str, side: str) -> Mapping[str, Any]:
    matches = [
        module for module in code.get("modules", [])
        if re.split(r"[\\/]", str(module.get("path", "")))[-1] == exact_name
    ]
    if len(matches) != 1:
        raise ValueError(f"approved lineage {side} module must resolve exactly once")
    return matches[0]


def _embedded_resource_sha256(rows: list[str]) -> str:
    pieces = []
    for row in rows:
        match = re.search(r'&\s*"([A-Za-z0-9+/=]+)"\s*$', row)
        if match:
            pieces.append(match.group(1))
    if not pieces:
        raise ValueError("approved lineage resource Base64 was not found")
    try:
        decoded = base64.b64decode("".join(pieces), validate=True)
    except (binascii.Error, ValueError) as exc:
        raise ValueError("approved lineage resource Base64 is invalid") from exc
    return hashlib.sha256(decoded).hexdigest()


def _resolve_approved_code_segment(
    candidate_code: Mapping[str, Any], reference_code: Mapping[str, Any], entry: Mapping[str, Any]
) -> dict[str, Any]:
    candidate_module = _select_module(candidate_code, str(entry.get("candidate_module", "")), "candidate")
    reference_module = _select_module(reference_code, str(entry.get("reference_module", "")), "reference")
    if candidate_module.get("hash") != entry.get("candidate_module_semantic_sha256"):
        raise ValueError("approved lineage candidate module semantic hash does not match")
    if reference_module.get("hash") != entry.get("reference_module_semantic_sha256"):
        raise ValueError("approved lineage reference module semantic hash does not match")
    count = entry.get("meaningful_row_count")
    if isinstance(count, bool) or not isinstance(count, int) or count < 1:
        raise ValueError("approved lineage meaningful row count is invalid")
    runs = _matching_runs(candidate_module, reference_module, minimum=count)
    matches = [
        run for run in runs
        if run["count"] == count
        and run["rows_sha256"] == entry.get("matching_rows_sha256")
        and run["candidate_procedure"] == entry.get("candidate_procedure")
        and run["reference_procedure"] == entry.get("reference_procedure")
    ]
    if len(matches) != 1:
        raise ValueError("approved lineage procedure run did not resolve exactly once")
    run = matches[0]
    candidate_rows = candidate_module["rows"][run["candidate_start"]:run["candidate_start"] + count]
    reference_rows = reference_module["rows"][run["reference_start"]:run["reference_start"] + count]
    candidate_resource_sha = ""
    reference_resource_sha = ""
    if entry.get("kind", "embedded-resource") == "embedded-resource":
        candidate_resource_sha = _embedded_resource_sha256(candidate_rows)
        reference_resource_sha = _embedded_resource_sha256(reference_rows)
        if candidate_resource_sha != entry.get("resource_sha256") or reference_resource_sha != entry.get("resource_sha256"):
            raise ValueError("approved lineage resource hash does not match")
    evidence = dict(entry)
    evidence.update({
        "status": "PASS",
        "candidate_start": run["candidate_start"],
        "reference_start": run["reference_start"],
    })
    if candidate_resource_sha:
        evidence["candidate_resource_sha256"] = candidate_resource_sha
        evidence["reference_resource_sha256"] = reference_resource_sha
    return {
        "candidate_module_path": str(candidate_module["path"]),
        "candidate_start": run["candidate_start"],
        "reference_module_path": str(reference_module["path"]),
        "reference_start": run["reference_start"],
        "count": count,
        "evidence": evidence,
    }


def _apply_approved_code_entry(
    candidate_code: Mapping[str, Any], reference_code: Mapping[str, Any], entry: Mapping[str, Any]
) -> tuple[dict[str, Any], dict[str, Any], dict[str, Any]]:
    """Remove one exact v3.4-approved procedure run and return auditable evidence."""
    resolution = _resolve_approved_code_segment(candidate_code, reference_code, entry)
    candidate_transformed = _replace_approved_run(
        candidate_code,
        resolution["candidate_module_path"],
        resolution["candidate_start"],
        resolution["count"],
        "__NX_APPROVED_CANDIDATE_BOUNDARY__",
    )
    reference_transformed = _replace_approved_run(
        reference_code,
        resolution["reference_module_path"],
        resolution["reference_start"],
        resolution["count"],
        "~ NX APPROVED REFERENCE BOUNDARY ~",
    )
    return candidate_transformed, reference_transformed, resolution["evidence"]


def _apply_approved_feature_entry(
    candidate_code: Mapping[str, Any], reference_code: Mapping[str, Any], entry: Mapping[str, Any]
) -> tuple[dict[str, Any], dict[str, Any], dict[str, Any]]:
    """Validate all closed segments against original modules, then remove them as one feature."""
    segments = entry.get("segments")
    if not isinstance(segments, list) or not segments:
        raise ValueError("approved lineage feature has no closed segments")
    resolutions = [
        _resolve_approved_code_segment(candidate_code, reference_code, segment)
        for segment in segments
    ]
    for side in ("candidate", "reference"):
        spans: dict[str, list[tuple[int, int]]] = {}
        for resolution in resolutions:
            path = resolution[f"{side}_module_path"]
            start = resolution[f"{side}_start"]
            end = start + resolution["count"]
            spans.setdefault(path, []).append((start, end))
        for path, ranges in spans.items():
            ordered = sorted(ranges)
            if any(left[1] > right[0] for left, right in zip(ordered, ordered[1:])):
                raise ValueError(f"approved lineage {side} segments overlap in {path}")
    candidate_transformed: Mapping[str, Any] = candidate_code
    for index, resolution in sorted(
        enumerate(resolutions),
        key=lambda item: (item[1]["candidate_module_path"], item[1]["candidate_start"]),
        reverse=True,
    ):
        candidate_transformed = _replace_approved_run(
            candidate_transformed,
            resolution["candidate_module_path"],
            resolution["candidate_start"],
            resolution["count"],
            f"__NX_APPROVED_CANDIDATE_SEGMENT_{index}__",
        )
    reference_transformed: Mapping[str, Any] = reference_code
    for index, resolution in sorted(
        enumerate(resolutions),
        key=lambda item: (item[1]["reference_module_path"], item[1]["reference_start"]),
        reverse=True,
    ):
        reference_transformed = _replace_approved_run(
            reference_transformed,
            resolution["reference_module_path"],
            resolution["reference_start"],
            resolution["count"],
            f"~ NX APPROVED REFERENCE SEGMENT {index} ~",
        )
    evidence = {
        key: entry.get(key)
        for key in ("id", "axis", "reference_identity", "feature_id", "approval_scope")
    }
    evidence["status"] = "PASS"
    evidence["segments"] = [resolution["evidence"] for resolution in resolutions]
    return dict(candidate_transformed), dict(reference_transformed), evidence


def _code_compare(a: Mapping[str, Any], b: Mapping[str, Any]) -> tuple[str, str]:
    ma, mb = _axis(a, "code"), _axis(b, "code")
    if not ma.get("complete", bool(_values(a,"code"))) or not mb.get("complete", bool(_values(b,"code"))): return "INCOMPLETE", "code extraction incomplete"
    if set(ma.get("items", [])) & set(mb.get("items", [])): return "FAIL", "exact normalized module hash match"
    run = max(
        (
            match["count"]
            for left in ma.get("modules", [])
            for right in mb.get("modules", [])
            for match in _matching_runs(left, right, minimum=5)
        ),
        default=0,
    )
    if run >= 8: return "FAIL", f"{run} consecutive meaningful rows match"
    if run >= 5: return "WARN", f"{run} consecutive meaningful rows match"
    for pa in [p for m in ma.get("modules", []) for p in m.get("procedures", [])]:
        for pb in [p for m in mb.get("modules", []) for p in m.get("procedures", [])]:
            tokens_a, tokens_b = pa.get("tokens", []), pb.get("tokens", [])
            if pa.get("token_count", len(tokens_a)) < 30 or pb.get("token_count", len(tokens_b)) < 30: continue
            x,y=set(tokens_a),set(tokens_b); j=len(x&y)/len(x|y) if x|y else 0
            if j >= .70 and pa.get("structure") == pb.get("structure"): return "FAIL", f"procedure token Jaccard {j:.2f} with matching structure"
    return "PASS", "no code lineage match"

def _axis_compare(a: Mapping[str, Any], b: Mapping[str, Any], axis: str) -> tuple[str, str]:
    aa, bb = _axis(a, axis), _axis(b, axis)
    if axis in SEMANTIC_MIN_COUNTS and (len(_values(a, axis)) < SEMANTIC_MIN_COUNTS[axis] or len(_values(b, axis)) < SEMANTIC_MIN_COUNTS[axis]):
        return "INCOMPLETE", f"{axis} semantic measurement count below {SEMANTIC_MIN_COUNTS[axis]}"
    if not aa.get("complete", bool(_values(a, axis))) or not bb.get("complete", bool(_values(b, axis))): return "INCOMPLETE", f"{axis} extraction incomplete"
    if axis == "code": return _code_compare(a, b)
    if axis == "ui_wording":
        if aa.get("callbacks", []) == bb.get("callbacks", []) and aa.get("callbacks", []): return "FAIL", "exact reference-only callback sequence"
        if aa.get("labels", []) == bb.get("labels", []) and aa.get("labels", []): return "FAIL", "exact reference-only label sequence"
        common = set(aa.get("ngrams", [])) & set(bb.get("ngrams", []))
        return ("FAIL", "exact wording 5-gram match") if common else (("WARN", "partial wording 5-gram overlap") if set(_values(a,axis)) & set(_values(b,axis)) else ("PASS", "no UI lineage match"))
    left, right = Counter(_values(a, axis)), Counter(_values(b, axis))
    if left == right and left: return "FAIL", f"exact {axis} match"
    return ("WARN", f"partial {axis} overlap") if left & right else ("PASS", f"no {axis} lineage match")

def _provenance_status(candidate: Mapping[str, Any], reference: Mapping[str, Any]) -> tuple[str, str]:
    p = candidate.get("provenance", candidate)
    if not isinstance(p, Mapping): p = candidate
    required = ("origin", "spec_ids", "license", "source_tree_sha256", "artifact_sha256")
    missing = [x for x in required if not p.get(x)]
    if missing: return "INCOMPLETE", "missing provenance fields: " + ", ".join(missing)
    if not _HEX64.fullmatch(str(p.get("source_tree_sha256", ""))): return "INCOMPLETE", "candidate source tree digest is invalid"
    if p.get("origin") != "project-original" or p.get("license") != "Proprietary": return "FAIL", "candidate provenance is not project-original/Proprietary"
    rp = reference.get("provenance", reference)
    if not isinstance(rp, Mapping): rp = reference
    if not _HEX64.fullmatch(str(rp.get("source_tree_sha256", ""))): return "INCOMPLETE", "reference source tree digest is missing or invalid"
    if p.get("source_tree_sha256") == rp.get("source_tree_sha256") and p.get("source_tree_sha256"): return "FAIL", "source tree digest matches reference"
    return "PASS", "provenance fields and source digest are independent"

def compare_subjects(
    candidate: Mapping[str, Any],
    reference: Mapping[str, Any],
) -> dict[str, Any]:
    axes = {axis: dict(zip(("status", "reason"), _axis_compare(candidate, reference, axis))) for axis in LINEAGE_AXES}
    axes["provenance"] = dict(zip(("status", "reason"), _provenance_status(candidate, reference)))
    worst = max((x["status"] for x in axes.values()), key=lambda s: _RANK[s])
    return {"schema_version": 2, "axes": axes, "status": worst}


def _generic_ui_allowlist_terms(allowlist: Mapping[str, Any] | None) -> set[str]:
    if not isinstance(allowlist, Mapping):
        return set()
    return {
        str(entry["value"])
        for entry in allowlist.get("entries", [])
        if isinstance(entry, Mapping)
        and entry.get("axis") == "ui_wording"
        and entry.get("kind") == "exact-term"
        and isinstance(entry.get("value"), str)
        and entry["value"]
        and isinstance(entry.get("max_count"), int)
        and not isinstance(entry.get("max_count"), bool)
        and entry["max_count"] >= 1
    }


def _without_generic_ui_terms(subject: Mapping[str, Any], terms: set[str]) -> dict[str, Any]:
    transformed = deepcopy(subject)
    wording = transformed.get("ui_wording")
    if not terms or not isinstance(wording, Mapping):
        return transformed
    wording = dict(wording)
    for field in ("items", "labels"):
        values = wording.get(field)
        if isinstance(values, list):
            wording[field] = [value for value in values if str(value) not in terms]
    transformed["ui_wording"] = wording
    return transformed

def compare_many(
    candidate: Mapping[str, Any],
    references: Mapping[str, Mapping[str, Any]] | list[Mapping[str, Any]],
    approved_lineage_policy: Mapping[str, Any] | None = None,
    allowlist: Mapping[str, Any] | None = None,
) -> dict[str, Any]:
    refs = list(references.items()) if isinstance(references, Mapping) else [(str(i), x) for i, x in enumerate(references)]
    policy_errors = validate_approved_lineage_policy(approved_lineage_policy) if approved_lineage_policy is not None else []
    identity_error = set(name for name, _ in refs) != {"original", "v3x"}
    bindings: list[dict[str, Any]] = []
    binding_by_identity: dict[str, dict[str, Any]] = {}
    for name, ref in refs:
        prov = ref.get("provenance", {}) if isinstance(ref.get("provenance", {}), Mapping) else {}
        identity = prov.get("reference_identity", ref.get("reference_identity", name))
        inventory = prov.get("reference_root_inventory", ref.get("reference_root_inventory"))
        digest = inventory.get("inventory_sha256", inventory.get("root_sha256", "")) if isinstance(inventory, Mapping) else ""
        records = inventory.get("records") if isinstance(inventory, Mapping) else None
        inventory_ok = isinstance(records, list) and bool(records) and all(
            isinstance(row, Mapping)
            and set(row) == {"path", "sha256", "size"}
            and isinstance(row.get("path"), str)
            and row["path"]
            and "\\" not in row["path"]
            and not Path(row["path"]).is_absolute()
            and not any(part in {"", ".", ".."} for part in row["path"].split("/"))
            and _HEX64.fullmatch(str(row.get("sha256", "")))
            and isinstance(row.get("size"), int)
            and not isinstance(row.get("size"), bool)
            and row["size"] >= 0
            for row in records
        )
        if inventory_ok:
            paths = [row["path"] for row in records]
            inventory_ok = paths == sorted(paths, key=lambda value: value.encode("utf-8")) and len(set(paths)) == len(paths) and str(digest).lower() == _inventory_digest(records)
        if name not in {"original", "v3x"} or identity != name or not _HEX64.fullmatch(str(digest)) or not inventory_ok:
            identity_error = True
        binding = {"identity": name, "inventory_sha256": digest}
        bindings.append(binding)
        binding_by_identity[name] = binding

    approved_identity = None
    if approved_lineage_policy is not None and not policy_errors:
        approved_identity = approved_lineage_policy["approved_reference_identity"]
    comparison_refs = [
        (name, ref)
        for name, ref in refs
        if approved_identity is None or name != approved_identity
    ]
    allowed_ui_terms = _generic_ui_allowlist_terms(allowlist)
    results = {
        name: compare_subjects(
            _without_generic_ui_terms(candidate, allowed_ui_terms),
            _without_generic_ui_terms(ref, allowed_ui_terms),
        )
        for name, ref in comparison_refs
    }
    for name, row in results.items():
        row["reference_binding"] = dict(binding_by_identity[name])
    worst = max((v["status"] for v in results.values()), key=lambda s: _RANK[s], default="INCOMPLETE")
    aggregate = {}
    for axis in AXES:
        vals = [(name, row["axes"][axis]) for name, row in results.items()]
        name, item = max(vals, key=lambda pair: _RANK[pair[1]["status"]], default=("", {"status": "INCOMPLETE", "reason": "no references"}))
        aggregate[axis] = {"status": item["status"], "reason": item["reason"], "reference": name}
    detector = run_negative_detector(candidate, candidate)
    identity_keys = ("frozen_source_commit", "source_tree_sha256", "artifact_sha256")
    if detector.get("status") != "PASS": worst = "FAIL"
    if identity_error: worst = "INCOMPLETE"
    lineage_errors: list[str] = []
    approved_source = None
    if approved_lineage_policy is not None and not policy_errors:
        binding = binding_by_identity.get(str(approved_identity))
        if binding is None:
            lineage_errors.append("approved v3.4 reference binding is missing")
        else:
            approved_source = approved_lineage_source_contract(
                str(binding.get("inventory_sha256", "")), approved_lineage_policy
            )
            lineage_errors.extend(validate_approved_lineage_source(approved_source, approved_lineage_policy))
    if policy_errors or lineage_errors:
        worst = "FAIL"
    binding_payload = json.dumps(bindings, sort_keys=True, ensure_ascii=False)
    result = {"schema_version": 2, "candidate_identity": {k: candidate.get("provenance", {}).get(k, candidate.get(k, "")) if isinstance(candidate.get("provenance", {}), Mapping) else candidate.get(k, "") for k in identity_keys}, "axes": aggregate, "references": results, "reference_bindings": bindings, "reference_binding_sha256": _sha(binding_payload), "detector": detector, "status": worst}
    if approved_lineage_policy is not None:
        result["approved_lineage_policy_validation"] = {
            "status": "PASS" if not policy_errors else "FAIL",
            "errors": policy_errors,
        }
        if approved_source is not None:
            result["approved_lineage_source"] = approved_source
        result["approved_lineage_errors"] = lineage_errors
    if allowlist is not None:
        result["comparison_allowlist"] = {
            "status": "PASS",
            "ui_wording_terms": sorted(allowed_ui_terms, key=lambda value: value.encode("utf-8")),
        }
    return result

def check_allowlist(subject: Mapping[str, Any], allowlist: Mapping[str, Any]) -> dict[str, Any]:
    errors, counts = [], {}
    for entry in allowlist.get("entries", []):
        if not isinstance(entry, Mapping): errors.append("allowlist entry must be an object"); continue
        axis, value, maximum = entry.get("axis"), entry.get("value"), entry.get("max_count")
        if axis not in AXES or not isinstance(value, str) or isinstance(maximum, bool) or not isinstance(maximum, int) or maximum < 1: errors.append("invalid allowlist entry"); continue
        count = _values(subject, axis).count(value); counts[f"{axis}:{value}"] = count
        if count > maximum: errors.append(f"allowlist count exceeded: {axis}:{value}={count}>{maximum}")
    return {"status": "PASS" if not errors else "FAIL", "counts": counts, "matches": [{"axis": k.split(":", 1)[0], "value": k.split(":", 1)[1], "count": v} for k, v in counts.items()], "errors": errors}

def run_negative_detector(candidate: Mapping[str, Any], reference: Mapping[str, Any]) -> dict[str, Any]:
    detected = [axis for axis in LINEAGE_AXES if _values(candidate, axis) and Counter(_values(candidate, axis)) == Counter(_values(reference, axis))]
    incomplete = [axis for axis in LINEAGE_AXES if not _values(candidate, axis) or (axis in SEMANTIC_MIN_COUNTS and not _axis(candidate, axis).get("complete", bool(_values(candidate, axis))))]
    status = "INCOMPLETE" if incomplete else ("PASS" if set(detected) == set(LINEAGE_AXES) else "FAIL")
    return {"status": status, "detected_axes": detected, "required_axes": list(LINEAGE_AXES), "incomplete_axes": incomplete}

def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--candidate", type=Path, required=True)
    parser.add_argument("--reference", type=Path, nargs="+", required=True)
    parser.add_argument("--output", type=Path)
    parser.add_argument("--allowlist", type=Path)
    parser.add_argument(
        "--approved-lineage-policy",
        "--approved-lineage-exceptions",
        dest="approved_lineage_policy",
        type=Path,
        help="Bound first-party v3.4 lineage policy (legacy option name remains an alias).",
    )
    args = parser.parse_args(argv)
    candidate = subject_from_path(args.candidate, "candidate")
    refs = {p.name: subject_from_path(p, "reference") for p in args.reference}
    policy = None
    if args.approved_lineage_policy:
        policy_bytes = args.approved_lineage_policy.read_bytes()
        policy = json.loads(policy_bytes.decode("utf-8"))
    allowlist = json.loads(args.allowlist.read_text(encoding="utf-8")) if args.allowlist else None
    result = compare_many(candidate, refs, approved_lineage_policy=policy, allowlist=allowlist)
    if args.approved_lineage_policy:
        result["approved_lineage_policy_sha256"] = hashlib.sha256(policy_bytes).hexdigest()
    if args.allowlist:
        allow = check_allowlist(candidate, allowlist)
        result["allowlist"] = allow
        if allow["status"] != "PASS": result["status"] = "FAIL"
    data = json.dumps(result, ensure_ascii=False, indent=2)
    if args.output: args.output.write_text(data + "\n", encoding="utf-8")
    print(data)
    return 0 if result["status"] == "PASS" else 1

if __name__ == "__main__": raise SystemExit(main())
