#!/usr/bin/env python3
"""Build and verify the deterministic Product XLAM input inventory."""

from __future__ import annotations

import argparse
import hashlib
import json
import os
import tempfile
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
OUTPUT = ROOT / "build" / "manifests" / "Product.json"
RIBBON = ROOT / "src" / "ribbon" / "customUI14.xml"
VBA_ROOT = ROOT / "src" / "vba"
FEATURE_CONTRACT = ROOT / "contracts" / "feature-contract.json"
R46_CONTRACTS = {
    "bridge": ROOT / "contracts" / "bridge-contract.json",
    "build_profiles": ROOT / "contracts" / "build-profile-contract.json",
    "commands": ROOT / "contracts" / "command-contract.json",
}
DIAGNOSTIC_ONLY_FORMS = {
    "FNxStart",
    "FNxAllFunctions",
    "FNxData",
    "FNxDraw",
    "FNxFile",
    "FNxTemplate",
    "FNxFocusPalette",
    "FNxFocusOverlay",
}
DIAGNOSTIC_ONLY_SOURCE_PATHS = {
    *(f"src/vba/ui/{name}.form.json" for name in {"FNxStart", "FNxAllFunctions"}),
    *(f"src/vba/ui/{name}.vba" for name in {"FNxStart", "FNxAllFunctions"}),
    *(f"src/vba/features/data/ui/{name}.form.json" for name in {"FNxData", "FNxFocusPalette"}),
    *(f"src/vba/features/data/ui/{name}.vba" for name in {"FNxData", "FNxFocusPalette"}),
    "src/vba/features/data/focus/NxFocusOverlay.bas",
    "src/vba/features/data/focus/NxFocusGeometry.bas",
    "src/vba/features/data/focus/CNxFocusGeometry.cls",
    "src/vba/features/data/ui/FNxFocusOverlay.form.json",
    "src/vba/features/data/ui/FNxFocusOverlay.vba",
    "src/vba/features/draw/ui/FNxDraw.form.json",
    "src/vba/features/draw/ui/FNxDraw.vba",
    "src/vba/features/file/ui/FNxFile.form.json",
    "src/vba/features/file/ui/FNxFile.vba",
    "src/vba/features/template/ui/FNxTemplate.form.json",
    "src/vba/features/template/ui/FNxTemplate.vba",
}
HANGUL_TABLE_SEND_REQUIREMENT = {
    "feature_id": "NX-HANGUL-TABLE-SEND",
    "name": "아래한글 표 전송",
    "reference": "v3.4",
    "status": "implemented",
    "transport": "embedded-HWPX",
}
SOURCE_SUFFIXES = {".bas", ".cls", ".vba"}


def sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def relative(path: Path) -> str:
    return path.relative_to(ROOT).as_posix()


def source_record(path: Path) -> dict[str, str]:
    rel = relative(path)
    if rel.endswith(".form.json"):
        kind = "form_layout"
    elif path.suffix.lower() == ".vba":
        kind = "form_code"
    elif path.suffix.lower() == ".cls":
        kind = "class_module"
    else:
        kind = "standard_module"
    record: dict[str, str] = {
        "path": rel,
        "kind": kind,
        "owner_phase": "product",
        "sha256": sha256(path),
    }
    if kind == "form_layout":
        document = json.loads(path.read_text(encoding="utf-8"))
        code_path = document.get("code_path")
        if not isinstance(code_path, str) or not code_path:
            raise ValueError(f"form layout missing code_path: {rel}")
        record["code_path"] = code_path.replace("\\", "/")
    return record


def build_manifest() -> dict[str, object]:
    if not VBA_ROOT.is_dir():
        raise ValueError(f"Product VBA source root missing: {VBA_ROOT}")
    paths = sorted(
        (
            path
            for path in VBA_ROOT.rglob("*")
            if path.is_file()
            and (path.suffix.lower() in SOURCE_SUFFIXES or path.name.endswith(".form.json"))
            and relative(path) not in DIAGNOSTIC_ONLY_SOURCE_PATHS
        ),
        key=lambda path: relative(path),
    )
    if not paths:
        raise ValueError("Product VBA inventory is empty")
    records = [source_record(path) for path in paths]
    seen = [record["path"] for record in records]
    if len(seen) != len(set(seen)):
        raise ValueError("duplicate Product source path")

    layouts = [record for record in records if record["kind"] == "form_layout"]
    codes = {record["path"] for record in records if record["kind"] == "form_code"}
    bindings = []
    for layout in layouts:
        code = layout.get("code_path")
        if code not in codes:
            raise ValueError(f"form code missing: {code}")
        bindings.append({"layout": layout["path"], "code": code})
    if len({(item["layout"], item["code"]) for item in bindings}) != len(bindings):
        raise ValueError("duplicate Product form binding")
    if len(bindings) != len(codes):
        raise ValueError("Product form bindings must be a bijection")
    if not RIBBON.is_file():
        raise ValueError(f"Product ribbon source missing: {relative(RIBBON)}")
    if not FEATURE_CONTRACT.is_file():
        raise ValueError("Product feature contract missing")
    missing_r46_contracts = [name for name, path in R46_CONTRACTS.items() if not path.is_file()]
    if missing_r46_contracts:
        raise ValueError(f"Product r46 contracts missing: {missing_r46_contracts}")
    feature_contract = json.loads(FEATURE_CONTRACT.read_text(encoding="utf-8"))
    build_profiles = json.loads(R46_CONTRACTS["build_profiles"].read_text(encoding="utf-8"))
    if build_profiles.get("default_profile") != "enhanced-dll":
        raise ValueError("Product default build profile rejected")
    lineage = feature_contract.get("lineage", {})
    catalog_ids = {feature.get("id") for feature in feature_contract.get("features", [])}
    release_ids = lineage.get("release_feature_ids")
    if not isinstance(release_ids, list) or not release_ids or len(release_ids) != len(set(release_ids)):
        raise ValueError("Product release feature inventory rejected")
    if set(release_ids) != catalog_ids:
        raise ValueError("Product release features must exactly cover the catalog")
    if lineage.get("hangul_table_send_requirement") != HANGUL_TABLE_SEND_REQUIREMENT:
        raise ValueError("Product direct Hangul table-send requirement rejected")

    return {
        "schema_version": 3,
        "suite": "Product",
        "owner_phase": "product",
        "default_build_profile": build_profiles["default_profile"],
        "r46_contracts": {
            name: {"path": relative(path), "sha256": sha256(path)}
            for name, path in sorted(R46_CONTRACTS.items())
        },
        "modules": records,
        "form_bindings": bindings,
        "ribbon": {
            "path": relative(RIBBON),
            "sha256": sha256(RIBBON),
            "package_part": "customUI/customUI14.xml",
            "relationship_type": "http://schemas.microsoft.com/office/2007/relationships/ui/extensibility",
        },
        "release_scope": {
            "feature_contract": {"path": relative(FEATURE_CONTRACT), "sha256": sha256(FEATURE_CONTRACT)},
            "feature_count": len(release_ids),
            "included_feature_ids": release_ids,
            "excluded_feature_ids": [],
        },
        "diagnostic_only": {
            "forms": sorted(DIAGNOSTIC_ONLY_FORMS),
            "source_paths": sorted(DIAGNOSTIC_ONLY_SOURCE_PATHS),
        },
        "forbidden_prefixes": ["tests/", "tests\\"],
    }


def render(document: dict[str, object]) -> bytes:
    return (json.dumps(document, ensure_ascii=False, indent=2) + "\n").encode("utf-8")


def atomic_write(path: Path, content: bytes) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    fd, temporary = tempfile.mkstemp(prefix=f".{path.name}.", dir=path.parent)
    try:
        with os.fdopen(fd, "wb") as handle:
            handle.write(content)
            handle.flush()
            os.fsync(handle.fileno())
        os.replace(temporary, path)
    finally:
        if os.path.exists(temporary):
            os.unlink(temporary)


def main() -> int:
    parser = argparse.ArgumentParser()
    mode = parser.add_mutually_exclusive_group(required=True)
    mode.add_argument("--write", action="store_true")
    mode.add_argument("--check", action="store_true")
    args = parser.parse_args()
    expected = render(build_manifest())
    if args.check:
        if not OUTPUT.is_file() or OUTPUT.read_bytes() != expected:
            raise SystemExit(f"stale Product manifest: {OUTPUT}")
        return 0
    atomic_write(OUTPUT, expected)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
