#!/usr/bin/env python3
"""Verify profile artifacts from a real payload directory."""
from __future__ import annotations

import argparse
import hashlib
import json
import re
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
FORBIDDEN_TERMS = (b"kutools", b"detong", b"kteloader")
NETWORK_API_TERMS = (
    b"system.net.http",
    b"webclient",
    b"winhttprequest",
    b"xmlhttp",
    b"internetopen",
    b"urldownloadtofile",
)


def sha(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def contract_profile(profile: str) -> dict:
    data = json.loads((ROOT / "contracts" / "build-profile-contract.json").read_text(encoding="utf-8"))
    for item in data["profiles"]:
        if item["id"] == profile:
            return item
    raise ValueError(f"unknown build profile: {profile}")


def safe_rel(path: Path, root: Path) -> str:
    if path.is_symlink():
        raise ValueError(f"reparse/symlink payload entry: {path.name}")
    rel = path.relative_to(root).as_posix()
    if not rel or rel.startswith("../") or ".." in Path(rel).parts:
        raise ValueError(f"unsafe payload path: {rel}")
    return rel


def parse_sums(path: Path) -> dict[str, str]:
    records: dict[str, str] = {}
    previous = ""
    for line in path.read_text(encoding="utf-8-sig").splitlines():
        m = re.fullmatch(r"([0-9a-f]{64})  ([^/\\]+(?:/[^/\\]+)*)", line)
        if not m or m.group(2) == "SHA256SUMS" or (previous and m.group(2) <= previous):
            raise ValueError("malformed or unsorted SHA256SUMS")
        records[m.group(2)] = m.group(1)
        previous = m.group(2)
    return records


def inventory(profile: str, payload_root: Path) -> dict:
    spec = contract_profile(profile)
    root = payload_root.resolve()
    if not root.is_dir():
        raise ValueError(f"payload root is not a directory: {payload_root}")
    files = sorted((p for p in root.rglob("*") if p.is_file()), key=lambda p: p.relative_to(root).as_posix())
    rows = [{"path": safe_rel(p, root), "sha256": sha(p)} for p in files]
    names = {row["path"] for row in rows}
    expected = set(spec.get("artifact_files", [spec["host_artifact"], *spec.get("allowed_sidecars", [])]))
    metadata = set(spec.get("metadata_paths", ["SHA256SUMS", "docs/r105-build.md"]))
    if names - expected - metadata:
        raise ValueError(f"unexpected payload entries: {sorted(names - expected - metadata)}")
    if expected - names:
        raise ValueError(f"missing payload entries: {sorted(expected - names)}")
    if "SHA256SUMS" not in names:
        raise ValueError("SHA256SUMS missing")
    sums = parse_sums(root / "SHA256SUMS")
    sum_targets = names - {"SHA256SUMS"}
    if set(sums) != sum_targets:
        raise ValueError("SHA256SUMS does not cover exact payload and metadata set")
    by_name = {row["path"]: row["sha256"] for row in rows}
    if any(by_name[name] != digest for name, digest in sums.items()):
        raise ValueError("SHA256SUMS digest mismatch")
    forbidden: list[str] = []
    for row in rows:
        data = (root / row["path"]).read_bytes()
        lowered = row["path"].casefold().encode() + b"\0" + data.lower()
        if (
            any(term in lowered for term in FORBIDDEN_TERMS)
            or any(term in data.lower() for term in NETWORK_API_TERMS)
            or Path(row["path"]).suffix.casefold() in {".exe", ".pyd", ".zip"}
        ):
            forbidden.append(row["path"])
    if forbidden:
        raise ValueError(f"forbidden payload content: {sorted(set(forbidden))}")
    return {"schema_version": 2, "profile": profile, "final_name": spec.get("final_name", "내엑셀 v0.01_r105"), "source_check": False,
            "payload_files": [row for row in rows if row["path"] in expected],
            "artifact_files": [row for row in rows if row["path"] in expected],
            "metadata_files": [row for row in rows if row["path"] not in expected],
            "forbidden_matches": [], "requires_com_registration": bool(spec["requires_com_registration"]),
            "com_registration_performed": False, "kutools_included": False, "deployment_package_generated": False}


def source_report(profile: str) -> dict:
    """Check source policy without inventing artifact hashes."""
    spec = contract_profile(profile)
    manifest = ROOT / "build" / "manifests" / "Product.json"
    if not manifest.is_file():
        raise ValueError("Product manifest missing")
    source_files = [manifest, *sorted((ROOT / "src" / "dotnet" / "NxHost").glob("*.cs"))]
    source_rows = [{"path": p.relative_to(ROOT).as_posix(), "exists": p.is_file()} for p in source_files]
    if not all(row["exists"] for row in source_rows):
        raise ValueError("source policy references a missing source file")
    return {"schema_version": 2, "profile": profile, "final_name": spec.get("final_name", "내엑셀 v0.01_r105"), "source_check": True,
            "source_files": source_rows, "payload_files": [], "artifact_files": [], "metadata_files": [], "forbidden_matches": [],
            "source_policy": {"artifact_hash_bound": False, "metadata_paths": list(spec.get("metadata_paths", []))},
            "requires_com_registration": bool(spec["requires_com_registration"]), "com_registration_performed": False,
            "kutools_included": False, "artifact_hash_bound": False, "deployment_package_generated": False}


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--profile", required=True, choices=("internal-xlam", "enhanced-dll"))
    parser.add_argument("--payload-root", type=Path)
    parser.add_argument("--check-source", action="store_true")
    args = parser.parse_args()
    try:
        if args.payload_root is not None and args.check_source:
            parser.error("choose --payload-root or --check-source")
        report = source_report(args.profile) if args.check_source else inventory(args.profile, args.payload_root) if args.payload_root else parser.error("--payload-root or --check-source is required")
    except (OSError, ValueError) as exc:
        parser.error(str(exc))
    print(json.dumps(report, ensure_ascii=True, sort_keys=True, separators=(",", ":")))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
