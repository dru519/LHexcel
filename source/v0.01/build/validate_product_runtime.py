"""Validate the Product embedded HWPX resource trust boundary."""

from __future__ import annotations

import argparse
import hashlib
import json
import subprocess
import sys
from pathlib import Path


RESOURCE_PATHS = ["tools/lhexcel_hwpx_table_export.ps1", "templates/표.hwpx"]
FORBIDDEN_SUFFIXES = {".exe", ".dll", ".pyd"}


def _sha(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def _contained_file(root: Path, relative: str) -> Path:
    if "\\" in relative or relative.startswith("/") or any(part in {"", ".", ".."} for part in relative.split("/")):
        raise ValueError(f"unsafe runtime resource path: {relative}")
    path = root / relative
    resolved_root = root.resolve(strict=True)
    resolved = path.resolve(strict=True)
    if resolved_root not in resolved.parents or not resolved.is_file() or path.is_symlink():
        raise ValueError(f"runtime resource escaped source root: {relative}")
    return path


def validate_runtime(data_root: Path, manifest: dict[str, object]) -> None:
    contract_path = _contained_file(data_root, "contracts/hwpx-runtime-contract.json")
    contract = json.loads(contract_path.read_text(encoding="utf-8"))
    runtime = contract.get("runtime")
    deployment = contract.get("deployment")
    if not isinstance(runtime, dict) or runtime.get("launcher") != "embedded-powershell-5.1":
        raise ValueError("Product HWPX embedded launcher contract rejected")
    if not isinstance(deployment, dict) or deployment.get("single_file") is not True:
        raise ValueError("Product HWPX single-file deployment contract rejected")
    if set(deployment.get("forbidden_extensions", [])) != FORBIDDEN_SUFFIXES:
        raise ValueError("Product forbidden runtime extensions rejected")
    resources = runtime.get("resources")
    if not isinstance(resources, list) or [row.get("path") for row in resources if isinstance(row, dict)] != RESOURCE_PATHS:
        raise ValueError("Product embedded resource inventory rejected")
    for row in resources:
        path = _contained_file(data_root, row["path"])
        if path.suffix.lower() in FORBIDDEN_SUFFIXES or _sha(path) != row.get("sha256"):
            raise ValueError(f"Product embedded resource digest rejected: {row['path']}")

    declared_runtime = manifest.get("runtime")
    if not isinstance(declared_runtime, dict):
        raise ValueError("Product manifest runtime missing")
    if declared_runtime.get("launcher") != "embedded-powershell-5.1" or declared_runtime.get("single_file") is not True:
        raise ValueError("Product manifest runtime identity rejected")
    if declared_runtime.get("contract") != {"path": "contracts/hwpx-runtime-contract.json", "sha256": _sha(contract_path)}:
        raise ValueError("Product manifest contract digest rejected")
    expected = [{"path": relative, "sha256": _sha(_contained_file(data_root, relative))} for relative in RESOURCE_PATHS]
    if declared_runtime.get("resources") != expected:
        raise ValueError("Product manifest embedded resource digests rejected")

    generator = _contained_file(data_root, "tools/build_hwpx_embedded_resources.py")
    completed = subprocess.run([sys.executable, str(generator), "--root", str(data_root), "--check"], capture_output=True, text=True)
    if completed.returncode != 0:
        raise ValueError(completed.stderr.strip() or completed.stdout.strip() or "generated embedded resource module is stale")


def validate_form_bindings(manifest: dict[str, object]) -> None:
    modules = manifest.get("modules")
    bindings = manifest.get("form_bindings")
    if not isinstance(modules, list) or not isinstance(bindings, list) or not bindings:
        raise ValueError("Product form binding inventory rejected")
    layouts = {row.get("path"): row for row in modules if row.get("kind") == "form_layout"}
    codes = {row.get("path"): row for row in modules if row.get("kind") == "form_code"}
    pairs = [(row.get("layout"), row.get("code")) for row in bindings]
    if len(pairs) != len(set(pairs)) or {layout for layout, _ in pairs} != set(layouts) or {code for _, code in pairs} != set(codes):
        raise ValueError("Product form binding bijection rejected")
    for layout, code in pairs:
        if layouts[layout].get("code_path") != code:
            raise ValueError("Product form layout/code binding rejected")


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--data-root", type=Path, required=True)
    parser.add_argument("--manifest", type=Path, required=True)
    args = parser.parse_args()
    data_root = args.data_root.resolve(strict=True)
    manifest = json.loads(args.manifest.read_text(encoding="utf-8"))
    validate_runtime(data_root, manifest)
    validate_form_bindings(manifest)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
