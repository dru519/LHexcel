#!/usr/bin/env python3
"""Build one final r105 profile into a staging directory and verify it."""
from __future__ import annotations

import argparse
import hashlib
import json
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

from windows_powershell import windows_powershell_environment

ROOT = Path(__file__).resolve().parents[1]
VERIFIER = ROOT / "tools" / "verify_build_profile.py"


def run_ps(script: Path, args: list[str]) -> None:
    command = ["powershell", "-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass", "-File", str(script), *args]
    # Windows PowerShell may emit the shared UNC path as UTF-8 while the
    # parent process still advertises the legacy CP949 code page.  Capture
    # bytes and decode only the error path so a successful build cannot fail
    # in the reader thread before the artifact is published.
    completed = subprocess.run(
        command,
        cwd=ROOT,
        env=windows_powershell_environment(),
        text=False,
        capture_output=True,
    )
    if completed.returncode:
        raw = completed.stderr or completed.stdout or b"PowerShell build failed"
        raise RuntimeError(raw.decode("utf-8", errors="replace").strip())


def write_sums(root: Path) -> None:
    rows = []
    for path in sorted((p for p in root.rglob("*") if p.is_file() and p.name != "SHA256SUMS"), key=lambda p: p.relative_to(root).as_posix()):
        rows.append(f"{hashlib.sha256(path.read_bytes()).hexdigest()}  {path.relative_to(root).as_posix()}")
    (root / "SHA256SUMS").write_text("\n".join(rows) + "\n", encoding="utf-8", newline="\n")


def main() -> int:
    parser = argparse.ArgumentParser(description="내엑셀 v0.01_r105 final-only build")
    parser.add_argument("--profile", choices=("internal-xlam", "enhanced-dll"), default="enhanced-dll")
    parser.add_argument("--output-root", type=Path,
                        help="empty caller-owned staging/output directory; no deployment is performed")
    parser.add_argument("--deployment-package", action="store_true", help=argparse.SUPPRESS)
    args = parser.parse_args()
    if args.deployment_package:
        parser.error("deployment package disabled/rejected for final-only r105 builds")
    if args.output_root is None:
        parser.error("--output-root is required; refusing implicit repository output")
    output = args.output_root.resolve()
    if output.exists() and any(output.iterdir()):
        parser.error(f"refusing non-empty output root: {output}")
    output.mkdir(parents=True, exist_ok=True)
    staging = Path(tempfile.mkdtemp(prefix="naeexcel-r105-"))
    try:
        payload = staging / "payload"
        payload.mkdir()
        staged_xlam = payload / "Product.xlam"
        run_ps(ROOT / "build" / "Build-Xlam.ps1", ["-DataRoot", str(ROOT), "-OutputPath", str(staged_xlam), "-ManifestPath", str(ROOT / "build" / "manifests" / "Product.json"), "-RibbonPath", str(ROOT / "src" / "ribbon" / "customUI14.xml")])
        if not staged_xlam.is_file():
            raise RuntimeError("Build-Xlam did not produce Product.xlam")
        if args.profile == "enhanced-dll":
            host_stage = staging / "NxHost"
            run_ps(ROOT / "build" / "Build-NxHost.ps1", ["-OutputRoot", str(host_stage), "-Configuration", "Release", "-ProtectedLoader"])
            for source, target in ((host_stage / arch / name, payload / name) for arch, suffix in (("x86", "32"), ("x64", "64")) for name in (f"NxHost{suffix}.dll", f"NxCore{suffix}.dll")):
                if not source.is_file():
                    raise RuntimeError(f"Build-NxHost missing {source.name}")
                shutil.copy2(source, target)
        docs = payload / "docs"
        docs.mkdir()
        profiles = json.loads((ROOT / "contracts/build-profile-contract.json").read_text(encoding="utf-8"))
        spec = next(item for item in profiles["profiles"] if item["id"] == args.profile)
        build_docs = [entry for entry in spec["metadata_paths"] if entry.startswith("docs/")]
        if len(build_docs) != 1 or Path(build_docs[0]).parts[0] != "docs" or ".." in Path(build_docs[0]).parts:
            raise ValueError("Invalid build document in profile")
        shutil.copy2(ROOT / build_docs[0], payload / build_docs[0])
        write_sums(payload)
        check = subprocess.run([sys.executable, str(VERIFIER), "--profile", args.profile, "--payload-root", str(payload)], cwd=ROOT, text=True, capture_output=True)
        if check.returncode:
            raise RuntimeError(check.stderr or check.stdout)
        report = json.loads(check.stdout)
        # Publish atomically at the directory-content level only after verification.
        for source in payload.iterdir():
            target = output / source.name
            if source.is_dir():
                shutil.copytree(source, target)
            else:
                shutil.copy2(source, target)
        report["build_mode"] = "final-only"
        print(json.dumps(report, ensure_ascii=True, sort_keys=True, separators=(",", ":")))
        return 0
    except (OSError, RuntimeError) as exc:
        parser.error(str(exc))
    finally:
        shutil.rmtree(staging, ignore_errors=True)


if __name__ == "__main__":
    raise SystemExit(main())
