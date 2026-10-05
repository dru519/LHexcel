"""Advance active release tokens only; historical suite filenames remain intact."""
import argparse
import json
import subprocess
import sys
from pathlib import Path
from prepare_r64_identity import FULL

ROOT = Path(__file__).resolve().parents[1]
def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--write", action="store_true")
    args = parser.parse_args()
    targets = {ROOT / p for p in FULL}
    targets.update((ROOT / "tests/python").glob("*.py"))
    targets.update((ROOT / "tests/windows").glob("*.ps1"))
    targets.add(ROOT / "tools/check_contracts.py")
    changed = []
    for path in sorted(targets):
        raw = path.read_bytes()
        relative = path.relative_to(ROOT).as_posix()
        updated = raw.replace(b"r69", b"r70") if relative in FULL else raw.replace(b"v0.01_r69", b"v0.01_r70").replace(b"docs/r69-build.md", b"docs/r70-build.md")
        if updated != raw:
            changed.append(relative)
            if args.write:
                path.write_bytes(updated)
    if args.write:
        # These are generated contracts, not literal-token migration targets.
        for script, flags in (
            ("sync_nxhost_handshake.py", ["--write"]),
            ("build_exhaustive_route_catalog.py", ["--root", str(ROOT)]),
            ("build_hwpx_embedded_resources.py", ["--root", str(ROOT)]),
            ("build_product_manifest.py", ["--write"]),
            ("refresh_extension_manifests.py", []),
        ):
            subprocess.run([sys.executable, "-X", "utf8", str(ROOT / "tools" / script), *flags], cwd=ROOT, check=True)
    print(json.dumps({"write": args.write, "changed": changed}, indent=2))
if __name__ == "__main__":
    main()

