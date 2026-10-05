"""Reviewed r64 to r65 active release-token migration; no published artifacts."""
import argparse
import json
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
        new = raw.replace(b"r64", b"r65") if relative in FULL else raw.replace(b"v0.01_r64", b"v0.01_r65").replace(b"docs/r64-build.md", b"docs/r65-build.md")
        if raw != new:
            changed.append(relative)
            if args.write:
                path.write_bytes(new)
    print(json.dumps({"write": args.write, "changed": changed}, indent=2))

if __name__ == "__main__":
    main()
