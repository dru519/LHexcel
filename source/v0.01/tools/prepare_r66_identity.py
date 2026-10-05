"""Reviewed active release-token migration; published artifacts are untouched."""
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
        new = raw.replace(b"r65", b"r66") if relative in FULL else raw.replace(b"v0.01_r65", b"v0.01_r66").replace(b"docs/r65-build.md", b"docs/r66-build.md")
        # Suite names are stable workflow identities, not product versions.
        new = new.replace(b"Run-R66", b"Run-R65") if relative in FULL else new
        if raw != new:
            changed.append(relative)
            if args.write: path.write_bytes(new)
    print(json.dumps({"write": args.write, "changed": changed}, indent=2))

if __name__ == "__main__": main()
