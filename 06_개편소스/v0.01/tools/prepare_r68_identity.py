"""Advance reviewed release identifiers to r68, retaining historical suite names."""
import argparse
import json
from pathlib import Path
from prepare_r64_identity import FULL

ROOT = Path(__file__).resolve().parents[1]

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--write', action='store_true')
    args = parser.parse_args()
    targets = {ROOT / p for p in FULL}
    targets.update((ROOT / 'tests/python').glob('*.py'))
    targets.update((ROOT / 'tests/windows').glob('*.ps1'))
    targets.add(ROOT / 'tools/check_contracts.py')
    changed = []
    for path in sorted(targets):
        raw = path.read_bytes()
        relative = path.relative_to(ROOT).as_posix()
        updated = raw.replace(b'r67', b'r68') if relative in FULL else raw.replace(b'v0.01_r67', b'v0.01_r68').replace(b'docs/r67-build.md', b'docs/r68-build.md')
        if raw != updated:
            changed.append(relative)
            if args.write:
                path.write_bytes(updated)
    print(json.dumps({'write': args.write, 'changed': changed}, indent=2))

if __name__ == '__main__':
    main()
