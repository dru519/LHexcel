"""Reviewed r63-to-r64 token migration. Historical reports and fixtures stay intact."""
import argparse
import json
from pathlib import Path
from prepare_r63_identity import FULL as PREVIOUS_SCOPE

ROOT=Path(__file__).resolve().parents[1]
FULL=set(PREVIOUS_SCOPE)|{'build/Build-NxSetup.ps1',
 'tests/python/test_r57_release_identity_contract.py','tests/python/test_nxhost_installer_contract.py',
 'tests/windows/Run-NxHostVisualSession.ps1','tests/windows/Run-R63SetupSuite.ps1'}
def main():
    parser=argparse.ArgumentParser()
    parser.add_argument('--write',action='store_true')
    args=parser.parse_args()
    targets={ROOT/p for p in FULL}
    targets.update((ROOT/'tests/python').glob('*.py'))
    targets.update((ROOT/'tests/windows').glob('*.ps1'))
    targets.add(ROOT/'tools/check_contracts.py')
    changes=[]
    for path in sorted(targets):
        raw=path.read_bytes();rel=path.relative_to(ROOT).as_posix()
        updated=raw.replace(b'r63',b'r64') if rel in FULL else raw.replace(b'v0.01_r63',b'v0.01_r64').replace(b'docs/r63-build.md',b'docs/r64-build.md')
        if rel=='tests/python/test_r57_release_identity_contract.py':
            updated=updated.replace(b'Run-R64',b'Run-R63')
        if updated!=raw:
            changes.append(rel)
            if args.write:path.write_bytes(updated)
    print(json.dumps({'mode':'write' if args.write else 'preview','paths':changes},indent=2))
if __name__=='__main__':main()
