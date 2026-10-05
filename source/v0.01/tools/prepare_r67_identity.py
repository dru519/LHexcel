"""Reviewed r66-to-r67 release identity migration; historical workflow names stay stable."""
import argparse,json
from pathlib import Path
from prepare_r64_identity import FULL
ROOT=Path(__file__).resolve().parents[1]
def main():
 p=argparse.ArgumentParser();p.add_argument('--write',action='store_true');a=p.parse_args()
 targets={ROOT/x for x in FULL}
 targets.update((ROOT/'tests/python').glob('*.py'));targets.update((ROOT/'tests/windows').glob('*.ps1'))
 targets.add(ROOT/'tools/check_contracts.py')
 changed=[]
 for path in sorted(targets):
  raw=path.read_bytes();relative=path.relative_to(ROOT).as_posix()
  new=raw.replace(b'r66',b'r67') if relative in FULL else raw.replace(b'v0.01_r66',b'v0.01_r67').replace(b'docs/r66-build.md',b'docs/r67-build.md')
  if raw!=new:
   changed.append(relative)
   if a.write:path.write_bytes(new)
 print(json.dumps({'write':a.write,'changed':changed},indent=2))
if __name__=='__main__':main()

