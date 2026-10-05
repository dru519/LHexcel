"""One-shot reviewed release-token migration, excluding historical fixtures/docs."""
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
FULL = '''build/Verify-NxEnhancedPackage.ps1
build/Uninstall-NxEnhanced.ps1
build/Register-NxHost.ps1
build/Install-NxEnhanced.ps1
tools/verify_build_profile.py
tools/build_r47_distribution.py
tools/build_r46_complete.py
tools/build_hwpx_embedded_resources.py
tools/build_exhaustive_route_catalog.py
tools/build_enhanced_distribution.py
contracts/bridge-contract.json
contracts/build-profile-contract.json
src/vba/ui/NxProductAbout.bas
src/vba/ui/FNxAbout.form.json
docs/dll-install-guide.md'''.splitlines()

def main():
    changed = []
    for relative in FULL:
        path = ROOT / relative
        raw = path.read_bytes()
        updated = raw.replace(b'r62', b'r63')
        if raw != updated:
            path.write_bytes(updated)
            changed.append(relative)
    # Keep the r62 removal fixtures and historical suite filenames unchanged.
    for path in [ROOT / 'tools/check_contracts.py', *sorted((ROOT / 'tests/python').glob('*.py'))]:
        raw = path.read_bytes()
        updated = raw.replace(b'v0.01_r62', b'v0.01_r63').replace(b'docs/r62-build.md', b'docs/r63-build.md')
        if path.name == 'test_r57_release_identity_contract.py':
            updated = updated.replace(b'r62', b'r63')
            # Native suite names describe when a workflow was introduced.
            updated = updated.replace(b'Run-R63', b'Run-R62')
        if raw != updated:
            path.write_bytes(updated)
            changed.append(path.relative_to(ROOT).as_posix())
    print('\n'.join(changed))

if __name__ == '__main__':
    main()
