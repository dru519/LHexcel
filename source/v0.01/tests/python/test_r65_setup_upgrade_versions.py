"""Exercise generated BAT arguments; no installer or Excel is executed."""
import contextlib
import importlib.util
import io
import json
from pathlib import Path
import sys
import tempfile
import unittest
from unittest.mock import patch

ROOT = Path(__file__).resolve().parents[2]
RUNNER = ROOT.parents[1] / 'development/scripts/lhexcel_windows_ascii_build.py'
spec = importlib.util.spec_from_file_location('r65_setup_runner', RUNNER)
runner = importlib.util.module_from_spec(spec)
sys.modules[spec.name] = runner
spec.loader.exec_module(runner)


class R65SetupUpgradeVersions(unittest.TestCase):
    def test_manual_visual_runner_never_activates_uia_harness(self):
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            product = root / 'fixture.xlam'
            product.write_bytes(b'not executed')
            for name in ('NxHost32.dll', 'NxHost64.dll', 'NxCore32.dll', 'NxCore64.dll'):
                (root / name).write_bytes(b'not executed')
            paths = runner.V001ReleasePaths('v0.01_r101', 'v001r65', ROOT, root, root / 'run.bat')
            with contextlib.redirect_stdout(io.StringIO()):
                runner.prepare_v001_verification(paths, 'manual-ui', 'dll-navigator', 'manual', product, root)
            script = (root / 'verify-manual-ui.ps1').read_text(encoding='ascii')
            self.assertIn('Run-NxHostVisualSession.ps1', script)
            self.assertNotIn('-Automated', script)
            self.assertNotIn('-Enhancements', script)
            self.assertNotIn('-DocumentOnly', script)
            self.assertIn('exact-host-manual-ui', script)

    def generate(self, previous=None):
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            setup = root / 'fixture.exe'
            setup.write_bytes(b'not executed')
            args = ['setup-acceptance', '--setup-exe', str(setup), '--tag', 'versions', '--version', 'v0.01_r101']
            if previous:
                args += ['--previous-setup-exe', str(setup), '--previous-version', previous]
            with patch.object(runner, 'TMP_ROOT', root / 'runners'), contextlib.redirect_stdout(io.StringIO()) as output:
                self.assertEqual(runner.main(args), 0)
            return Path(json.loads(output.getvalue())['bat']).read_text(encoding='ascii')

    def test_r64_to_r65_are_explicit_in_bat(self):
        bat = self.generate('v0.01_r64')
        self.assertIn('-PreviousVersion "v0.01_r64"', bat)
        self.assertIn('-ExpectedVersion "v0.01_r101"', bat)

    def test_clean_install_has_explicit_target(self):
        bat = self.generate()
        self.assertIn('-ExpectedVersion "v0.01_r101"', bat)
        self.assertNotIn('-PreviousVersion', bat)

    def test_older_source_version_is_not_silently_replaced(self):
        self.assertIn('-PreviousVersion "v0.01_r63"', self.generate('v0.01_r63'))

    def test_invalid_version_rejected_before_staging(self):
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            with patch.object(runner, 'TMP_ROOT', root / 'runners'), contextlib.redirect_stderr(io.StringIO()):
                with self.assertRaises(SystemExit) as error:
                    runner.main(['setup-acceptance', '--setup-exe', 'not-read.exe', '--tag', 'invalid', '--version', 'v0.01_r101" & whoami'])
            self.assertEqual(error.exception.code, 2)
            self.assertFalse((root / 'runners').exists())

    def test_suite_uses_version_for_paths_and_child_acceptance(self):
        code = (ROOT / 'tests/windows/Run-NxSetupUpgradeSuite.ps1').read_text(encoding='utf-8-sig')
        self.assertIn('-ExpectedVersion $PreviousVersion', code)
        self.assertIn('-ExpectedVersion $ExpectedVersion', code)
        self.assertIn("+' '+$PreviousVersion+'.xlam'", code)
        self.assertNotIn("-ExpectedVersion v0.01_r63", code)

    def test_installed_acceptance_covers_new_services(self):
        code = (ROOT / 'tests/windows/Run-R63SetupSuite.ps1').read_text(encoding='utf-8-sig')
        identities = json.loads((ROOT / 'contracts/nxhost-contract.json').read_text(encoding='utf-8'))['com_identities']
        for name in ('picture_preview', 'workbook_compare'):
            self.assertIn(identities[name]['prog_id'], code)
            self.assertIn(identities[name]['class_id'], code)


if __name__ == '__main__':
    unittest.main()
