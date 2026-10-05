import importlib.util
import sys
import unittest
from pathlib import Path
from types import SimpleNamespace

ROOT = Path(__file__).resolve().parents[4]
SCRIPT = ROOT / "03_개발자료/작업스크립트/lhexcel_windows_ascii_build.py"
spec = importlib.util.spec_from_file_location("public_runner_contract", SCRIPT)
runner = importlib.util.module_from_spec(spec)
sys.modules[spec.name] = runner
spec.loader.exec_module(runner)

class PublicRunnerTests(unittest.TestCase):
    def test_generated_bats_remain_ascii_under_korean_checkout(self):
        parent = Path("C:/검증 경로/내엑셀/workspace/tmp")
        paths = SimpleNamespace(tmp_dir=parent / "lhexcel_v001r105_direct_run_ascii_local01",
                                launcher_bat=parent / "run_local01.bat",
                                version="v0.01_r105", token="v001r105")
        for text in (runner.v001_release_run_bat_text(paths), runner.launcher_bat_text(paths)):
            text.encode("ascii")
            self.assertIn("%~dp0", text)
            self.assertNotIn("검증 경로", text)
    def test_output_is_confined_to_export_workspace(self):
        self.assertEqual(ROOT / "workspace/tmp", runner.TMP_ROOT)
