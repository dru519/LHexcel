import json
import unittest
from pathlib import Path


ROOT = Path(__file__).resolve().parents[2]


class R47DistributionPolicyContractTests(unittest.TestCase):
    def test_r47_distribution_builder_exists(self):
        self.assertTrue((ROOT / "tools/build_r47_distribution.py").exists())

    def test_public_policy_keeps_expiry_without_personal_notice(self):
        script = (ROOT / "tools/build_r47_distribution.py").read_text(encoding="utf-8")
        for token in ("2027-12-31", "2028-01-01", "LHE_BIRTHDAY_MONTH As Long = 0", "LHE_BIRTHDAY_DAY As Long = 0"):
            self.assertIn(token, script)
        self.assertIn('"BirthdayNoticeEnabled": "false"', script)

    def test_distribution_policy_uses_only_v001_runtime_symbols(self):
        script = (ROOT / "tools/build_r47_distribution.py").read_text(encoding="utf-8")
        for token in ("NxLHexcelProfileRoot", "NxShortcutsApplySavedBindings", "distribution-notice-v1.cfg"):
            self.assertIn(token, script)
        for legacy_only in (
            "LHExcelCfgPath", "GetPrivateProfileString", "WritePrivateProfileString",
            "LHExcelCloseAndRemoveLegacyXLStartCfg", "EnsureLHExcelCfg",
            "ApplyShortcutSettings", "LHExcelEnsureVisibleWorkbook",
        ):
            self.assertNotIn(legacy_only, script)

    def test_distribution_contract_is_single_xlam(self):
        contract = json.loads((ROOT / "contracts/build-profile-contract.json").read_text(encoding="utf-8"))
        profiles = contract["profiles"]
        internal = next(row for row in profiles if row["id"] == "internal-xlam")
        self.assertEqual(["Product.xlam"], internal["artifact_files"])


if __name__ == "__main__":
    unittest.main()
