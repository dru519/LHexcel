from __future__ import annotations

import json
import os
import subprocess
import tempfile
import unittest
from pathlib import Path


ROOT = Path(__file__).resolve().parents[2]
GATE = ROOT / "build" / "Artifact-Antivirus.ps1"
BUILD = ROOT / "build" / "Build-NxSetup.ps1"


class SetupAntivirusGateTests(unittest.TestCase):
    def test_gate_precedes_execution_and_publication(self):
        code = BUILD.read_text(encoding="utf-8-sig")
        first = code.index("Invoke-NxArtifactAntivirus")
        verify = code.index("Start-Process -FilePath $candidate")
        second = code.index("Invoke-NxArtifactAntivirus", first + 1)
        publish = code.index("[IO.File]::Copy($candidate, $OutputPath, $false)")
        self.assertLess(first, verify)
        self.assertLess(verify, second)
        self.assertLess(second, publish)
        self.assertIn("LOCAL_DEFENDER_PASS", code)
        self.assertIn("compiler_signature", code)
        self.assertIn("source_hashes", code)

    def test_no_security_configuration_changes_or_bypass(self):
        code = GATE.read_text(encoding="utf-8-sig")
        for forbidden in ("Set-MpPreference", "Add-MpPreference", "-Restore", "SkipAntivirus"):
            self.assertNotIn(forbidden, code)
        self.assertIn("-DisableRemediation", code)
        self.assertIn("Get-MpComputerStatus", code)
        self.assertIn("Get-MpThreatDetection", code)

    @unittest.skipUnless(os.name == "nt", "Windows PowerShell required")
    def test_fail_closed_decisions_without_live_threats(self):
        good = dict(
            error=None, scanner_trusted=True,
            antivirus_before=True, antivirus_after=True,
            realtime_before=True, realtime_after=True,
            signature_age_hours_before=1, signature_age_hours_after=1,
            scan_exit_code=0, scan_output="Scanning sample.exe found no threats.",
            sha256_before="a" * 64, sha256_after="a" * 64, new_threats=[],
        )
        cases = [({}, "LOCAL_DEFENDER_PASS")]
        for change, expected in (
            ({"error": "denied"}, "SCAN_ERROR"),
            ({"scanner_trusted": False}, "SCANNER_NOT_TRUSTED"),
            ({"antivirus_before": False}, "PROTECTION_NOT_ACTIVE"),
            ({"realtime_after": False}, "PROTECTION_NOT_ACTIVE"),
            ({"signature_age_hours_before": 49}, "SIGNATURES_NOT_CURRENT"),
            ({"signature_age_hours_after": -1}, "SIGNATURES_NOT_CURRENT"),
            ({"scan_exit_code": 2}, "SCAN_NOT_CLEAN"),
            ({"scan_exit_code": None}, "SCAN_NOT_CLEAN"),
            ({"scan_output": "Scan skipped"}, "CLEAN_RESULT_NOT_CONFIRMED"),
            ({"sha256_before": ""}, "ARTIFACT_CHANGED_OR_MISSING"),
            ({"sha256_after": None}, "ARTIFACT_CHANGED_OR_MISSING"),
            ({"sha256_after": "b" * 64}, "ARTIFACT_CHANGED_OR_MISSING"),
            ({"new_threats": [{"ThreatID": 1}]}, "THREAT_RECORDED"),
        ):
            cases.append((change, expected))
        fixtures = [{"evidence": dict(good, **change), "expected": expected}
                    for change, expected in cases]
        env = dict(os.environ, NX_AV_TEST_CASES=json.dumps(fixtures))
        escaped = str(GATE).replace("'", "''")
        command = (
            "$ErrorActionPreference='Stop';"
            f". '{escaped}';"
            "$cases=ConvertFrom-Json $env:NX_AV_TEST_CASES;"
            "foreach($case in $cases){"
            "$actual=Get-NxAntivirusVerdict -Evidence $case.evidence;"
            "if($actual -ne $case.expected){throw ($case.expected+' != '+$actual)}};"
            "Write-Output ('PASS '+$cases.Count)"
        )
        result = subprocess.run(
            ["powershell.exe", "-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass", "-Command", command],
            env=env, capture_output=True, text=True, encoding="utf-8", errors="replace",
        )
        self.assertEqual(0, result.returncode, result.stdout + result.stderr)
        self.assertIn("PASS 14", result.stdout)

    @unittest.skipUnless(os.name == "nt", "Windows PowerShell required")
    def test_missing_file_is_rejected_with_receipt(self):
        with tempfile.TemporaryDirectory() as folder:
            report = Path(folder) / "scan.json"
            env = dict(os.environ, NX_GATE=str(GATE), NX_SCAN_FILE=str(Path(folder) / "missing.exe"),
                       NX_SCAN_REPORT=str(report))
            command = (
                "$ErrorActionPreference='Stop';. $env:NX_GATE;"
                "try {Invoke-NxArtifactAntivirus -FilePath $env:NX_SCAN_FILE "
                "-ReportPath $env:NX_SCAN_REPORT;exit 2} "
                "catch {if($_.Exception.Message -notmatch 'Antivirus gate HOLD'){throw}};exit 0"
            )
            result = subprocess.run(
                ["powershell.exe", "-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass",
                 "-Command", command], env=env, capture_output=True, text=True,
                encoding="utf-8", errors="replace",
            )
            self.assertEqual(0, result.returncode, result.stdout + result.stderr)
            evidence = json.loads(report.read_text(encoding="utf-8-sig"))
            self.assertEqual("SCAN_ERROR", evidence["status"])
            self.assertEqual("Scan target is missing", evidence["error"])
            self.assertIsNone(evidence["scan_exit_code"])


if __name__ == "__main__":
    unittest.main()
