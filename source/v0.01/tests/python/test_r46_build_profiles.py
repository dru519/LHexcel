from __future__ import annotations

import hashlib
import importlib.util
import json
import os
import subprocess
import sys
import unittest
import tempfile
from pathlib import Path


ROOT = Path(__file__).resolve().parents[2]
SOURCE_ROOT = ROOT.parent
BUILD_PROFILES = ROOT / "contracts" / "build-profile-contract.json"
V34_MAPPING = ROOT / "provenance" / "v34-r46-mapping.json"
PROFILE_VERIFIER = ROOT / "tools" / "verify_build_profile.py"
COMPLETE_BUILD = ROOT / "tools" / "build_r46_complete.py"
ENHANCED_BUILD = ROOT / "tools" / "build_enhanced_distribution.py"
WINDOWS_POWERSHELL_ENV = ROOT / "tools" / "windows_powershell.py"

EXPECTED_PAYLOADS = {
    "internal-xlam": {"Product.xlam"},
    "enhanced-dll": {"Product.xlam", "NxHost32.dll", "NxHost64.dll", "NxCore32.dll", "NxCore64.dll"},
}
FORBIDDEN_PAYLOAD_SUFFIXES = {".exe", ".pyd", ".zip"}


def load_json(path: Path) -> dict:
    return json.loads(path.read_text(encoding="utf-8"))


def run_profile_check(profile: str) -> dict:
    if not PROFILE_VERIFIER.is_file():
        raise AssertionError(
            "missing planned build-profile source verifier: tools/verify_build_profile.py"
        )
    completed = subprocess.run(
        [sys.executable, str(PROFILE_VERIFIER), "--profile", profile, "--check-source"],
        cwd=ROOT,
        capture_output=True,
        text=True,
        encoding="utf-8",
        check=False,
    )
    if completed.returncode != 0:
        raise AssertionError((completed.stdout or "") + (completed.stderr or ""))
    return json.loads(completed.stdout)


class R46BuildProfileTests(unittest.TestCase):
    @unittest.skipUnless(os.name == "nt", "Windows PSModulePath semantics: NOT_RUN on this host")
    def test_python_launches_windows_powershell_without_core_module_shadowing(self) -> None:
        self.assertTrue(WINDOWS_POWERSHELL_ENV.is_file())
        spec = importlib.util.spec_from_file_location("windows_powershell", WINDOWS_POWERSHELL_ENV)
        self.assertIsNotNone(spec)
        self.assertIsNotNone(spec.loader)
        module = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(module)
        source = {
            "SystemRoot": r"C:\Windows",
            "PSModulePath": os.pathsep.join((
                r"C:\Program Files\PowerShell\Modules",
                r"C:\runtime\native\powershell\Modules",
                r"C:\Users\tester\Documents\WindowsPowerShell\Modules",
                r"C:\Program Files\WindowsPowerShell\Modules",
            )),
        }
        sanitized = module.windows_powershell_environment(source)
        parts = sanitized["PSModulePath"].split(os.pathsep)
        self.assertTrue(parts)
        self.assertTrue(all("\\windowspowershell\\" in part.casefold() for part in parts))
        self.assertFalse(any(r"\native\powershell\modules" in part.casefold() for part in parts))
        for builder in (COMPLETE_BUILD, ENHANCED_BUILD):
            self.assertIn(
                "windows_powershell_environment()",
                builder.read_text(encoding="utf-8"),
                builder.name,
            )

    def test_profile_verifier_exists_as_a_source_check_tool(self) -> None:
        self.assertTrue(
            PROFILE_VERIFIER.is_file(),
            "missing planned build-profile source verifier: tools/verify_build_profile.py",
        )

    def test_internal_xlam_source_check_has_only_xlam_docs_and_hashes_without_com_registration(self) -> None:
        report = run_profile_check("internal-xlam")
        payload = report["artifact_files"]
        payload_names = {item["path"] for item in payload}

        self.assertEqual(set(), payload_names)
        self.assertEqual([], report["metadata_files"])
        self.assertFalse(report["artifact_hash_bound"])
        self.assertFalse(report["requires_com_registration"])
        self.assertFalse(report["com_registration_performed"])
        self.assertTrue(all(name == "Product.xlam" or name == "SHA256SUMS" or name.startswith("docs/") for name in payload_names))

    def test_enhanced_dll_source_check_has_exact_host_and_both_dlls(self) -> None:
        report = run_profile_check("enhanced-dll")
        payload_names = {item["path"] for item in report["artifact_files"]}

        self.assertEqual(set(), payload_names)
        self.assertEqual([], report["metadata_files"])
        self.assertTrue(report["requires_com_registration"])
        self.assertTrue(all(name == "SHA256SUMS" or name.startswith("docs/") or name in EXPECTED_PAYLOADS["enhanced-dll"] for name in payload_names))

    def test_profile_verifier_binds_a_sha256_to_every_payload_dll(self) -> None:
        report = run_profile_check("enhanced-dll")
        hashes = {item["path"]: item["sha256"] for item in report["artifact_files"]}

        self.assertEqual({}, hashes)

    def test_profile_verifier_rejects_executables_python_extensions_and_zip_payloads(self) -> None:
        reports = [run_profile_check(profile) for profile in EXPECTED_PAYLOADS]
        payload_names = {
            item["path"]
            for report in reports
            for item in report["payload_files"]
        }

        self.assertFalse(
            [
                name for name in payload_names
                if Path(name).suffix.casefold() in FORBIDDEN_PAYLOAD_SUFFIXES
                or Path(name).suffix.casefold() == ".dll" and name not in EXPECTED_PAYLOADS["enhanced-dll"]
            ],
        )

    def test_profile_verifier_reports_no_kutools_path_name_or_asset(self) -> None:
        reports = [run_profile_check(profile) for profile in EXPECTED_PAYLOADS]
        self.assertEqual([], [
            item
            for report in reports
            for item in report["payload_files"]
            if "kutools" in item["path"].casefold()
        ])
        self.assertTrue(all(report["forbidden_matches"] == [] for report in reports))

    def test_profile_source_check_is_deterministic(self) -> None:
        self.assertEqual(run_profile_check("enhanced-dll"), run_profile_check("enhanced-dll"))

    @unittest.skipUnless((ROOT.parent / "v3.4/candidate/src").is_dir(), "Historical v3.4 source is outside the public snapshot")
    def test_v34_mapping_rows_bind_every_source_hash_to_one_target_and_target_files(self) -> None:
        mapping = load_json(V34_MAPPING)
        rows = mapping["per_target_records"]

        self.assertTrue(rows)
        for row in rows:
            with self.subTest(source=row.get("source_path"), target=row.get("target_id")):
                self.assertEqual(
                    {"source_path", "source_sha256", "target_id", "target_files", "source_family"},
                    set(row),
                )
                self.assertEqual("v3.4", row["source_family"])
                self.assertTrue(row["source_path"].startswith("../v3.4/candidate/"))
                source = (ROOT / row["source_path"]).resolve()
                self.assertTrue(source.is_file(), row["source_path"])
                self.assertEqual(hashlib.sha256(source.read_bytes()).hexdigest(), row["source_sha256"])
                self.assertRegex(row["target_id"], r"^NX-[A-Z0-9-]+$")
                self.assertTrue(row["target_files"])
                self.assertTrue(all(isinstance(path, str) and path.startswith("src/") for path in row["target_files"]))

    def test_complete_build_uses_final_only_r62_title(self) -> None:
        self.assertTrue(
            COMPLETE_BUILD.is_file(),
            "missing planned final-only build entry point: tools/build_r46_complete.py",
        )
        source = COMPLETE_BUILD.read_text(encoding="utf-8-sig")
        self.assertIn("내엑셀 v0.01_r105", source)
        self.assertIn("final-only", source.casefold())
        self.assertLess(source.index("check = subprocess.run"), source.index("for source in payload.iterdir()"))

    def test_complete_build_rejects_the_explicit_deployment_package_flag(self) -> None:
        if not COMPLETE_BUILD.is_file():
            self.fail("missing planned final-only build entry point: tools/build_r46_complete.py")
        completed = subprocess.run(
            [sys.executable, str(COMPLETE_BUILD), "--deployment-package"],
            cwd=ROOT,
            capture_output=True,
            text=True,
            encoding="utf-8",
            check=False,
        )

        self.assertNotEqual(0, completed.returncode)
        self.assertRegex(
            (completed.stdout or "") + (completed.stderr or ""),
            r"(?i)deployment.*(?:disabled|rejected|not supported)",
        )

    def test_inventory_binds_artifacts_and_metadata_and_rejects_forbidden_payload(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            (root / "Product.xlam").write_bytes(b"xlam")
            (root / "NxHost32.dll").write_bytes(b"host32")
            (root / "NxHost64.dll").write_bytes(b"host64")
            (root / "NxCore32.dll").write_bytes(b"core32")
            (root / "NxCore64.dll").write_bytes(b"core64")
            (root / "docs").mkdir()
            metadata = set(next(row for row in load_json(BUILD_PROFILES)["profiles"] if row["id"] == "enhanced-dll")["metadata_paths"])
            build_note = next(name for name in metadata if name.startswith("docs/"))
            (root / build_note).write_text("profile=enhanced-dll\n", encoding="utf-8")
            rows = []
            for path in sorted(root.rglob("*"), key=lambda item: item.relative_to(root).as_posix()):
                if path.is_file():
                    rows.append(f"{hashlib.sha256(path.read_bytes()).hexdigest()}  {path.relative_to(root).as_posix()}")
            (root / "SHA256SUMS").write_text("\n".join(rows) + "\n", encoding="utf-8")
            completed = subprocess.run(
                [sys.executable, str(PROFILE_VERIFIER), "--profile", "enhanced-dll", "--payload-root", str(root)],
                cwd=ROOT, capture_output=True, text=True, encoding="utf-8", errors="replace", check=False,
            )
            self.assertEqual(0, completed.returncode, completed.stderr)
            report = json.loads(completed.stdout)
            self.assertEqual(EXPECTED_PAYLOADS["enhanced-dll"], {row["path"] for row in report["artifact_files"]})
            self.assertEqual(metadata, {row["path"] for row in report["metadata_files"]})

            # Office Ribbon XML namespaces are identifiers, not network behavior.
            (root / "Product.xlam").write_bytes(b"http://schemas.microsoft.com/office/2009/07/customui")
            rows = []
            for path in sorted(root.rglob("*"), key=lambda item: item.relative_to(root).as_posix()):
                if path.is_file() and path.name != "SHA256SUMS":
                    rows.append(f"{hashlib.sha256(path.read_bytes()).hexdigest()}  {path.relative_to(root).as_posix()}")
            (root / "SHA256SUMS").write_text("\n".join(rows) + "\n", encoding="utf-8")
            completed = subprocess.run(
                [sys.executable, str(PROFILE_VERIFIER), "--profile", "enhanced-dll", "--payload-root", str(root)],
                cwd=ROOT, capture_output=True, text=True, encoding="utf-8", check=False,
            )
            self.assertEqual(0, completed.returncode, completed.stderr)

            (root / "Product.xlam").write_bytes(b"detong")
            rows = []
            for path in sorted(root.rglob("*"), key=lambda item: item.relative_to(root).as_posix()):
                if path.is_file() and path.name != "SHA256SUMS":
                    rows.append(f"{hashlib.sha256(path.read_bytes()).hexdigest()}  {path.relative_to(root).as_posix()}")
            (root / "SHA256SUMS").write_text("\n".join(rows) + "\n", encoding="utf-8")
            completed = subprocess.run(
                [sys.executable, str(PROFILE_VERIFIER), "--profile", "enhanced-dll", "--payload-root", str(root)],
                cwd=ROOT, capture_output=True, text=True, encoding="utf-8", check=False,
            )
            self.assertNotEqual(0, completed.returncode)


if __name__ == "__main__":
    unittest.main()
