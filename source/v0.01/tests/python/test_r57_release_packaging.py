from __future__ import annotations

import hashlib
import importlib.util
import sys
import tempfile
import unittest
from unittest.mock import Mock, patch
from types import SimpleNamespace
from pathlib import Path


ROOT = Path(__file__).resolve().parents[2]
INTERNAL_BUILDER = ROOT / "tools" / "build_r47_distribution.py"
ENHANCED_BUILDER = ROOT / "tools" / "build_enhanced_distribution.py"


def load_enhanced_builder():
    spec = importlib.util.spec_from_file_location("build_enhanced_distribution", ENHANCED_BUILDER)
    if spec is None or spec.loader is None:
        raise RuntimeError("enhanced distribution builder import failed")
    module = importlib.util.module_from_spec(spec)
    tool_path = str(ENHANCED_BUILDER.parent)
    sys.path.insert(0, tool_path)
    try:
        spec.loader.exec_module(module)
    finally:
        sys.path.remove(tool_path)
    return module


class R57ReleasePackagingTests(unittest.TestCase):
    def test_internal_distribution_refuses_existing_destination_and_publishes_atomically(self) -> None:
        source = INTERNAL_BUILDER.read_text(encoding="utf-8")
        self.assertIn("distribution destination already exists", source)
        self.assertIn("temporary_destination.replace(DISTRIBUTION_XLAM)", source)
        self.assertNotIn("def remove_existing_distribution_files", source)
        self.assertNotIn("child.unlink()", source)

    def test_enhanced_builder_declares_exact_r105_inventory_and_fail_closed_verification(self) -> None:
        source = ENHANCED_BUILDER.read_text(encoding="utf-8")
        for token in (
            "payload/Product.xlam",
            "payload/NxHost32.dll",
            "payload/NxHost64.dll",
            "payload/NxCore32.dll",
            "payload/NxCore64.dll",
            "payload/SHA256SUMS",
            "payload/docs/r105-build.md",
            "deployment/Install-NxEnhanced.ps1",
            "PACKAGE_SHA256SUMS",
            "refusing existing enhanced destination",
            "Verify-NxEnhancedPackage.ps1",
            "zipfile.ZipFile",
        ):
            self.assertIn(token, source)
        self.assertIn('prefix=f"staging-{destination_root.name}-"', source)
        self.assertNotIn('prefix=f".{destination_root.name}."', source)

    def test_assemble_copies_exact_payload_and_writes_sorted_hash_manifest(self) -> None:
        module = load_enhanced_builder()
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            complete = root / "complete"
            source = root / "source"
            staging = root / "staging"
            (complete / "docs").mkdir(parents=True)
            (source / "build").mkdir(parents=True)
            for name, data in (
                ("Product.xlam", b"xlam"),
                ("NxHost32.dll", b"x86"),
                ("NxHost64.dll", b"x64"),
                ("NxCore32.dll", b"core-x86"),
                ("NxCore64.dll", b"core-x64"),
                ("docs/r105-build.md", b"r105\n"),
            ):
                path = complete / name
                path.parent.mkdir(parents=True, exist_ok=True)
                path.write_bytes(data)
            module.write_manifest(complete, "SHA256SUMS")
            for name in module.DEPLOYMENT_FILES:
                (source / "build" / name).write_text(name + "\n", encoding="ascii")

            # Assembly is tested with synthetic bytes. Profile validation has
            # its own real-XLAM tests and is required on the published artifact.
            profile_check = Mock()
            with patch.dict(sys.modules, {"distribution_source_profile": SimpleNamespace(verify_enhanced=profile_check)}):
                module.assemble(complete, source, staging)
            profile_check.assert_called_once_with(complete / "Product.xlam")

            actual = sorted(
                path.relative_to(staging).as_posix()
                for path in staging.rglob("*")
                if path.is_file()
            )
            self.assertEqual(sorted(module.PACKAGE_FILES), actual)
            rows = (staging / "PACKAGE_SHA256SUMS").read_text(encoding="utf-8").splitlines()
            paths = [row.split("  ", 1)[1] for row in rows]
            self.assertEqual(sorted(paths, key=lambda value: value.encode("utf-8")), paths)
            for row in rows:
                digest, relative = row.split("  ", 1)
                self.assertEqual(hashlib.sha256((staging / relative).read_bytes()).hexdigest(), digest)

            self.assertNotIn("정식 배포 금지", (staging / "README.md").read_text(encoding="utf-8"))
            readme = (staging / "README.md").read_text(encoding="utf-8")
            self.assertIn((ROOT / "docs/public-use-terms.md").read_text(encoding="utf-8"), readme)
            for instruction in ("설치·실행 안내", "COM", "제거", "일반 업무에 무료", "유료 제품·서비스"):
                self.assertIn(instruction, readme)
            candidate = root / "candidate"
            module.assemble(complete, source, candidate, development_candidate=True)
            readme = (candidate / "README.md").read_text(encoding="utf-8")
            self.assertIn("정식 배포 금지", readme)
            self.assertIn("배포 보호를 적용하지 않았습니다", readme)
            self.assertIn("설치·실행 안내", readme)
            module.verify_manifest(candidate, "PACKAGE_SHA256SUMS")
            self.assertEqual((candidate / "payload/Product.xlam").read_bytes(), b"xlam")

    def test_publish_refuses_existing_folder_or_zip(self) -> None:
        module = load_enhanced_builder()
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            staging = root / "staging"
            staging.mkdir()
            (staging / "README.md").write_text("r62\n", encoding="utf-8")
            folder = root / "release"
            archive = root / "release.zip"
            folder.mkdir()
            with self.assertRaisesRegex(RuntimeError, "refusing existing enhanced destination"):
                module.publish(staging, folder, archive)


if __name__ == "__main__":
    unittest.main()
