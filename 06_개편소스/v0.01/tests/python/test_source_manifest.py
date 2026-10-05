from __future__ import annotations

import hashlib
import json
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path


ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT / "tools"))
from source_manifest import FreezeError, build_manifest, freeze, require_clean_source


def _git(root: Path, *args: str) -> None:
    subprocess.run(["git", "-C", str(root), *args], check=True, capture_output=True)


class SourceManifestTests(unittest.TestCase):
    def _repo(self) -> Path:
        tmp = tempfile.TemporaryDirectory()
        self.addCleanup(tmp.cleanup)
        root = Path(tmp.name)
        _git(root, "init", "-q")
        _git(root, "config", "user.email", "fixture@example.invalid")
        _git(root, "config", "user.name", "fixture")
        return root

    def test_builder_covers_declared_file_and_hashes_it(self) -> None:
        root = self._repo()
        (root / "src").mkdir()
        source = root / "src" / "A.bas"
        source.write_text('Attribute VB_Name = "A"\n', encoding="ascii")
        _git(root, "add", "src/A.bas")
        _git(root, "commit", "-qm", "fixture")
        commit = subprocess.check_output(["git", "-C", str(root), "rev-parse", "HEAD"], text=True).strip()

        manifest = build_manifest(
            root,
            ("src",),
            [{"pattern": "src/**", "origin": "clean-room", "spec_ids": ["P09"], "license": "Proprietary"}],
        )
        self.assertEqual(["src/A.bas"], [row["path"] for row in manifest["files"]])
        self.assertEqual(hashlib.sha256(source.read_bytes()).hexdigest(), manifest["files"][0]["sha256"])
        self.assertEqual(commit, manifest["files"][0]["created_commit"])

    def test_builder_accepts_explicit_fixture_commit_without_git_history(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            (root / "src").mkdir()
            (root / "src" / "A.bas").write_text("new\n", encoding="utf-8")
            manifest = build_manifest(
                root,
                ("src",),
                [{"glob": "src/**", "origin": "clean-room", "spec_ids": ["P09"], "license": "project"}],
                {"src/A.bas": "a" * 40},
            )
            self.assertEqual("a" * 40, manifest["files"][0]["created_commit"])

    def test_untracked_and_dirty_source_fail_closed(self) -> None:
        root = self._repo()
        (root / "src").mkdir()
        source = root / "src" / "A.bas"
        source.write_text("new\n", encoding="utf-8")
        with self.assertRaisesRegex(FreezeError, "SOURCE_UNTRACKED_OR_DIRTY"):
            require_clean_source([source], root=root)
        _git(root, "add", "src/A.bas")
        _git(root, "commit", "-qm", "fixture")
        source.write_text("dirty\n", encoding="utf-8")
        with self.assertRaisesRegex(FreezeError, "SOURCE_UNTRACKED_OR_DIRTY"):
            require_clean_source([source], root=root)

    def test_ignored_source_fails_closed_even_when_status_is_clean(self) -> None:
        root = self._repo()
        (root / ".gitignore").write_text("src/\n", encoding="utf-8")
        (root / "src").mkdir()
        source = root / "src" / "A.bas"
        source.write_text("ignored\n", encoding="utf-8")
        _git(root, "add", ".gitignore")
        _git(root, "commit", "-qm", "fixture")
        self.assertEqual("", subprocess.check_output(["git", "-C", str(root), "status", "--porcelain", "--", "src/A.bas"], text=True))
        with self.assertRaisesRegex(FreezeError, "SOURCE_UNTRACKED_OR_DIRTY: src/A.bas"):
            require_clean_source([source], root=root)

    def test_duplicate_rule_and_missing_commit_are_rejected(self) -> None:
        root = self._repo()
        (root / "src").mkdir()
        (root / "src" / "A.bas").write_text("new\n", encoding="utf-8")
        with self.assertRaisesRegex(FreezeError, "exactly one"):
            build_manifest(
                root,
                ("src",),
                [
                    {"pattern": "src/**", "origin": "x", "spec_ids": ["P09"], "license": "p"},
                    {"pattern": "src/*.bas", "origin": "x", "spec_ids": ["P09"], "license": "p"},
                ],
                {"src/A.bas": "a" * 40},
            )

    def test_freeze_writes_commit_and_artifact_identity(self) -> None:
        root = self._repo()
        (root / "src").mkdir()
        source = root / "src" / "A.bas"
        source.write_text("new\n", encoding="utf-8")
        (root / "provenance").mkdir()
        (root / "provenance" / "origins.json").write_text(
            json.dumps({"rules": [{"pattern": "src/**", "origin": "x", "spec_ids": ["P09"], "license": "p"}]}),
            encoding="utf-8",
        )
        candidate = root / "candidate.xlam"
        candidate.write_bytes(b"fixture artifact")
        _git(root, "add", "src/A.bas", "provenance/origins.json")
        _git(root, "commit", "-qm", "fixture")
        result = freeze(root, candidate, ("src",))
        self.assertEqual(hashlib.sha256(candidate.read_bytes()).hexdigest(), result["artifact_sha256"])
        preserved = Path(result["preserved"])
        self.assertTrue((preserved / "source-manifest.json").is_file())
        self.assertTrue((preserved / "freeze-manifest.json").is_file())

    def test_product_freeze_preserves_single_xlam_and_rejects_external_runtime(self) -> None:
        root = self._repo()
        (root / "src").mkdir()
        (root / "src" / "A.bas").write_text("new\n", encoding="utf-8")
        (root / "provenance").mkdir()
        (root / "provenance" / "origins.json").write_text(
            json.dumps({"rules": [{"pattern": "src/**", "origin": "x", "spec_ids": ["P09"], "license": "p"}]}),
            encoding="utf-8",
        )
        _git(root, "add", "src/A.bas", "provenance/origins.json")
        _git(root, "commit", "-qm", "fixture")
        product = root / "Product.xlam"
        product.write_bytes(b"product")
        result = freeze(root, product, ("src",))
        current = root / "out" / "frozen" / "current"
        self.assertEqual(["Product.xlam", "freeze-manifest.json", "source-manifest.json"], sorted(path.name for path in current.iterdir()))
        manifest = json.loads((current / "freeze-manifest.json").read_text(encoding="utf-8"))
        self.assertEqual(result["product_bundle_sha256"], manifest["product_bundle_sha256"])
        self.assertEqual([{"path": "Product.xlam", "sha256": hashlib.sha256(product.read_bytes()).hexdigest()}], manifest["product_files"])

        (root / "python.exe").write_bytes(b"forbidden")
        with self.assertRaisesRegex(FreezeError, "forbidden external runtime"):
            freeze(root, product, ("src",))

    def test_product_freeze_needs_no_sidecar_manifest(self) -> None:
        root = self._repo()
        (root / "src").mkdir()
        (root / "src" / "A.bas").write_text("new\n", encoding="utf-8")
        (root / "provenance").mkdir()
        (root / "provenance" / "origins.json").write_text(
            json.dumps({"rules": [{"pattern": "src/**", "origin": "x", "spec_ids": ["P09"], "license": "p"}]}),
            encoding="utf-8",
        )
        _git(root, "add", "src/A.bas", "provenance/origins.json")
        _git(root, "commit", "-qm", "fixture")
        product = root / "Product.xlam"
        product.write_bytes(b"product")
        result = freeze(root, product, ("src",))
        self.assertEqual(["Product.xlam"], [row["path"] for row in result["product_files"]])


if __name__ == "__main__":
    unittest.main()
