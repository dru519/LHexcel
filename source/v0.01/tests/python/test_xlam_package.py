import hashlib
import importlib.util
import json
import os
import shutil
import subprocess
import sys
import tempfile
import unittest
import zipfile
from pathlib import Path
from unittest import mock
from xml.etree import ElementTree as ET

ROOT = Path(__file__).resolve().parents[2]
BUILD = ROOT / "build"
MANIFEST = BUILD / "manifests/Product.json"
PYTHON_SUBPROCESS_ENV = {
    **os.environ,
    "PYTHONUTF8": "1",
    "PYTHONIOENCODING": "utf-8",
}


class ProductXlamPackageContractTests(unittest.TestCase):
    def test_product_manifest_is_reproducible(self):
        result = subprocess.run(
            [sys.executable, "tools/build_product_manifest.py", "--check"],
            cwd=ROOT,
            text=True,
            capture_output=True,
            encoding="utf-8",
            env=PYTHON_SUBPROCESS_ENV,
        )
        self.assertEqual(0, result.returncode, result.stderr)

    def test_ribbon_relationship_uses_custom_ui_2007_contract(self):
        source = (BUILD / "inject_ribbon.py").read_text(encoding="utf-8")
        manifest = json.loads(MANIFEST.read_text(encoding="utf-8"))
        expected = "http://schemas.microsoft.com/office/2007/relationships/ui/extensibility"
        self.assertEqual(expected, manifest["ribbon"]["relationship_type"])
        self.assertIn(
            f'UI_REL = "{expected}"',
            source,
        )
        self.assertNotIn(
            'UI_REL = "http://schemas.microsoft.com/office/2006/relationships/ui/extensibility"',
            source.splitlines(),
        )

    def test_manifest_is_flat_hash_bound_and_excludes_tests(self):
        manifest = json.loads(MANIFEST.read_text(encoding="utf-8"))
        self.assertEqual(
            {
                "schema_version",
                "suite",
                "owner_phase",
                "default_build_profile",
                "r46_contracts",
                "modules",
                "form_bindings",
                "diagnostic_only",
                "forbidden_prefixes",
                "ribbon",
                "release_scope",
            },
            set(manifest),
        )
        self.assertEqual((3, "Product", "product"), (manifest["schema_version"], manifest["suite"], manifest["owner_phase"]))
        paths = [row["path"] for row in manifest["modules"]]
        self.assertEqual(len(paths), len(set(paths)))
        self.assertTrue(paths)
        for row in manifest["modules"]:
            self.assertTrue(row["path"].startswith("src/vba/"))
            self.assertNotIn("tests/", row["path"])
            source = ROOT / row["path"]
            self.assertTrue(source.is_file())
            self.assertEqual(hashlib.sha256(source.read_bytes()).hexdigest(), row["sha256"])
        bindings = manifest["form_bindings"]
        self.assertEqual(
            len([row for row in manifest["modules"] if row["kind"] == "form_layout"]),
            len(bindings),
        )
        self.assertNotIn(
            "src/vba/features/data/ui/FNxFocusOverlay.form.json",
            {row["layout"] for row in bindings},
        )
        for binding in bindings:
            layout = next(row for row in manifest["modules"] if row["path"] == binding["layout"])
            code = next(row for row in manifest["modules"] if row["path"] == binding["code"])
            self.assertEqual("form_layout", layout["kind"])
            self.assertEqual("form_code", code["kind"])
            self.assertEqual(binding["code"], layout["code_path"])
        self.assertFalse(any(path.startswith("src/vba/features/hwpx/") for path in paths))
        self.assertEqual(15, len({path for path in paths if path.startswith("src/vba/features/hangul/")}))
        self.assertFalse(any("Roster" in path for path in paths))
        scope = manifest["release_scope"]
        self.assertEqual(len(scope["included_feature_ids"]), scope["feature_count"])
        self.assertEqual([], scope["excluded_feature_ids"])
        contract = ROOT / scope["feature_contract"]["path"]
        self.assertEqual(hashlib.sha256(contract.read_bytes()).hexdigest(), scope["feature_contract"]["sha256"])

    def test_build_scripts_keep_product_boundary(self):
        build = (BUILD / "Build-Xlam.ps1").read_text(encoding="utf-8")
        importer = (BUILD / "Import-ProductVba.ps1").read_text(encoding="utf-8")
        self.assertIn("Product.json", build)
        self.assertIn("inject_ribbon.py", build)
        self.assertIn("Existing product XLAM destination is rejected", build)
        self.assertIn("tests/", importer)
        self.assertIn("Product manifest contains forbidden test paths", importer)
        self.assertIn("$expectedProductComponents.Count -ne ($records.Count - $formCodeCount)", importer)
        self.assertNotIn("$expectedProductComponents.Count -ne 161", importer)
        self.assertIn("$actualProductComponents.Count -ne $expectedProductComponents.Count", importer)
        self.assertNotIn("actual importable count is not 114", importer)
        self.assertIn("form_bindings", importer)
        self.assertNotIn("build_hwpx_embedded_resources.py", build)
        self.assertNotIn("validate_product_runtime.py", build)
        self.assertIn("Product Hangul HWPX source inventory rejected", build)
        self.assertIn("$removedModules.Count -ne 0", build)
        self.assertIn("$hangulModules.Count -lt 9", build)
        self.assertIn("NX-HANGUL-TABLE-SEND", build)
        self.assertNotIn("sidecar-manifest", build)
        self.assertNotIn("EmbeddedPythonRoot", build)
        self.assertNotIn("python.exe", build)
        self.assertIn("publishStage", build)
        self.assertIn("Publish transaction failed; phase=", build)
        self.assertIn("cleaning final outputs:", build)
        self.assertIn("Copy-Item -LiteralPath $publishStage -Destination $publishCandidate", build)
        self.assertIn("Move-Item -LiteralPath $publishCandidate -Destination $destination", build)
        self.assertIn("$ErrorActionPreference = 'Stop'", build)
        self.assertIn("Assert-NoReparsePath", build)
        self.assertIn("[string]$Boundary", build)
        self.assertIn("Product reparse boundary was not reached", build)
        self.assertIn("Get-VerifiedHash $RibbonPath $sourceRoot", build)
        self.assertFalse((BUILD / "Import-Vba.ps1").exists())
        self.assertNotIn("Build-TestXlam.ps1", build)

    def test_release_scope_is_runtime_free_and_hash_bound(self):
        manifest = json.loads(MANIFEST.read_text(encoding="utf-8"))
        self.assertNotIn("runtime", manifest)
        scope = manifest["release_scope"]
        contract = ROOT / scope["feature_contract"]["path"]
        self.assertEqual(hashlib.sha256(contract.read_bytes()).hexdigest(), scope["feature_contract"]["sha256"])
        feature_contract = json.loads(contract.read_text(encoding="utf-8"))
        self.assertEqual(
            set(feature_contract["lineage"]["release_feature_ids"]),
            set(scope["included_feature_ids"]),
        )

    def test_manifest_form_bindings_are_a_bijection(self):
        manifest = json.loads(MANIFEST.read_text(encoding="utf-8"))
        layouts = {row["path"] for row in manifest["modules"] if row["kind"] == "form_layout"}
        codes = {row["path"] for row in manifest["modules"] if row["kind"] == "form_code"}
        bindings = {(row["layout"], row["code"]) for row in manifest["form_bindings"]}
        self.assertEqual(layouts, {layout for layout, _ in bindings})
        self.assertEqual(codes, {code for _, code in bindings})

    def test_product_manifest_has_no_python_runtime_payload(self):
        manifest = json.loads(MANIFEST.read_text(encoding="utf-8"))
        serialized = json.dumps(manifest, ensure_ascii=False).lower()
        self.assertNotIn("python.exe", serialized)
        self.assertNotIn("src/python/", serialized)
        self.assertNotIn(".dll", serialized)
        self.assertNotIn(".pyd", serialized)

    def test_ribbon_injection_is_atomic_and_verifiable(self):
        module_path = BUILD / "inject_ribbon.py"
        spec = importlib.util.spec_from_file_location("inject_ribbon", module_path)
        module = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(module)
        ribbon = ROOT / "src/ribbon/customUI14.xml"
        with tempfile.TemporaryDirectory() as tmp:
            xlam = Path(tmp) / "product.xlam"
            with zipfile.ZipFile(xlam, "w") as z:
                z.writestr("[Content_Types].xml", '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>')
                z.writestr("_rels/.rels", '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>')
                z.writestr("xl/vbaProject.bin", b"vba")
            module.inject(xlam, ribbon)
            with zipfile.ZipFile(xlam) as z:
                self.assertEqual(1, z.namelist().count("customUI/customUI14.xml"))
                self.assertEqual(ribbon.read_bytes(), z.read("customUI/customUI14.xml"))
                rels = ET.fromstring(z.read("_rels/.rels"))
                matches = [
                    row for row in rels
                    if row.attrib.get("Type") == "http://schemas.microsoft.com/office/2007/relationships/ui/extensibility"
                ]
                self.assertEqual(1, len(matches))

    def test_ribbon_injection_closes_mkstemp_handle_before_zip_write(self):
        module_path = BUILD / "inject_ribbon.py"
        spec = importlib.util.spec_from_file_location("inject_ribbon_handle", module_path)
        module = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(module)
        ribbon = ROOT / "src/ribbon/customUI14.xml"
        with tempfile.TemporaryDirectory() as tmp:
            xlam = Path(tmp) / "product.xlam"
            with zipfile.ZipFile(xlam, "w") as package:
                package.writestr("[Content_Types].xml", '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>')
                package.writestr("_rels/.rels", '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>')
                package.writestr("xl/vbaProject.bin", b"vba")

            temp_path = Path(tmp) / "locked-on-windows.tmp"
            fd = os.open(temp_path, os.O_CREAT | os.O_RDWR)
            real_zip_file = zipfile.ZipFile

            def open_zip_after_handle_close(*args, **kwargs):
                if Path(args[0]) == temp_path:
                    with self.assertRaises(OSError):
                        os.fstat(fd)
                return real_zip_file(*args, **kwargs)

            with mock.patch.object(module.tempfile, "mkstemp", return_value=(fd, str(temp_path))):
                with mock.patch.object(module.zipfile, "ZipFile", side_effect=open_zip_after_handle_close):
                    module.inject(xlam, ribbon)

            self.assertFalse(temp_path.exists())

    def test_ribbon_injection_rejects_legacy_custom_ui_relationship(self):
        module_path = BUILD / "inject_ribbon.py"
        spec = importlib.util.spec_from_file_location("inject_ribbon_legacy", module_path)
        module = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(module)
        ribbon = ROOT / "src/ribbon/customUI14.xml"
        with tempfile.TemporaryDirectory() as tmp:
            xlam = Path(tmp) / "legacy.xlam"
            rels = (
                '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                '<Relationship Id="rIdLegacy" '
                'Type="http://schemas.microsoft.com/office/2006/relationships/ui/extensibility" '
                'Target="customUI/customUI14.xml"/>'
                '</Relationships>'
            )
            with zipfile.ZipFile(xlam, "w") as package:
                package.writestr("[Content_Types].xml", '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>')
                package.writestr("_rels/.rels", rels)
                package.writestr("xl/vbaProject.bin", b"vba")
            with self.assertRaisesRegex(ValueError, "existing customUI relationship rejected"):
                module.inject(xlam, ribbon)


if __name__ == "__main__":
    unittest.main()
