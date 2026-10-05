from __future__ import annotations

import json
import shutil
import subprocess
import sys
import tempfile
import time
import unittest
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT / "src" / "python"))

from naeexcel_hwpx.builder import build_hwpx
from naeexcel_hwpx.cleanup import cleanup_expired
from naeexcel_hwpx.validator import validate_hwpx


HP = "http://www.hancom.co.kr/hwpml/2011/paragraph"
HH = "http://www.hancom.co.kr/hwpml/2011/head"


class HwpxPackageTests(unittest.TestCase):
    def setUp(self) -> None:
        self.temporary = tempfile.TemporaryDirectory()
        self.addCleanup(self.temporary.cleanup)
        self.input = ROOT / "tests" / "fixtures" / "hwpx.json"
        self.output = Path(self.temporary.name) / "fixture.hwpx"

    def build(self) -> dict[str, object]:
        return build_hwpx(self.input, self.output)

    @unittest.skipUnless(shutil.which("powershell.exe") or shutil.which("powershell"), "Windows PowerShell contract: NOT_RUN on this host")
    def test_powershell_builder_accepts_product_wire_schema(self) -> None:
        powershell = shutil.which("powershell.exe") or shutil.which("powershell")
        self.assertIsNotNone(powershell, "Windows PowerShell is required for the embedded-runtime contract")
        script = ROOT / "tools" / "lhexcel_hwpx_table_export.ps1"
        template = ROOT / "templates" / "표.hwpx"
        completed = subprocess.run(
            [
                str(powershell),
                "-NoProfile",
                "-NonInteractive",
                "-ExecutionPolicy",
                "Bypass",
                "-File",
                str(script),
                "-InputPath",
                str(self.input),
                "-OutputPath",
                str(self.output),
                "-TemplatePath",
                str(template),
            ],
            text=True,
            capture_output=True,
            timeout=90,
        )
        self.assertEqual(0, completed.returncode, completed.stderr)
        self.assertTrue(self.output.is_file())
        with zipfile.ZipFile(self.output) as archive:
            table = ET.fromstring(archive.read("Contents/section0.xml")).find(f".//{{{HP}}}tbl")
        self.assertEqual("4", table.attrib["rowCnt"])
        self.assertEqual("3", table.attrib["colCnt"])
        cells = table.findall(f".//{{{HP}}}tc")
        self.assertEqual(11, len(cells))
        merge = next(
            cell.find(f"{{{HP}}}cellSpan").attrib
            for cell in cells
            if cell.find(f"{{{HP}}}cellAddr").attrib == {"colAddr": "1", "rowAddr": "1"}
        )
        self.assertEqual({"colSpan": "2", "rowSpan": "1"}, merge)

        verified = subprocess.run(
            [str(powershell), "-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass", "-File", str(script), "-VerifyPath", str(self.output)],
            text=True,
            capture_output=True,
            timeout=90,
        )
        self.assertEqual(0, verified.returncode, verified.stderr)
        self.assertTrue(json.loads(verified.stdout)["ok"])

    def test_mimetype_is_first_and_stored(self) -> None:
        result = self.build()
        self.assertEqual("PASS", result["status"])
        with zipfile.ZipFile(self.output) as archive:
            first = archive.infolist()[0]
            self.assertEqual("mimetype", first.filename)
            self.assertEqual(zipfile.ZIP_STORED, first.compress_type)
            self.assertEqual(b"application/hwp+zip", archive.read("mimetype"))

    def test_junggodik_hft_is_used_for_every_font_slot(self) -> None:
        self.build()
        with zipfile.ZipFile(self.output) as archive:
            header = ET.fromstring(archive.read("Contents/header.xml"))
        fontfaces = header.find(f".//{{{HH}}}fontfaces")
        self.assertIsNotNone(fontfaces)
        groups = list(fontfaces)
        self.assertEqual(7, len(groups))
        for group in groups:
            font = group.find(f"{{{HH}}}font")
            self.assertEqual("중고딕", font.attrib["face"])
            self.assertEqual("HFT", font.attrib["type"])

    def test_rows_columns_widths_and_merge_match_input(self) -> None:
        expected = json.loads(self.input.read_text(encoding="utf-8"))
        self.build()
        with zipfile.ZipFile(self.output) as archive:
            section = ET.fromstring(archive.read("Contents/section0.xml"))
        table = section.find(f".//{{{HP}}}tbl")
        self.assertEqual(str(expected["rows"]), table.attrib["rowCnt"])
        self.assertEqual(str(expected["columns"]), table.attrib["colCnt"])
        cells = table.findall(f".//{{{HP}}}tc")
        self.assertEqual(11, len(cells))
        widths = {}
        merge = None
        for cell in cells:
            address = cell.find(f"{{{HP}}}cellAddr")
            size = cell.find(f"{{{HP}}}cellSz")
            span = cell.find(f"{{{HP}}}cellSpan")
            widths[int(address.attrib["colAddr"])] = int(size.attrib["width"])
            if address.attrib == {"colAddr": "1", "rowAddr": "1"}:
                merge = span.attrib
        self.assertEqual(42520, sum(widths.values()))
        self.assertEqual({"colSpan": "2", "rowSpan": "1"}, merge)

    def test_runtime_validator_rejects_duplicates_external_macro_and_ole(self) -> None:
        self.build()
        self.assertEqual("PASS", validate_hwpx(self.output)["status"])
        for index, bad_name in enumerate(("word/vbaProject.bin", "Contents/oleObject.bin", "_rels/external.rels", "Contents/header.xml")):
            broken = self.output.with_name(f"broken-{index}.hwpx")
            shutil.copyfile(self.output, broken)
            with zipfile.ZipFile(broken, "a", zipfile.ZIP_DEFLATED) as archive:
                archive.writestr(bad_name, b"x")
            self.assertEqual("FAIL", validate_hwpx(broken)["status"], bad_name)

    def test_builder_rejects_source_metadata_and_overlapping_merges(self) -> None:
        payload = json.loads(self.input.read_text(encoding="utf-8"))
        payload["source_path"] = r"C:\\Users\\person\\source.xlsx"
        bad_input = Path(self.temporary.name) / "path.json"
        bad_input.write_text(json.dumps(payload, ensure_ascii=False), encoding="utf-8")
        with self.assertRaises(ValueError):
            build_hwpx(bad_input, self.output)
        payload = json.loads(self.input.read_text(encoding="utf-8"))
        payload["merges"].append({"row": 2, "column": 1, "row_span": 1, "column_span": 2})
        bad_input.write_text(json.dumps(payload, ensure_ascii=False), encoding="utf-8")
        with self.assertRaises(ValueError):
            build_hwpx(bad_input, self.output)

    def test_validator_rejects_broken_item_count_and_reference(self) -> None:
        self.build()
        with zipfile.ZipFile(self.output) as source:
            entries = {info.filename: source.read(info.filename) for info in source.infolist()}
        header = entries["Contents/header.xml"].replace(b'itemCnt="7"', b'itemCnt="8"', 1)
        content = entries["Contents/content.hpf"].replace(b"Contents/section0.xml", b"Contents/missing.xml", 1)
        for index, replacement in enumerate((("Contents/header.xml", header), ("Contents/content.hpf", content))):
            broken = self.output.with_name(f"contract-{index}.hwpx")
            with zipfile.ZipFile(broken, "w") as archive:
                archive.writestr("mimetype", entries["mimetype"], compress_type=zipfile.ZIP_STORED)
                for name, payload in entries.items():
                    if name != "mimetype":
                        archive.writestr(name, replacement[1] if name == replacement[0] else payload, compress_type=zipfile.ZIP_DEFLATED)
            self.assertEqual("FAIL", validate_hwpx(broken)["status"])

    def test_cleanup_removes_only_expired_owned_files(self) -> None:
        root = Path(self.temporary.name) / "cleanup"
        root.mkdir()
        expired = root / "expired.hwpx"
        current = root / "current.json"
        unrelated = root / "keep.txt"
        for path in (expired, current, unrelated):
            path.write_bytes(b"x")
        old = time.time() - 25 * 3600
        expired.touch()
        import os
        os.utime(expired, (old, old))
        result = cleanup_expired(root, 24)
        self.assertEqual("PASS", result["status"])
        self.assertFalse(expired.exists())
        self.assertTrue(current.exists())
        self.assertTrue(unrelated.exists())

if __name__ == "__main__":
    unittest.main()
