import sys
import tempfile
import unittest
import zipfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[2] / 'tools'))
import build_r47_distribution as builder


class DocumentNameTests(unittest.TestCase):
    def test_reads_new_and_legacy_document_names(self):
        for name in ('NxProductWorkbook', '현재_통합_문서'):
            with self.subTest(name=name), tempfile.TemporaryDirectory() as folder:
                path = Path(folder)/'Product.xlam'
                with zipfile.ZipFile(path, 'w') as z:
                    z.writestr('xl/workbook.xml', '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><workbookPr codeName="'+name+'"/></workbook>')
                self.assertEqual(name, builder.workbook_module_name(path))

    def test_missing_name_fails_closed(self):
        with tempfile.TemporaryDirectory() as folder:
            path = Path(folder)/'Product.xlam'
            with zipfile.ZipFile(path, 'w') as z:
                z.writestr('xl/workbook.xml', '<workbook/>')
            with self.assertRaisesRegex(RuntimeError, 'codeName'):
                builder.workbook_module_name(path)
