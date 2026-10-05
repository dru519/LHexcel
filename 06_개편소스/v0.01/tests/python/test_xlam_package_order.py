import sys
import tempfile
import unittest
import zipfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[2]/'tools'))
from xlam_package_order import normalize, ordered_names

class PackageOrderTests(unittest.TestCase):
    def test_order_and_contents_and_idempotence(self):
        with tempfile.TemporaryDirectory() as root:
            p = Path(root)/'fixture.xlam'
            members = {'customUI/customUI14.xml':b'ui', 'xl/vbaProject.bin':b'vba',
                       'xl/workbook.xml':b'workbook', '_rels/.rels':b'rels', '[Content_Types].xml':b'types'}
            with zipfile.ZipFile(p,'w',zipfile.ZIP_DEFLATED) as z:
                for n,b in members.items(): z.writestr(n,b)
            normalize(p)
            with zipfile.ZipFile(p) as z:
                self.assertEqual(['[Content_Types].xml','_rels/.rels','xl/workbook.xml'],z.namelist()[:3])
                self.assertEqual(members,{n:z.read(n) for n in z.namelist()})
            before=p.read_bytes(); normalize(p)
            self.assertEqual(before,p.read_bytes())

    def test_duplicate_rejected(self):
        with self.assertRaises(ValueError): ordered_names(['a','a'])

    def test_signed_input_preserved(self):
        with tempfile.TemporaryDirectory() as root:
            p=Path(root)/'signed.xlam'
            with zipfile.ZipFile(p,'w') as z: z.writestr('xl/vbaProjectSignature.bin',b'signature')
            before=p.read_bytes()
            with self.assertRaises(ValueError): normalize(p)
            self.assertEqual(before,p.read_bytes())
