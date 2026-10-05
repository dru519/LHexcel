import sys,unittest
from pathlib import Path
ROOT=Path(__file__).resolve().parents[2]
sys.path.insert(0,str(ROOT/'tools'))
from distribution_private_names import compact_private_names

class PrivateNameTests(unittest.TestCase):
    def test_private_helpers_rename_but_public_callback_and_literals_stay(self):
        source='Option Explicit\nPrivate Const NX_VALUE As Long = 7\nPrivate Function NxHidden() As Long\nNxHidden = NX_VALUE\nEnd Function\nPublic Sub Callback()\nDebug.Print NxHidden(), "NX_OTHER"\nEnd Sub\n'
        transformed,names=compact_private_names(source)
        self.assertEqual(set(names),{'nxhidden','nx_value'})
        self.assertIn('Public Sub Callback()',transformed)
        self.assertIn('"NX_OTHER"',transformed)
        self.assertNotIn('NxHidden',transformed)
        self.assertEqual(compact_private_names(transformed)[0],transformed)
    def test_runtime_strings_named_arguments_and_public_parameters_prevent_rename(self):
        source='Private Sub NxHandler()\nEnd Sub\n'
        self.assertEqual(compact_private_names(source,'"NxHandler"')[0],source)
        self.assertEqual(compact_private_names(source+'Call foo(NxHandler:=1)')[1],{})
        self.assertEqual(compact_private_names(source+'Public Sub Read(ByVal NxHandler As Long)\nEnd Sub')[1],{})
    def test_comments_and_escaped_string_preserved(self):
        source='Private Const NX_VALUE As Long = 1\n\x27 NX_VALUE\nRem NX_VALUE\nDebug.Print NX_VALUE, "a ""b"""'
        result,names=compact_private_names(source)
        self.assertIn('\x27 NX_VALUE',result)
        self.assertIn('Rem NX_VALUE',result)
        self.assertIn('"a ""b"""',result)
        self.assertIn('nx_value',names)

if __name__=='__main__':unittest.main()
