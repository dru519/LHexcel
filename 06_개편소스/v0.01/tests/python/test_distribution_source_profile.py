from pathlib import Path
import sys, unittest
ROOT=Path(__file__).resolve().parents[2]
sys.path.insert(0,str(ROOT/'tools'))
from distribution_source_profile import compact_comments,transform_source,verify_enhanced_source,REQUIRED_HOST

class SourceProfiles(unittest.TestCase):
    def test_literals_attributes_directives_and_inline_comments_survive(self):
        code='Attribute VB_Name = "Demo"\nOption Explicit\n#If VBA7 Then\nConst x = "Rem test \' quote"\n#End If\nSub A()\nRem developer notes\n\' generated source path\nx = "a" \' inline\nEnd Sub\n'
        reduced=compact_comments(code)
        self.assertEqual(reduced,code.replace('Rem developer notes\n','').replace("' generated source path\n",''))
    def test_attribution_is_preserved(self):
        code="' Copyright owner\n' license terms\nOption Explicit\n"
        self.assertEqual(compact_comments(code),code)
    def test_profile_keeps_internal_algorithm_only(self):
        source=(ROOT/'src/vba/features/file/compare/NxWorkbookCompare.bas').read_text(encoding='utf-8')
        internal=transform_source('NxWorkbookCompare',source,'internal-xlam')
        enhanced=transform_source('NxWorkbookCompare',source,'enhanced-dll')
        self.assertIn('If leftRows(ordinal, 0) <> rightRows(ordinal, 0)',internal)
        self.assertNotIn('If leftRows(ordinal, 0) <> rightRows(ordinal, 0)',enhanced)
        self.assertIn('If service Is Nothing Then NxRaiseContractError',enhanced)
        self.assertIn('NxWorkbookCompareFormatSignature',enhanced)
        self.assertIn('If service Is Nothing Then Exit Function',internal)
    def test_changed_shape_fails_closed(self):
        with self.assertRaises(ValueError):transform_source('NxWorkbookCompare','Option Explicit','enhanced-dll')
    def test_other_module_preserves_executable_source(self):
        self.assertEqual(transform_source('Other','Option Explicit\n','enhanced-dll'),'Option Explicit\n')
    def test_compiled_vba_identifier_casing_cannot_hide_fallback(self):
        with self.assertRaises(ValueError):
            verify_enhanced_source(REQUIRED_HOST + '\nIf CBool(baseCell.hasFormula) Then')

if __name__=='__main__':unittest.main()
