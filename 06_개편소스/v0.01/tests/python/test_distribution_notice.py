import sys
import tempfile
import unittest
import zipfile
from unittest.mock import MagicMock, patch
from datetime import date
from pathlib import Path
from xml.etree import ElementTree as ET

ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT / 'tools'))
import distribution_notice as notice
import build_r47_distribution as builder


class DistributionNoticeTests(unittest.TestCase):
    def test_notice_roundtrips_korean_vba_codepage(self):
        self.assertEqual(notice.NOTICE_TEXT, notice.NOTICE_TEXT.encode('cp949').decode('cp949'))

    def test_vba_notice_preserves_code_attributes_and_endings(self):
        for nl in ('\n', '\r\n'):
            original = nl.join(['Attribute VB_Name = "Example"', 'Attribute VB_GlobalNameSpace = False',
                                'Option Explicit', 'Public Sub Run()', 'End Sub', ''])
            changed = notice.with_vba_notice(original)
            self.assertTrue(changed.startswith(original.split('Option Explicit')[0]))
            self.assertEqual(original, notice.without_vba_notice(changed))
            self.assertEqual(changed, notice.with_vba_notice(changed))
            self.assertEqual(1, changed.count(notice.MARKER_BEGIN))

    def test_empty_and_generated_source(self):
        for source in ('', "' GENERATED FILE\nOption Explicit\n", 'Option Explicit'):
            self.assertEqual(source, notice.without_vba_notice(notice.with_vba_notice(source)))

    def test_every_notice_line_is_a_comment(self):
        self.assertTrue(all(line.startswith("' ") for line in notice.with_vba_notice('').splitlines()))
        self.assertIn('차단하는 장치는 아닙니다', notice.NOTICE_TEXT)

    def test_unexpected_marker_fails_closed(self):
        with self.assertRaises(ValueError):
            notice.with_vba_notice("' " + notice.MARKER_BEGIN + '\nOption Explicit\n')

    def test_upgrade_preserves_code_from_both_historical_notice_formats(self):
        for newline in ('\n', '\r\n'):
            source = newline.join(['Attribute VB_Name = "Example"', 'Option Explicit',
                                   'Public Sub Run()', 'End Sub', ''])
            offset = source.index('Option Explicit')
            for old_notice in (notice.LEGACY_NOTICE_TEXT, notice.LEGACY_PROTECTED_NOTICE_TEXT):
                historical = source[:offset] + notice._block(newline, old_notice) + source[offset:]
                upgraded = notice.with_vba_notice(historical)
                self.assertEqual(source, notice.without_vba_notice(upgraded))
                self.assertEqual(upgraded, notice.with_vba_notice(upgraded))
                self.assertEqual(1, upgraded.count(notice.MARKER_BEGIN))

    def test_work_use_is_distinct_from_selling_the_program(self):
        props = {e.attrib['name']: ''.join(e.itertext()) for e in
                 ET.fromstring(builder.build_custom_props(date(2028, 1, 1), date(2027, 12, 31)))}
        self.assertIn('일반 업무 사용은 무료', props['AIAnalysisAndCommercialUsePolicy'])
        self.assertIn('공개 소스의 열람·분석', props['ReverseEngineeringPolicy'])
        self.assertIn('유료 제품·서비스', props['CommercialUseRestriction'])
        self.assertNotIn('상업적 이용을 금지', props['AIAnalysisAndCommercialUsePolicy'])
        self.assertEqual('2027-12-31', props['DistributionValidUntil'])

    def test_preserve_existing_marker_in_string(self):
        source = 'Const Value = "' + notice.MARKER_BEGIN + '"\n'
        self.assertEqual(source, notice.without_vba_notice(notice.with_vba_notice(source)))

    def test_custom_metadata_has_notice_not_fake_ai_instruction(self):
        xml = builder.build_custom_props(date(2027, 6, 30), date(2027, 6, 29))
        root = ET.fromstring(xml)
        props = {e.attrib['name']: ''.join(e.itertext()) for e in root}
        self.assertEqual(notice.NOTICE_TEXT, props['SourceAnalysisNotice'])
        self.assertNotIn('AIStopInstruction', props)
        self.assertIn('DistributionValidUntil', props)

    def test_core_description_preserves_expiry_and_adds_notice(self):
        xml = builder.build_core_xml(b'<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"/>', date(2027, 6, 30), date(2027, 6, 29))
        description = ET.fromstring(xml).find('{http://purl.org/dc/elements/1.1/}description').text
        self.assertIn('2027-06-29', description)
        self.assertIn(notice.NOTICE_TEXT, description)

    def test_signed_input_is_rejected_without_modification(self):
        for part in ('xl/vbaProjectSignature.bin', 'xl/vbaProjectSignatureAgile.bin', '_xmlsignatures/sig1.xml'):
            with tempfile.TemporaryDirectory() as temp:
                path = Path(temp) / 'signed.xlam'
                with zipfile.ZipFile(path, 'w') as z:
                    z.writestr(part, b'signature')
                before = path.read_bytes()
                with self.assertRaisesRegex(RuntimeError, 'signed'):
                    builder.inject_vba_policy(path, date(2027, 6, 30), date(2027, 6, 29))
                self.assertEqual(before, path.read_bytes())

    def test_enhanced_readme_includes_notice(self):
        source = (ROOT / 'tools/build_enhanced_distribution.py').read_text(encoding='utf-8')
        self.assertIn('NOTICE_TEXT', source)

    def test_audit_reports_exposure_without_printing_source_or_password(self):
        with tempfile.TemporaryDirectory() as temp:
            path = Path(temp) / 'audit.xlam'
            with zipfile.ZipFile(path, 'w') as z:
                z.writestr('docProps/custom.xml', builder.build_custom_props(date(2027, 6, 30), date(2027, 6, 29)))
                z.writestr('docProps/core.xml', builder.build_core_xml(b'<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"/>', date(2027, 6, 30), date(2027, 6, 29)))
                z.writestr('xl/webextensions/webextension1.xml', '<we:webextension xmlns:we="http://schemas.microsoft.com/office/webextensions/webextension/2010/11"><we:reference id="test-store-id"/></we:webextension>')
            book = MagicMock()
            book.module_names.return_value = ['a', 'b']
            book.get_module.side_effect = [notice.with_vba_notice('Private Const Secret = "source-not-for-report"\n'), 'Option Explicit\n']
            book.vba_project.return_value.protection.has_password = True
            original = path.read_bytes()
            with patch.object(notice, 'ExcelFile') as fake:
                fake.return_value.__enter__.return_value = book
                result = notice.audit_distribution(path)
            self.assertEqual(2, result['readable_module_count'])
            self.assertEqual(1, result['notified_module_count'])
            self.assertEqual(1, result['web_extension_count'])
            self.assertTrue(result['vba_password_present'])
            self.assertTrue(result['custom_notice_present'] and result['core_notice_present'])
            self.assertNotIn('source-not-for-report', str(result))
            self.assertEqual(original, path.read_bytes())


if __name__ == '__main__':
    unittest.main()
