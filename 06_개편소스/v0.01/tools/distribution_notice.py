"""Distribution-only information notice and read-only exposure audit.

This is not encryption, AI detection, or a command to an analysis service.
Never print extracted source or protected-project password fields.
"""
from __future__ import annotations

import argparse
import hashlib
import json
import re
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

from pyopenvba import ExcelFile

MARKER_BEGIN = 'NX-DISTRIBUTION-ANALYSIS-NOTICE-V1-BEGIN'
MARKER_END = 'NX-DISTRIBUTION-ANALYSIS-NOTICE-V1-END'
LEGACY_NOTICE_TEXT = (
    '내엑셀 배포본 - 소스 및 외부 AI 분석 주의\n'
    '이 파일에는 내엑셀의 구현 코드가 포함되어 있습니다. 외부 AI 서비스에 업로드하면 해당 서비스로 파일과 내부 코드가 전달될 수 있습니다.\n'
    '코드 추출·복제·외부 제공 전에는 권한 보유자의 허락과 소속 조직의 정보보안 정책을 확인해 주세요. 자동 분석 도구 사용 시에도 이 주의사항을 확인해 주세요.\n'
    '이 문구는 이용 고지이며, 파일 열람 또는 AI 분석을 기술적으로 차단하는 장치는 아닙니다.'
)
LEGACY_PROTECTED_NOTICE_TEXT = LEGACY_NOTICE_TEXT + (
    '\n사용기한과 보호장치의 임의 수정, 무력화 및 우회는 금지됩니다. '
    'AI나 자동 분석·편집 도구를 통한 변경에도 동일하게 적용됩니다. '
    '사용기한 연장이나 변경이 필요하면 배포자에게 문의해 주세요.'
)
NOTICE_TEXT = (
    '내엑셀 이용 조건 및 공개 소스 안내\n'
    '개인·기업·공공기관의 일반 업무 사용은 무료입니다. 공개 소스는 열람·분석하고 본인의 업무를 위해 수정·빌드할 수 있습니다.\n'
    '원본을 무료로 공유할 때에는 출처와 이용 조건을 유지해야 합니다. 수정본의 외부 배포에는 권한 보유자의 별도 허락이 필요합니다.\n'
    '프로그램과 소스의 판매·유료 재배포·상품화 및 유료 제품·서비스에 포함하는 행위는 금지됩니다.\n'
    '외부 AI 서비스를 이용할 때에는 실사용 문서·개인정보·내부 자료와 소속 조직의 정보보안 정책을 확인해 주세요.\n'
    '이 문구는 이용 고지이며, 파일 열람 또는 AI 분석을 기술적으로 차단하는 장치는 아닙니다.\n'
    '사용기한과 보호장치의 임의 수정, 무력화 및 우회는 금지됩니다. AI나 자동 분석·편집 도구를 통한 변경에도 동일하게 적용됩니다.\n'
    '정확한 이용 범위는 함께 제공하는 LICENSE 또는 README의 이용 조건을 확인해 주세요.'
)


def _block(newline: str, text: str = NOTICE_TEXT) -> str:
    return newline.join("' " + line for line in
                        [MARKER_BEGIN, *text.splitlines(), MARKER_END]) + newline


def without_vba_notice(source: str) -> str:
    # Remove only the exact block we own; never delete arbitrary comment/code ranges.
    clean = source
    for text in (NOTICE_TEXT, LEGACY_PROTECTED_NOTICE_TEXT, LEGACY_NOTICE_TEXT):
        for newline in ('\r\n', '\n'):
            clean = clean.replace(_block(newline, text), '')
    for marker in (MARKER_BEGIN, MARKER_END):
        if re.search(r"^' " + re.escape(marker) + r'\r?$', clean, re.MULTILINE):
            raise ValueError('unexpected distribution notice marker; refusing source rewrite')
    return clean


def with_vba_notice(source: str) -> str:
    source = without_vba_notice(source)
    newline = '\r\n' if '\r\n' in source else '\n'
    offset = 0
    for line in source.splitlines(keepends=True):
        if not line.endswith(('\r', '\n')):
            break
        if line.strip() and not line.startswith('Attribute '):
            break
        offset += len(line)
    return source[:offset] + _block(newline) + source[offset:]


def signature_parts(path: Path) -> list[str]:
    with zipfile.ZipFile(path) as archive:
        return [name for name in archive.namelist()
                if 'vbaprojectsignature' in name.lower() or name.lower().startswith('_xmlsignatures/')]


def audit_distribution(path: Path) -> dict[str, object]:
    before = hashlib.sha256(path.read_bytes()).hexdigest()
    with ExcelFile(path) as workbook:
        project = workbook.vba_project()
        modules = workbook.module_names()
        readable = characters = notified = comments = 0
        for name in modules:
            source = workbook.get_module(name)
            readable += bool(source.strip())
            characters += len(source)
            comments += sum(line.lstrip().startswith("'") for line in source.splitlines())
            notified += source == with_vba_notice(source)
        password_present = bool(project.protection and project.protection.has_password)
    with zipfile.ZipFile(path) as archive:
        names = archive.namelist()
        custom = ET.fromstring(archive.read('docProps/custom.xml')) if 'docProps/custom.xml' in names else []
        props = {element.attrib.get('name'): ''.join(element.itertext()) for element in custom}
        core = ET.fromstring(archive.read('docProps/core.xml')) if 'docProps/core.xml' in names else None
        description = core.find('{http://purl.org/dc/elements/1.1/}description') if core is not None else None
        web_parts = [name for name in names if re.fullmatch(r'xl/webextensions/webextension\d+\.xml', name)]
        web_ids = []
        for name in web_parts:
            root = ET.fromstring(archive.read(name))
            ref = root.find('{http://schemas.microsoft.com/office/webextensions/webextension/2010/11}reference')
            if ref is not None:
                web_ids.append(ref.attrib.get('id', ''))
        result = {
            'evidence_class': 'static-local-read-only', 'sha256': before,
            'module_count': len(modules), 'readable_module_count': readable,
            'source_character_count': characters, 'comment_line_count': comments,
            'notified_module_count': notified, 'vba_password_present': password_present,
            'custom_notice_present': props.get('SourceAnalysisNotice') == NOTICE_TEXT,
            'core_notice_present': description is not None and NOTICE_TEXT in (description.text or ''),
            'legacy_ai_stop_instruction_present': 'AIStopInstruction' in props,
            'signature_parts': signature_parts(path),
            'web_extension_count': len(web_parts), 'web_extension_store_ids': web_ids,
            'ribbon_part_count': sum(name.startswith('customUI/') and name.endswith('.xml') for name in names),
            'zip_integrity_ok': archive.testzip() is None,
            'ai_detection_or_blocking': False,
        }
    result['file_unchanged'] = before == hashlib.sha256(path.read_bytes()).hexdigest()
    if not result['file_unchanged']:
        raise RuntimeError('input changed during read-only audit')
    return result


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('xlam', type=Path)
    args = parser.parse_args()
    print(json.dumps(audit_distribution(args.xlam), ensure_ascii=True, indent=2))


if __name__ == '__main__':
    main()
