"""Reduce distribution-only source exposure without modifying development code."""
from __future__ import annotations
import re

COMPARE_MODULE = 'NxWorkbookCompare'
REQUIRED_HOST = '문서 비교 DLL을 사용할 수 없습니다. 고도화판 설치를 확인하거나 내부망용 배포본을 사용해 주세요.'

def compact_comments(source: str) -> str:
    # Preserve complete attribution blocks when present, including their license terms.
    if re.search(r"(?im)^\s*(?:'|Rem\s).*?(copyright|SPDX|license|저작권)", source):
        return source
    return ''.join(line for line in source.splitlines(keepends=True)
                   if not re.match(r"^\s*(?:'|Rem(?:\s|$))", line, re.I))

def enhanced_compare(source: str) -> str:
    # Keep the public API and Excel capture/report adapter. Remove the duplicate
    # decision algorithm from the enhanced distribution, retaining it internally.
    pattern = r'(?ims)^Private Function NxWorkbookCompareCellDifference\([^\r\n]*\r?\n.*?^End Function'
    matches = list(re.finditer(pattern, source))
    if len(matches) != 1:
        raise ValueError('Expected one VBA compare fallback')
    signature = matches[0].group(0).splitlines()[0]
    replacement = signature + '\n    NxRaiseContractError "' + REQUIRED_HOST + '"\nEnd Function'
    source = re.sub(pattern, lambda _: replacement, source)
    old = '    If service Is Nothing Then Exit Function'
    # Only the comparison factory uses this exact statement in this module.
    if source.count(old) != 1:
        raise ValueError('Expected one compare host availability guard')
    source = source.replace(old, '    If service Is Nothing Then NxRaiseContractError "' + REQUIRED_HOST + '"')
    append = r'(?ims)^Private Function NxWorkbookCompareAppendReason\([^\r\n]*\r?\n.*?^End Function'
    source, count = re.subn(append, '', source)
    if count != 1:
        raise ValueError('Expected one compare reason helper')
    return source

def enhanced_hwpx(source: str) -> str:
    pattern = r'(?ims)^    If Not NxHwpxTryNative\(payload, outputPath\) Then\r?\n.*?^    End If'
    if len(re.findall(pattern, source)) != 1:
        raise ValueError('Expected one HWPX fallback boundary')
    return re.sub(pattern, '    If Not NxHwpxTryNative(payload, outputPath) Then NxRaiseContractError NxHostIntegrityLastError()', source)

def enhanced_hwpx_pictures(source: str) -> str:
    pattern = r'(?ims)^    If Not NxHwpxTryNative\(payload, outputPath, 120\) Then\r?\n.*?^    End If'
    if len(re.findall(pattern, source)) != 1:
        raise ValueError('Expected one picture HWPX fallback boundary')
    return re.sub(pattern, '    If Not NxHwpxTryNative(payload, outputPath, 120) Then NxRaiseContractError "아래한글 그림 전송 DLL 서비스를 사용할 수 없습니다: " & NxHostIntegrityLastError()', source)


def enhanced_integrity(source: str, host_hashes: dict[str, str] | None) -> str:
    if not host_hashes or set(host_hashes) != {'x86', 'x64'} or any(not re.fullmatch('[0-9a-f]{64}', v) for v in host_hashes.values()):
        raise ValueError('Enhanced distribution requires exact x86/x64 DLL hashes')
    for old, new in [('NX_HOST_VERIFY_FILES As Boolean = False', 'NX_HOST_VERIFY_FILES As Boolean = True'),
                     *[(f'NX_HOST_SHA256_{arch.upper()} As String = ""', f'NX_HOST_SHA256_{arch.upper()} As String = "{host_hashes[arch]}"') for arch in ('x86', 'x64')]]:
        if source.count(old) != 1:raise ValueError('Host integrity source boundary changed')
        source = source.replace(old, new)
    return source


def transform_source(name: str, source: str, profile: str, host_hashes: dict[str, str] | None = None) -> str:
    if profile not in ('internal-xlam', 'enhanced-dll'):
        raise ValueError('Unknown distribution profile')
    if profile == 'enhanced-dll' and name == COMPARE_MODULE:
        source = enhanced_compare(source)
    if profile == 'enhanced-dll' and name == 'NxHostIntegrity':
        source = enhanced_integrity(source, host_hashes)
    if profile == 'enhanced-dll' and name == 'NxHwpxController':
        source = enhanced_hwpx(source)
    if profile == 'enhanced-dll' and name == 'NxHangulPictureTransfer':
        source = enhanced_hwpx_pictures(source)
    if profile == 'enhanced-dll' and name == 'NxPowerShell':
        pattern = r'(?ims)^(?:Public|Private) Function (NxPowerShellRunHwpx|RunPowerShellCommand)\([^\r\n]*\r?\n.*?^End Function'
        matches = list(re.finditer(pattern, source))
        if len(matches) != 2:raise ValueError('Expected HWPX script execution boundaries')
        source = re.sub(pattern, lambda match: match.group(0).splitlines()[0] + '\n    NxRaiseContractError "고도화판은 아래한글 DLL 서비스를 사용합니다."\nEnd Function', source)
    if profile == 'enhanced-dll' and name == 'NxHwpxEmbeddedResources':
        pattern = r'(?ims)^Public Function NxHwpxEmbeddedScriptBase64\(\) As String\r?\n.*?^End Function'
        if len(re.findall(pattern, source)) != 1:raise ValueError('Expected one embedded PowerShell resource')
        source = re.sub(pattern, 'Public Function NxHwpxEmbeddedScriptBase64() As String\n    NxRaiseContractError "고도화판은 아래한글 DLL 서비스를 사용합니다."\nEnd Function', source)
    return compact_comments(source)

def verify_enhanced_source(code: str) -> None:
    if REQUIRED_HOST not in code or re.search(r'If\s+CBool\(baseCell\.HasFormula\)|If\s+leftRows\(ordinal,\s*0\)|Private Function NxWorkbookCompareAppendReason', code, re.I):
        raise ValueError('Enhanced payload must use --profile enhanced-dll; VBA compare fallback remains or DLL guard is absent')


def verify_enhanced(path):
    from pyopenvba import ExcelFile
    with ExcelFile(path) as workbook:
        code = workbook.get_module(COMPARE_MODULE)
    verify_enhanced_source(code)

if __name__ == '__main__':
    import argparse
    from pathlib import Path
    parser=argparse.ArgumentParser()
    parser.add_argument('--source',type=Path,required=True)
    parser.add_argument('--destination',type=Path,required=True)
    parser.add_argument('--profile',choices=('internal-xlam','enhanced-dll'),required=True)
    args=parser.parse_args()
    if args.destination.exists():raise SystemExit('Destination exists')
    if args.source.suffix.lower() == '.xlam':
        import shutil
        import hashlib
        from pyopenvba import ExcelFile
        from distribution_notice import signature_parts
        if signature_parts(args.source):raise SystemExit('Signed source requires re-signing workflow')
        host_hashes = None
        if args.profile == 'enhanced-dll':
            host_hashes = {arch: hashlib.sha256((args.source.parent / filename).read_bytes()).hexdigest()
                           for arch, filename in [('x86', 'NxHost32.dll'), ('x64', 'NxHost64.dll')]}
        shutil.copy2(args.source,args.destination)
        with ExcelFile(args.destination) as workbook:
            for name in workbook.module_names():
                workbook.set_module(name,transform_source(name,workbook.get_module(name),args.profile,host_hashes))
            workbook.save()
    else:
        args.destination.write_text(transform_source(COMPARE_MODULE,args.source.read_text(encoding='utf-8-sig'),args.profile),encoding='utf-8')
