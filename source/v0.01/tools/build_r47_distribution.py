#!/usr/bin/env python3
from __future__ import annotations

import argparse
import hashlib
import json
import re
import shutil
import tempfile
import zipfile
from datetime import date
from pathlib import Path
from xml.etree import ElementTree as ET

from pyopenvba import ExcelFile
from pyopenvba.cfb import CFB
from pyopenvba.vba import VBAModuleKind
from distribution_notice import NOTICE_TEXT, audit_distribution, signature_parts, with_vba_notice
from ooxml_webextensions import clean_file
from xlam_package_order import normalize as normalize_package_order
from distribution_source_profile import transform_source
from distribution_release_audit import audit_release
from distribution_gate import signed_policy, gate_source, gate_ribbon
from distribution_private_names import compact_private_names, TOKEN as VBA_TOKEN


def find_project_root(start: Path) -> Path:
    for path in (start, *start.parents):
        if (path / "source/v0.01/contracts").is_dir() and (path / "development/scripts").is_dir():
            return path
        if (path / "01_완성본").exists() and (path / "source").exists():
            return path
    raise RuntimeError(f"project root not found from {start}")


ROOT = find_project_root(Path(__file__).resolve())
VERSION_LABEL = "v0.01_r105"
SOURCE_XLAM = ROOT / "01_완성본" / f"내엑셀 {VERSION_LABEL}.xlam"
DISTRIBUTION_DIR = ROOT / "02_배포본(고도화)"
DISTRIBUTION_XLAM = DISTRIBUTION_DIR / f"내엑셀 {VERSION_LABEL}_배포본.xlam"
REPORT_PATH = ROOT / "source" / "v0.01" / "reports" / "r105-distribution-policy-build.json"

CORE_NS = {
    "cp": "http://schemas.openxmlformats.org/package/2006/metadata/core-properties",
    "dc": "http://purl.org/dc/elements/1.1/",
    "dcterms": "http://purl.org/dc/terms/",
    "dcmitype": "http://purl.org/dc/dcmitype/",
    "xsi": "http://www.w3.org/2001/XMLSchema-instance",
}
CUSTOM_PROPS_NS = "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties"
VT_NS = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"
RELS_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
CONTENT_TYPES_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
CUSTOM_PROPS_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.custom-properties+xml"
CUSTOM_PROPS_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties"


for prefix, uri in CORE_NS.items():
    ET.register_namespace(prefix, uri)
ET.register_namespace("", CUSTOM_PROPS_NS)
ET.register_namespace("vt", VT_NS)


def sha256_bytes(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def sha256_file(path: Path) -> str:
    return sha256_bytes(path.read_bytes())


def read_zip_entry(path: Path, entry: str) -> bytes:
    with zipfile.ZipFile(path) as archive:
        return archive.read(entry)


def rewrite_zip(path: Path, replacements: dict[str, bytes], removals: set[str] | None = None) -> None:
    removals = removals or set()
    # Keep the temporary archive beside the destination.  The project lives
    # on a shared UNC volume while Python's default temp directory is local;
    # os.replace cannot atomically cross those volumes (WinError 17).
    with tempfile.NamedTemporaryFile(delete=False, suffix=path.suffix, dir=path.parent) as handle:
        tmp = Path(handle.name)
    try:
        with zipfile.ZipFile(path, "r") as src, zipfile.ZipFile(tmp, "w") as dst:
            seen: set[str] = set()
            for info in src.infolist():
                if info.filename in seen:
                    continue
                if info.filename in removals:
                    seen.add(info.filename)
                    continue
                seen.add(info.filename)
                data = replacements.get(info.filename, src.read(info.filename))
                out_info = zipfile.ZipInfo(info.filename, info.date_time)
                out_info.compress_type = info.compress_type
                out_info.external_attr = info.external_attr
                out_info.create_system = info.create_system
                dst.writestr(out_info, data)
            for name, data in replacements.items():
                if name in seen:
                    continue
                info = zipfile.ZipInfo(name)
                info.compress_type = zipfile.ZIP_DEFLATED
                dst.writestr(info, data)
        tmp.replace(path)
    finally:
        if tmp.exists():
            tmp.unlink()


def project_stream_from_xlam(path: Path) -> bytes:
    cfb = CFB.from_bytes(read_zip_entry(path, "xl/vbaProject.bin"))
    return cfb.get_stream("PROJECT")


def extract_protection(path: Path) -> dict[str, str]:
    text = project_stream_from_xlam(path).decode("latin-1", errors="replace")
    fields: dict[str, str] = {}
    for key in ("ID", "CMG", "DPB", "GC"):
        match = re.search(rf'^{key}="([^"]*)"\r?$', text, re.MULTILINE)
        if not match:
            raise RuntimeError(f"PROJECT protection field missing: {key}")
        fields[key] = match.group(1)
    if len(fields["DPB"]) <= 16:
        raise RuntimeError("existing distribution does not contain a password DPB block")
    return fields


def replace_project_field(text: str, key: str, value: str) -> str:
    pattern = rf'^{key}="[^"]*"\r?$'
    replacement = f'{key}="{value}"'
    new_text, count = re.subn(pattern, replacement, text, count=1, flags=re.MULTILINE)
    if count != 1:
        raise RuntimeError(f"could not replace PROJECT field: {key}")
    return new_text


def apply_project_protection(path: Path, protection: dict[str, str]) -> None:
    vba_bin = read_zip_entry(path, "xl/vbaProject.bin")
    cfb = CFB.from_bytes(vba_bin)
    text = cfb.get_stream("PROJECT").decode("latin-1", errors="replace")
    for key in ("ID", "CMG", "DPB", "GC"):
        text = replace_project_field(text, key, protection[key])
    cfb.write_stream("PROJECT", text.encode("latin-1"))
    rewrite_zip(path, {"xl/vbaProject.bin": cfb.to_bytes()})


def distribution_policy_module(block_from: date, valid_until: date) -> str:
    distribution_title = f"내엑셀 {VERSION_LABEL} 배포본"
    return rf'''Attribute VB_Name = "mdLHEDistributionPolicy"
Option Explicit

Private Const LHE_DISTRIBUTION_BLOCK_YEAR As Long = {block_from.year}
Private Const LHE_DISTRIBUTION_BLOCK_MONTH As Long = {block_from.month}
Private Const LHE_DISTRIBUTION_BLOCK_DAY As Long = {block_from.day}
Private Const LHE_DISTRIBUTION_VALID_UNTIL As String = "{valid_until:%Y-%m-%d}"
Private Const LHE_BIRTHDAY_MONTH As Long = 0
Private Const LHE_BIRTHDAY_DAY As Long = 0
Private Const LHE_BIRTHDAY_MESSAGE As String = ""
Private Const LHE_DISTRIBUTION_NOTICE_FILE As String = "distribution-notice-v1.cfg"

Public Function LHEDistributionPolicyCanOpen() As Boolean
    On Error GoTo FailClosed
    If NxDistributionCanExecute() Then
        LHEDistributionPolicyCanOpen = True
        Exit Function
    End If
    MsgBox NxDistributionMessage(), _
           vbExclamation, "{distribution_title}"
    LHEDistributionPolicyCanOpen = False
    Exit Function

FailClosed:
    LHEDistributionPolicyCanOpen = False
End Function

Public Sub LHEDistributionPolicyShowBirthdayIfNeeded()
    On Error GoTo CleanExit
    If Month(Date) <> LHE_BIRTHDAY_MONTH Then Exit Sub
    If Day(Date) <> LHE_BIRTHDAY_DAY Then Exit Sub

    Dim cfgPath As String, thisYear As String, lastShownYear As String
    cfgPath = NxLHexcelProfileRoot() & "\Settings\" & LHE_DISTRIBUTION_NOTICE_FILE
    thisYear = CStr(Year(Date))
    lastShownYear = LHEDistributionPolicyReadYear(cfgPath)
    ' Preserve notices recorded by releases with the missing Settings separator.
    If Len(lastShownYear) = 0 Then lastShownYear = LHEDistributionPolicyReadYear( _
        NxLHexcelProfileRoot() & "\Settings" & LHE_DISTRIBUTION_NOTICE_FILE)
    If StrComp(lastShownYear, thisYear, vbTextCompare) = 0 Then Exit Sub

    MsgBox LHE_BIRTHDAY_MESSAGE, vbInformation, "내엑셀"
    LHEDistributionPolicyWriteYear cfgPath, thisYear
CleanExit:
End Sub

Private Function LHEDistributionPolicyReadYear(ByVal cfgPath As String) As String
    Dim handle As Integer, value As String
    On Error GoTo CleanExit
    If Len(Dir$(cfgPath)) = 0 Then Exit Function
    handle = FreeFile
    Open cfgPath For Input Access Read As #handle
    If Not EOF(handle) Then Line Input #handle, value
    LHEDistributionPolicyReadYear = Trim$(value)
CleanExit:
    On Error Resume Next
    If handle > 0 Then Close #handle
End Function

Private Sub LHEDistributionPolicyWriteYear(ByVal cfgPath As String, ByVal value As String)
    Dim root As String, settingsFolder As String, temporary As String, handle As Integer
    On Error GoTo CleanExit
    root = NxLHexcelProfileRoot()
    settingsFolder = root & "\Settings"
    If Len(Dir$(root, vbDirectory)) = 0 Then MkDir root
    If Len(Dir$(settingsFolder, vbDirectory)) = 0 Then MkDir settingsFolder
    temporary = cfgPath & ".tmp"
    If Len(Dir$(temporary)) > 0 Then Kill temporary
    handle = FreeFile
    Open temporary For Output Access Write Lock Read Write As #handle
    Print #handle, value
    Close #handle
    handle = 0
    If Len(Dir$(cfgPath)) > 0 Then Kill cfgPath
    Name temporary As cfgPath
CleanExit:
    On Error Resume Next
    If handle > 0 Then Close #handle
    If Len(temporary) > 0 And Len(Dir$(temporary)) > 0 Then Kill temporary
End Sub
'''


def patch_workbook_open(source: str) -> str:
    replacement = '''Private Sub Workbook_Open()
    On Error Resume Next
    NxDistributionStart
    If Not LHEDistributionPolicyCanOpen() Then
        Exit Sub
    End If
    LHEDistributionPolicyShowBirthdayIfNeeded
    NxShortcutsApplySavedBindings
    On Error GoTo 0
End Sub'''
    pattern = r"Private Sub Workbook_Open\(\).*?End Sub"
    patched, count = re.subn(pattern, replacement, source, count=1, flags=re.DOTALL)
    if count == 1:
        return patched
    # The r105 XLAM source may not yet have a document-event procedure.
    # Add the guarded entry point to the ThisWorkbook module instead of
    # silently omitting the expiry/birthday policy.
    return source.rstrip() + "\n\n" + replacement + "\n"


def workbook_module_name(path: Path) -> str:
    with zipfile.ZipFile(path) as archive:
        root = ET.fromstring(archive.read('xl/workbook.xml'))
    properties = root.find('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}workbookPr')
    name = properties.get('codeName') if properties is not None else None
    if not name:
        raise RuntimeError('Workbook codeName is required for distribution event policy')
    return name


def patch_workbook_close(source: str) -> str:
    signature = "Private Sub Workbook_BeforeClose(Cancel As Boolean)"
    pattern = r"(?im)^Private Sub Workbook_BeforeClose\([^\r\n]*\)\s*\r?\n"
    if re.search(pattern, source):
        return re.sub(pattern, lambda m: m.group(0) + "    NxDistributionStop\n", source, count=1)
    return source.rstrip() + "\n\n" + signature + "\n    NxDistributionStop\nEnd Sub\n"


def inject_vba_policy(path: Path, block_from: date, valid_until: date, profile: str = "internal-xlam", host_hashes=None) -> dict[str, object]:
    if signature_parts(path):
        raise RuntimeError('signed input requires a separate re-signing workflow')
    document_module = workbook_module_name(path)
    signed = signed_policy(VERSION_LABEL, profile, block_from)
    with ExcelFile(path) as workbook:
        project = workbook.vba_project()
        before_modules = project.module_names()
        if "NxDistributionGate" not in before_modules:
            raise RuntimeError("Distribution gate source is required; rebuild the product first")
        workbook.set_module("NxDistributionGate", gate_source(workbook.get_module("NxDistributionGate"), signed))
        about = workbook.get_module("NxProductAbout")
        about, version_count = re.subn(r'(NxProductVersionText\s*=\s*)"[^"\r\n]*"', lambda m: m.group(1) + '"내엑셀 ' + VERSION_LABEL + '"', about)
        about, build_count = re.subn(r'(NxProductBuildIdText\s*=\s*)"[^"\r\n]*"', lambda m: m.group(1) + '"' + VERSION_LABEL.split("_")[-1] + ' · 내부망/DLL 공통 기능"', about)
        if version_count != 1 or build_count != 1:
            raise RuntimeError("Expected one product version and build label")
        workbook.set_module("NxProductAbout", about)
        module_source = distribution_policy_module(block_from, valid_until)
        if "mdLHEDistributionPolicy" in before_modules:
            workbook.set_module("mdLHEDistributionPolicy", module_source)
            module_action = "updated"
        else:
            project.add_module(
                "mdLHEDistributionPolicy",
                module_source,
                kind=VBAModuleKind.standard,
            )
            module_action = "added"

        current = workbook.get_module(document_module)
        workbook.set_module(document_module, patch_workbook_close(patch_workbook_open(current)))
        literals = "\n".join(
            token.group() for module_name in workbook.module_names()
            for token in VBA_TOKEN.finditer(workbook.get_module(module_name))
            if token.group().startswith('"')
        )
        standard_names = {p.stem for p in (Path(__file__).resolve().parents[1] / "src/vba").rglob("*.bas")}
        compacted_identifiers = 0
        for name in workbook.module_names():
            current = workbook.get_module(name)
            transformed = transform_source(name, current, profile, host_hashes)
            # Keep the small policy and host-verification modules readable for
            # explicit inclusion/audit checks. Only implementation helpers shrink.
            if name in standard_names and name not in {"NxDistributionGate", "NxHostIntegrity", "NxWorkbookCompare"}:
                transformed, aliases = compact_private_names(transformed, literals)
                compacted_identifiers += len(aliases)
            notified = with_vba_notice(transformed)
            if notified != current:
                workbook.set_module(name, notified)
        validation = project.validate()
        workbook.save()

    with zipfile.ZipFile(path) as archive:
        ribbon_updates = {
            name: gate_ribbon(archive.read(name))
            for name in archive.namelist()
            if name.startswith("customUI/") and name.endswith(".xml")
        }
    rewrite_zip(path, ribbon_updates)
    with ExcelFile(path) as workbook:
        modules = workbook.module_names()
        policy_module = workbook.get_module("mdLHEDistributionPolicy")
        thisworkbook = workbook.get_module(document_module)

    return {
        "module_action": module_action,
        "private_identifiers_compacted": compacted_identifiers,
        "module_count_before": len(before_modules),
        "module_count_after": len(modules),
        "validation_before_save": validation,
        "policy_module_present": "mdLHEDistributionPolicy" in modules,
        "workbook_open_calls_can_open": "LHEDistributionPolicyCanOpen" in thisworkbook,
        "workbook_open_calls_birthday": "LHEDistributionPolicyShowBirthdayIfNeeded" in thisworkbook,
        "expiry_block_expression_present": f"DateSerial({block_from.year}, {block_from.month}, {block_from.day})" in policy_module
        or f"LHE_DISTRIBUTION_BLOCK_YEAR As Long = {block_from.year}" in policy_module,
        "birthday_message_present": 'LHE_BIRTHDAY_MONTH As Long = 0' in policy_module and 'LHE_BIRTHDAY_DAY As Long = 0' in policy_module and 'LHE_BIRTHDAY_MESSAGE As String = ""' in policy_module,
    }


def clean_workbook_xml(text: str) -> tuple[str, dict[str, object]]:
    before_ignorable = re.search(r'mc:Ignorable="([^"]*)"', text)
    new_text, removed_alt = re.subn(
        r"<mc:AlternateContent\b[^>]*>.*?<x15ac:absPath\b[^>]*/>.*?</mc:AlternateContent>",
        "",
        text,
        flags=re.DOTALL,
    )
    new_text, removed_direct = re.subn(r"<x15ac:absPath\b[^>]*/>", "", new_text)
    new_text = re.sub(r'\s+xmlns:xr="[^"]*"', "", new_text)
    new_text = re.sub(r'\s+xmlns:xr6="[^"]*"', "", new_text)
    new_text = re.sub(r'\s+xmlns:xr10="[^"]*"', "", new_text)
    new_text = re.sub(r'mc:Ignorable="[^"]*"', 'mc:Ignorable="x15 xr2"', new_text, count=1)
    after_ignorable = re.search(r'mc:Ignorable="([^"]*)"', new_text)
    return new_text, {
        "removed_absPath_blocks": removed_alt + removed_direct,
        "mc_ignorable_before": before_ignorable.group(1) if before_ignorable else "",
        "mc_ignorable_after": after_ignorable.group(1) if after_ignorable else "",
    }


def set_core_child(root: ET.Element, qname: str, value: str, attrs: dict[str, str] | None = None) -> None:
    child = root.find(qname)
    if child is None:
        child = ET.SubElement(root, qname)
    child.text = value
    if attrs:
        for key, attr_value in attrs.items():
            child.set(key, attr_value)


def build_core_xml(existing: bytes, block_from: date, valid_until: date) -> bytes:
    root = ET.fromstring(existing)
    set_core_child(root, f"{{{CORE_NS['dc']}}}title", f"내엑셀 {VERSION_LABEL} 배포본")
    set_core_child(root, f"{{{CORE_NS['dc']}}}subject", f"내엑셀 {VERSION_LABEL} 보안 배포본")
    set_core_child(
        root,
        f"{{{CORE_NS['dc']}}}description",
        f"내엑셀 {VERSION_LABEL} 배포본. {block_from:%Y-%m-%d}부터 실행 차단. 유효 사용일 {valid_until:%Y-%m-%d}까지.\n\n" + NOTICE_TEXT,
    )
    set_core_child(
        root,
        f"{{{CORE_NS['cp']}}}keywords",
        f"내엑셀, {VERSION_LABEL}, 배포본, 소스공개, 일반업무무료, 유료재배포금지, VBA보호, {block_from:%Y-%m-%d}차단",
    )
    now = date.today().isoformat() + "T00:00:00Z"
    set_core_child(
        root,
        f"{{{CORE_NS['dcterms']}}}modified",
        now,
        {f"{{{CORE_NS['xsi']}}}type": "dcterms:W3CDTF"},
    )
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def build_custom_props(block_from: date, valid_until: date, profile: str = "internal-xlam") -> bytes:
    root = ET.Element(f"{{{CUSTOM_PROPS_NS}}}Properties")
    props = {
        "DistributionProfile": profile,
        "CompiledOnlyFeatures": "NX-FILE-WORKBOOK-COMPARE,NX-HANGUL-TABLE-SEND,NX-HANGUL-PICTURE-SEND" if profile == "enhanced-dll" else "",
        "DistributionPolicyEnabled": "true",
        "DistributionVersion": f"내엑셀 {VERSION_LABEL}_배포본",
        "DistributionValidUntil": valid_until.isoformat(),
        "DistributionBlockFrom": block_from.isoformat(),
        "DistributionExpiryPolicy": f"본 배포본은 {block_from:%Y-%m-%d}부터 실행을 차단합니다.",
        "BirthdayNoticeEnabled": "false",
        "BirthdayNoticeRule": "",
        "BirthdayNoticeText": "",
        "AIAnalysisAndCommercialUsePolicy": "개인·기업·공공기관의 일반 업무 사용은 무료입니다. 공개 소스의 열람·분석과 본인의 업무를 위한 수정·빌드를 허용합니다. 실사용 자료의 외부 전달에는 소속 조직의 정보보안 정책을 적용합니다.",
        "SourceAnalysisNotice": NOTICE_TEXT,
        "CommercialUseRestriction": "프로그램과 소스의 판매·유료 재배포·상품화 및 유료 제품·서비스에 포함하는 행위를 금지합니다. 원본 무료 공유는 출처와 이용 조건 유지 시 허용하며 수정본의 외부 배포에는 별도 허락이 필요합니다.",
        "ReverseEngineeringPolicy": "공개 소스의 열람·분석과 본인의 업무를 위한 수정·빌드를 허용합니다. 사용기한·보호장치의 임의 수정·무력화·우회 및 이용 조건에 반하는 재배포는 금지합니다.",
        "DistributionFolderRule": f"02_배포본(고도화)의 내엑셀 {VERSION_LABEL} 내부망 배포본 및 DLL 배포본 폴더에 각각 배치합니다.",
        "EnglishPolicyNotice": "Free for ordinary work use. Source viewing, analysis and private work builds are permitted. Preserve attribution and terms when sharing unmodified copies for free. Modified external distribution requires permission. Sale, paid redistribution, commercialization and inclusion in paid products or services are prohibited. Do not bypass expiry or protection. See LICENSE.",
    }
    for pid, (name, value) in enumerate(props.items(), start=2):
        prop = ET.SubElement(
            root,
            f"{{{CUSTOM_PROPS_NS}}}property",
            {
                "fmtid": "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}",
                "pid": str(pid),
                "name": name,
            },
        )
        child = ET.SubElement(prop, f"{{{VT_NS}}}lpwstr")
        child.text = value
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def add_custom_props_relationship(rels_xml: bytes) -> tuple[bytes, bool]:
    root = ET.fromstring(rels_xml)
    for child in root:
        if child.attrib.get("Type") == CUSTOM_PROPS_REL_TYPE:
            child.set("Target", "docProps/custom.xml")
            return ET.tostring(root, encoding="utf-8", xml_declaration=True), False
    ids = []
    for child in root:
        rid = child.attrib.get("Id", "")
        match = re.fullmatch(r"rId(\d+)", rid)
        if match:
            ids.append(int(match.group(1)))
    ET.SubElement(
        root,
        f"{{{RELS_NS}}}Relationship",
        {
            "Id": f"rId{max(ids, default=0) + 1}",
            "Type": CUSTOM_PROPS_REL_TYPE,
            "Target": "docProps/custom.xml",
        },
    )
    return ET.tostring(root, encoding="utf-8", xml_declaration=True), True


def add_custom_props_content_type(content_types_xml: bytes) -> tuple[bytes, bool]:
    root = ET.fromstring(content_types_xml)
    for child in root:
        if child.attrib.get("PartName") == "/docProps/custom.xml":
            child.set("ContentType", CUSTOM_PROPS_CONTENT_TYPE)
            return ET.tostring(root, encoding="utf-8", xml_declaration=True), False
    ET.SubElement(
        root,
        f"{{{CONTENT_TYPES_NS}}}Override",
        {
            "PartName": "/docProps/custom.xml",
            "ContentType": CUSTOM_PROPS_CONTENT_TYPE,
        },
    )
    return ET.tostring(root, encoding="utf-8", xml_declaration=True), True


def remove_calc_chain_relationship(workbook_rels_xml: bytes) -> tuple[bytes, bool]:
    root = ET.fromstring(workbook_rels_xml)
    removed = False
    for child in list(root):
        if child.attrib.get("Target") == "calcChain.xml" or child.attrib.get("Type", "").endswith("/calcChain"):
            root.remove(child)
            removed = True
    return ET.tostring(root, encoding="utf-8", xml_declaration=True), removed


def remove_calc_chain_content_type(content_types_xml: bytes) -> tuple[bytes, bool]:
    root = ET.fromstring(content_types_xml)
    removed = False
    for child in list(root):
        if child.attrib.get("PartName") == "/xl/calcChain.xml":
            root.remove(child)
            removed = True
    return ET.tostring(root, encoding="utf-8", xml_declaration=True), removed


def apply_package_policy(path: Path, block_from: date, valid_until: date, profile: str = "internal-xlam") -> dict[str, object]:
    workbook_xml = read_zip_entry(path, "xl/workbook.xml").decode("utf-8", errors="replace")
    workbook_xml, workbook_cleanup = clean_workbook_xml(workbook_xml)
    rels_xml, custom_relationship_added = add_custom_props_relationship(read_zip_entry(path, "_rels/.rels"))
    workbook_rels_xml, calc_chain_relationship_removed = remove_calc_chain_relationship(
        read_zip_entry(path, "xl/_rels/workbook.xml.rels")
    )
    content_types_xml, calc_chain_content_type_removed = remove_calc_chain_content_type(
        read_zip_entry(path, "[Content_Types].xml")
    )
    content_types_xml, custom_content_type_added = add_custom_props_content_type(content_types_xml)
    replacements = {
        "xl/workbook.xml": workbook_xml.encode("utf-8"),
        "xl/_rels/workbook.xml.rels": workbook_rels_xml,
        "docProps/core.xml": build_core_xml(read_zip_entry(path, "docProps/core.xml"), block_from, valid_until),
        "docProps/custom.xml": build_custom_props(block_from, valid_until, profile),
        "_rels/.rels": rels_xml,
        "[Content_Types].xml": content_types_xml,
    }
    rewrite_zip(path, replacements, removals={"xl/calcChain.xml"})
    return {
        "custom_relationship_added": custom_relationship_added,
        "custom_content_type_added": custom_content_type_added,
        "calc_chain_removed": calc_chain_relationship_removed or calc_chain_content_type_removed,
        "calc_chain_relationship_removed": calc_chain_relationship_removed,
        "calc_chain_content_type_removed": calc_chain_content_type_removed,
        "workbook_cleanup": workbook_cleanup,
    }


def distribution_dir_files() -> list[str]:
    return sorted(child.name for child in DISTRIBUTION_DIR.iterdir() if child.name != ".DS_Store")


def verify_policy(path: Path, block_from: date) -> dict[str, object]:
    document_module = workbook_module_name(path)
    with ExcelFile(path) as workbook:
        project = workbook.vba_project()
        policy = workbook.get_module("mdLHEDistributionPolicy")
        thisworkbook = workbook.get_module(document_module)
        protection = project.protection
        validation = project.validate()
    with zipfile.ZipFile(path) as archive:
        names = set(archive.namelist())
        workbook_xml = archive.read("xl/workbook.xml").decode("utf-8", errors="replace")
        custom_xml = archive.read("docProps/custom.xml").decode("utf-8", errors="replace")
        joined_xml = "\n".join(
            archive.read(name).decode("utf-8", errors="ignore")
            for name in names
            if name.endswith((".xml", ".rels"))
        )
    return {
        "vba_has_password": bool(protection and protection.has_password),
        "dpb_hex_length": len(protection.dpb) if protection else 0,
        "vba_validation": validation,
        "distribution_module_present": "LHE_DISTRIBUTION_BLOCK_YEAR" in policy,
        "block_from_constant_present": f"LHE_DISTRIBUTION_BLOCK_YEAR As Long = {block_from.year}" in policy
        and f"LHE_DISTRIBUTION_BLOCK_MONTH As Long = {block_from.month}" in policy
        and f"LHE_DISTRIBUTION_BLOCK_DAY As Long = {block_from.day}" in policy,
        "birthday_message_present": 'LHE_BIRTHDAY_MONTH As Long = 0' in policy and 'LHE_BIRTHDAY_DAY As Long = 0' in policy and 'LHE_BIRTHDAY_MESSAGE As String = ""' in policy,
        "workbook_open_guard_present": "LHEDistributionPolicyCanOpen" in thisworkbook
        and "LHEDistributionPolicyShowBirthdayIfNeeded" in thisworkbook,
        "custom_props_present": "docProps/custom.xml" in names,
        "custom_props_block_from_present": block_from.isoformat() in custom_xml,
        "custom_props_birthday_present": False,
        "abs_path_present": "absPath" in workbook_xml,
        "local_path_present": "Users/woo" in joined_xml or "file://" in joined_xml,
    }


def parse_date(value: str) -> date:
    return date.fromisoformat(value)


def resolve_project_path(path: Path) -> Path:
    return path if path.is_absolute() else ROOT / path


def main() -> int:
    global VERSION_LABEL, SOURCE_XLAM, DISTRIBUTION_XLAM, DISTRIBUTION_DIR, REPORT_PATH

    parser = argparse.ArgumentParser()
    parser.add_argument("--version-label", default=VERSION_LABEL)
    parser.add_argument("--source-xlam", type=Path)
    parser.add_argument("--destination-xlam", type=Path)
    parser.add_argument("--protection-seed", type=Path)
    parser.add_argument("--block-from", default="2028-01-01")
    parser.add_argument("--valid-until", default="2027-12-31")
    parser.add_argument("--json", type=Path)
    parser.add_argument("--profile", choices=("internal-xlam", "enhanced-dll"), default="internal-xlam")
    args = parser.parse_args()

    VERSION_LABEL = args.version_label
    version_token = VERSION_LABEL.replace(".", "")
    SOURCE_XLAM = resolve_project_path(args.source_xlam) if args.source_xlam else ROOT / "01_완성본" / f"내엑셀 {VERSION_LABEL}.xlam"
    DISTRIBUTION_XLAM = (
        resolve_project_path(args.destination_xlam)
        if args.destination_xlam
        else DISTRIBUTION_DIR / f"내엑셀 {VERSION_LABEL}_배포본.xlam"
    )
    DISTRIBUTION_DIR = DISTRIBUTION_XLAM.parent
    REPORT_PATH = (
        resolve_project_path(args.json)
        if args.json
        else ROOT / "source" / "v0.01" / "reports" / f"{version_token}-distribution-policy-build.json"
    )
    protection_seed = resolve_project_path(args.protection_seed) if args.protection_seed else DISTRIBUTION_XLAM

    block_from = parse_date(args.block_from)
    valid_until = parse_date(args.valid_until)
    if valid_until >= block_from:
        raise RuntimeError("valid-until must be earlier than block-from")
    if not SOURCE_XLAM.exists():
        raise FileNotFoundError(SOURCE_XLAM)

    source_sha = sha256_file(SOURCE_XLAM)
    host_hashes = None
    if args.profile == 'enhanced-dll':
        host_hashes = {arch: sha256_file(SOURCE_XLAM.parent / name) for arch, name in [('x86','NxHost32.dll'),('x64','NxHost64.dll')]}
    old_protection = extract_protection(protection_seed) if protection_seed.exists() else None
    if old_protection is None:
        raise RuntimeError(f"protected distribution workbook is required to preserve the VBA password block: {protection_seed}")

    DISTRIBUTION_DIR.mkdir(parents=True, exist_ok=True)
    if DISTRIBUTION_XLAM.exists():
        raise RuntimeError(f"distribution destination already exists: {DISTRIBUTION_XLAM}")
    with tempfile.NamedTemporaryFile(
        delete=False,
        suffix=DISTRIBUTION_XLAM.suffix,
        dir=DISTRIBUTION_DIR,
    ) as handle:
        temporary_destination = Path(handle.name)
    published = False
    try:
        shutil.copy2(SOURCE_XLAM, temporary_destination)
        vba_result = inject_vba_policy(temporary_destination, block_from, valid_until, args.profile, host_hashes)
        package_result = apply_package_policy(temporary_destination, block_from, valid_until, args.profile)
        package_result["removed_webextension_parts"] = clean_file(temporary_destination)
        apply_project_protection(temporary_destination, old_protection)
        normalize_package_order(temporary_destination)
        verify_result = verify_policy(temporary_destination, block_from)
        notice_result = audit_distribution(temporary_destination)
        release_audit = audit_release(temporary_destination, args.profile, host_hashes)

        failures: list[str] = []
        if notice_result['module_count'] == 0 or notice_result['notified_module_count'] != notice_result['module_count']:
            failures.append('distribution notice missing from VBA modules')
        if not notice_result['custom_notice_present'] or not notice_result['core_notice_present']:
            failures.append('distribution notice missing from document properties')
        if notice_result['legacy_ai_stop_instruction_present']:
            failures.append('misleading AI stop instruction remains')
        if verify_result["vba_validation"]:
            failures.append("vba project validation failed")
        if not verify_result["vba_has_password"]:
            failures.append("vba project password not detected")
        if not verify_result["distribution_module_present"]:
            failures.append("distribution policy module missing")
        if not verify_result["block_from_constant_present"]:
            failures.append("block-from constant missing")
        if not verify_result["birthday_message_present"]:
            failures.append("birthday message missing")
        if not verify_result["workbook_open_guard_present"]:
            failures.append("Workbook_Open policy guard missing")
        if not verify_result["custom_props_present"]:
            failures.append("custom properties missing")
        if not verify_result["custom_props_block_from_present"]:
            failures.append("custom props block-from missing")
        if verify_result["abs_path_present"]:
            failures.append("absPath remained in workbook.xml")
        if verify_result["local_path_present"]:
            failures.append("local path leaked into xml/rels payload")
        if not failures:
            temporary_destination.replace(DISTRIBUTION_XLAM)
            published = True
        verify_result["distribution_dir_files"] = distribution_dir_files()
        if published and DISTRIBUTION_XLAM.name not in verify_result["distribution_dir_files"]:
            failures.append("distribution folder is missing the current target")

        report = {
            "status": "pass" if not failures else "fail",
            "failures": failures,
            "source": str(SOURCE_XLAM),
            "profile": args.profile,
            "compiled_only_features": ["NX-FILE-WORKBOOK-COMPARE", "NX-HANGUL-TABLE-SEND", "NX-HANGUL-PICTURE-SEND"] if args.profile == "enhanced-dll" else [],
            "destination": str(DISTRIBUTION_XLAM),
            "protection_seed": str(protection_seed),
            "source_sha256": source_sha,
            "distribution_sha256": sha256_file(DISTRIBUTION_XLAM if published else temporary_destination),
            "source_size": SOURCE_XLAM.stat().st_size,
            "distribution_size": (DISTRIBUTION_XLAM if published else temporary_destination).stat().st_size,
            "deleted_existing_distribution_files": [],
            "published_atomically": published,
            "expiry": {
                "valid_until": valid_until.isoformat(),
                "block_from": block_from.isoformat(),
                "interpretation": f"{block_from.isoformat()}부터 차단, {valid_until.isoformat()}까지 사용 가능",
            },
            "vba": vba_result,
            "package": package_result,
            "protection": {
                "reused_existing_distribution_protection_block": True,
                "seed": str(protection_seed),
                "project_id": old_protection["ID"],
                "dpb_hex_length": len(old_protection["DPB"]),
                "protection_block_sha256": sha256_bytes(
                    ("\n".join(old_protection[key] for key in ("ID", "CMG", "DPB", "GC"))).encode("ascii")
                ),
                "password_plaintext_stored": False,
            },
            "verification": verify_result,
            "source_analysis_notice": notice_result,
            "release_audit": release_audit,
            "runtime_dll_hashes": host_hashes,
        }
        REPORT_PATH.parent.mkdir(parents=True, exist_ok=True)
        REPORT_PATH.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
        print(json.dumps(report, ensure_ascii=False, indent=2))
        return 0 if report["status"] == "pass" else 1
    finally:
        if temporary_destination.exists():
            temporary_destination.unlink()


if __name__ == "__main__":
    raise SystemExit(main())
