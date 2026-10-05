"""표 전용 HWPX의 최소 XML 파트를 생성한다."""

from __future__ import annotations

from collections.abc import Iterable
from xml.etree import ElementTree as ET

from .model import Cell, Merge, TableInput


NS = {
    "hv": "http://www.hancom.co.kr/hwpml/2011/version",
    "hs": "http://www.hancom.co.kr/hwpml/2011/section",
    "hp": "http://www.hancom.co.kr/hwpml/2011/paragraph",
    "hh": "http://www.hancom.co.kr/hwpml/2011/head",
    "hc": "http://www.hancom.co.kr/hwpml/2011/core",
    "opf": "http://www.idpf.org/2007/opf/",
    "odf": "urn:oasis:names:tc:opendocument:xmlns:manifest:1.0",
}
for prefix, uri in NS.items():
    ET.register_namespace(prefix, uri)


def q(prefix: str, tag: str) -> str:
    return f"{{{NS[prefix]}}}{tag}"


def _xml(root: ET.Element) -> bytes:
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def _sub(parent: ET.Element, prefix: str, tag: str, **attrs: object) -> ET.Element:
    return ET.SubElement(parent, q(prefix, tag), {key: str(value) for key, value in attrs.items()})


def normalize_widths(values: Iterable[float], total: int = 42520) -> tuple[int, ...]:
    raw = tuple(float(value) for value in values)
    denominator = sum(raw)
    exact = tuple(value * total / denominator for value in raw)
    widths = [int(value) for value in exact]
    remainder = total - sum(widths)
    order = sorted(range(len(raw)), key=lambda index: exact[index] - widths[index], reverse=True)
    for index in order[:remainder]:
        widths[index] += 1
    widths[-1] += total - sum(widths)
    return tuple(widths)


def _style_maps(table: TableInput) -> tuple[dict[tuple[bool, str], int], dict[tuple[str, str], int], dict[str, int]]:
    char_keys = sorted({(cell.font_bold, cell.font_rgb) for cell in table.cells})
    border_keys = sorted({(cell.fill_rgb, cell.borders) for cell in table.cells})
    alignments = sorted({cell.horizontal for cell in table.cells})
    return (
        {key: index + 1 for index, key in enumerate(char_keys)},
        {key: index + 3 for index, key in enumerate(border_keys)},
        {key: index + 1 for index, key in enumerate(alignments)},
    )


def version_xml() -> bytes:
    return _xml(ET.Element(q("hv", "HCFVersion"), {
        "tagetApplication": "WORDPROCESSOR", "major": "5", "minor": "1",
        "micro": "0", "buildNumber": "0", "os": "1", "xmlVersion": "1.4",
        "application": "LHexcel", "appVersion": "0.01",
    }))


def settings_xml() -> bytes:
    return _xml(ET.Element(q("hc", "settings"), {"version": "1.0"}))


def manifest_xml() -> bytes:
    root = ET.Element(q("odf", "manifest"), {q("odf", "version"): "1.0"})
    entries = (
        ("/", "application/hwp+zip"), ("version.xml", "application/xml"),
        ("settings.xml", "application/xml"), ("Contents/content.hpf", "application/oebps-package+xml"),
        ("Contents/header.xml", "application/xml"), ("Contents/section0.xml", "application/xml"),
        ("Preview/PrvText.txt", "text/plain"),
    )
    for path, media_type in entries:
        ET.SubElement(root, q("odf", "file-entry"), {
            q("odf", "full-path"): path, q("odf", "media-type"): media_type,
        })
    return _xml(root)


def content_hpf_xml() -> bytes:
    root = ET.Element(q("opf", "package"), {"version": "2.0", "unique-identifier": "bookid"})
    metadata = ET.SubElement(root, q("opf", "metadata"))
    ET.SubElement(metadata, q("opf", "meta"), {"name": "creator", "content": "LHexcel"})
    manifest = ET.SubElement(root, q("opf", "manifest"))
    ET.SubElement(manifest, q("opf", "item"), {"id": "header", "href": "Contents/header.xml", "media-type": "application/xml"})
    ET.SubElement(manifest, q("opf", "item"), {"id": "section0", "href": "Contents/section0.xml", "media-type": "application/xml"})
    spine = ET.SubElement(root, q("opf", "spine"))
    ET.SubElement(spine, q("opf", "itemref"), {"idref": "section0"})
    return _xml(root)


def header_xml(table: TableInput) -> bytes:
    char_map, border_map, align_map = _style_maps(table)
    root = ET.Element(q("hh", "head"), {"version": "1.4", "secCnt": "1"})
    _sub(root, "hh", "beginNum", page="1", footnote="1", endnote="1", pic="1", tbl="1", equation="1")
    ref = _sub(root, "hh", "refList")

    fontfaces = _sub(ref, "hh", "fontfaces", itemCnt="7")
    for language in ("HANGUL", "LATIN", "HANJA", "JAPANESE", "OTHER", "SYMBOL", "USER"):
        group = _sub(fontfaces, "hh", "fontface", lang=language, fontCnt="1")
        font = _sub(group, "hh", "font", id="0", face="중고딕", type="HFT", isEmbedded="0")
        _sub(font, "hh", "typeInfo", familyType="FCAT_GOTHIC", serifStyle="0", weight="6", proportion="4", contrast="0", strokeVariation="1", armStyle="1", letterform="1", midline="1", xHeight="1")

    border_fills = _sub(ref, "hh", "borderFills", itemCnt=str(2 + len(border_map)))
    for border_id in (1, 2):
        border = _sub(border_fills, "hh", "borderFill", id=str(border_id), threeD="0", shadow="0", centerLine="NONE", breakCellSeparateLine="0")
        for side in ("leftBorder", "rightBorder", "topBorder", "bottomBorder"):
            _sub(border, "hh", side, type="NONE", width="0.1 mm", color="#000000")
        _sub(border, "hh", "diagonal", type="SLASH", width="0.1 mm", color="#000000")
    for (fill_rgb, borders), border_id in border_map.items():
        border = _sub(border_fills, "hh", "borderFill", id=str(border_id), threeD="0", shadow="0", centerLine="NONE", breakCellSeparateLine="0")
        border_type = "SOLID" if borders == "all" else "NONE"
        for side in ("leftBorder", "rightBorder", "topBorder", "bottomBorder"):
            _sub(border, "hh", side, type=border_type, width="0.1 mm", color="#000000")
        _sub(border, "hh", "diagonal", type="SLASH", width="0.1 mm", color="#000000")
        fill_brush = _sub(border, "hc", "fillBrush")
        _sub(fill_brush, "hc", "winBrush", faceColor=f"#{fill_rgb}", hatchColor="#000000", alpha="0")

    char_properties = _sub(ref, "hh", "charProperties", itemCnt=str(1 + len(char_map)))
    for char_id, bold, color in [(0, False, "000000")] + [(ident, key[0], key[1]) for key, ident in char_map.items()]:
        char = _sub(char_properties, "hh", "charPr", id=str(char_id), height="1000", textColor=f"#{color}", shadeColor="none", useFontSpace="0", useKerning="0", symMark="NONE", borderFillIDRef="0")
        _sub(char, "hh", "fontRef", hangul="0", latin="0", hanja="0", japanese="0", other="0", symbol="0", user="0")
        _sub(char, "hh", "ratio", hangul="100", latin="100", hanja="100", japanese="100", other="100", symbol="100", user="100")
        _sub(char, "hh", "spacing", hangul="0", latin="0", hanja="0", japanese="0", other="0", symbol="0", user="0")
        _sub(char, "hh", "relSz", hangul="100", latin="100", hanja="100", japanese="100", other="100", symbol="100", user="100")
        _sub(char, "hh", "offset", hangul="0", latin="0", hanja="0", japanese="0", other="0", symbol="0", user="0")
        if bold:
            _sub(char, "hh", "bold")

    para_properties = _sub(ref, "hh", "paraProperties", itemCnt=str(1 + len(align_map)))
    for para_id, alignment in [(0, "left")] + [(ident, key) for key, ident in align_map.items()]:
        para = _sub(para_properties, "hh", "paraPr", id=str(para_id), tabPrIDRef="0", condense="0", fontLineHeight="0", snapToGrid="1", suppressLineNumbers="0", checked="0")
        _sub(para, "hh", "align", horizontal=alignment.upper(), vertical="BASELINE")
        _sub(para, "hh", "heading", type="NONE", idRef="0", level="0")
        _sub(para, "hh", "breakSetting", breakLatinWord="KEEP_WORD", breakNonLatinWord="KEEP_WORD", widowOrphan="0", keepWithNext="0", keepLines="0", pageBreakBefore="0", lineWrap="BREAK")
        _sub(para, "hh", "margin")
        _sub(para, "hh", "lineSpacing", type="PERCENT", value="160", unit="HWPUNIT")
        _sub(para, "hh", "border", borderFillIDRef="0", offsetLeft="0", offsetRight="0", offsetTop="0", offsetBottom="0", connect="0", ignoreMargin="0")

    tabs = _sub(ref, "hh", "tabProperties", itemCnt="1")
    _sub(tabs, "hh", "tabPr", id="0", autoTabLeft="0", autoTabRight="0")
    styles = _sub(ref, "hh", "styles", itemCnt="1")
    _sub(styles, "hh", "style", id="0", type="PARA", name="바탕글", engName="Normal", paraPrIDRef="0", charPrIDRef="0", nextStyleIDRef="0", langID="1042", lockForm="0")
    _sub(root, "hh", "compatibleDocument", targetProgram="NONE")
    _sub(root, "hh", "docOption")
    _sub(root, "hh", "metaTag")
    _sub(root, "hh", "trackChangeConfig", flags="0")
    return _xml(root)


def _merge_maps(merges: tuple[Merge, ...]) -> tuple[dict[tuple[int, int], Merge], set[tuple[int, int]]]:
    starts = {(merge.row, merge.column): merge for merge in merges}
    covered: set[tuple[int, int]] = set()
    for merge in merges:
        for row in range(merge.row, merge.row + merge.row_span):
            for column in range(merge.column, merge.column + merge.column_span):
                if (row, column) != (merge.row, merge.column):
                    covered.add((row, column))
    return starts, covered


def section_xml(table: TableInput) -> bytes:
    char_map, border_map, align_map = _style_maps(table)
    widths = normalize_widths(table.column_widths)
    starts, covered = _merge_maps(table.merges)
    root = ET.Element(q("hs", "sec"), {"id": "0", "textDirection": "HORIZONTAL"})
    paragraph = _sub(root, "hp", "p", id="1000000001", paraPrIDRef="0", styleIDRef="0", pageBreak="0", columnBreak="0", merged="0")
    run = _sub(paragraph, "hp", "run", charPrIDRef="0")
    sec_pr = _sub(run, "hp", "secPr", id="1000000002", textDirection="HORIZONTAL", spaceColumns="1134", tabStop="8000", tabStopVal="4000", outlineShapeIDRef="0", memoShapeIDRef="0", textVerticalWidthHead="0")
    _sub(sec_pr, "hp", "pagePr", landscape="WIDELY", width="59528", height="84188", gutterType="LEFT_ONLY")
    _sub(sec_pr, "hp", "pageMargin", left="8504", right="8504", top="5668", bottom="4252", header="4252", footer="4252", gutter="0")
    _sub(run, "hp", "ctrl")
    _sub(run, "hp", "colPr", id="1000000003", type="NEWSPAPER", layout="LEFT", colCount="1", sameSz="1", sameGap="0")
    table_node = _sub(run, "hp", "tbl", id="1000000004", zOrder="0", numberingType="TABLE", textWrap="TOP_AND_BOTTOM", textFlow="BOTH_SIDES", lock="0", dropcapstyle="None", pageBreak="CELL", repeatHeader="1" if table.title_row else "0", rowCnt=str(table.rows), colCnt=str(table.columns), cellSpacing="0", borderFillIDRef="1", noAdjust="0")
    _sub(table_node, "hp", "sz", width="42520", widthRelTo="ABSOLUTE", height="0", heightRelTo="ABSOLUTE", protect="0")
    _sub(table_node, "hp", "pos", treatAsChar="1", affectLSpacing="0", flowWithText="1", allowOverlap="0", holdAnchorAndSO="0", vertRelTo="PARA", horzRelTo="PARA", vertAlign="TOP", horzAlign="LEFT", vertOffset="0", horzOffset="0")
    _sub(table_node, "hp", "outMargin", left="0", right="0", top="0", bottom="0")
    _sub(table_node, "hp", "inMargin", left="141", right="141", top="141", bottom="141")

    next_id = 1000000010
    for row in range(1, table.rows + 1):
        tr = _sub(table_node, "hp", "tr")
        for column in range(1, table.columns + 1):
            if (row, column) in covered:
                continue
            cell: Cell = table.cell_at(row, column)
            merge = starts.get((row, column))
            row_span = merge.row_span if merge else 1
            column_span = merge.column_span if merge else 1
            width = sum(widths[column - 1:column - 1 + column_span])
            height = round(sum(table.row_heights[row - 1:row - 1 + row_span]) * 100)
            tc = _sub(tr, "hp", "tc", name="", header="1" if table.title_row and row == 1 else "0", hasMargin="0", protect="0", editable="0", dirty="0", borderFillIDRef=str(border_map[(cell.fill_rgb, cell.borders)]))
            _sub(tc, "hp", "subList", id=str(next_id), textDirection="HORIZONTAL", lineWrap="BREAK", vertAlign=cell.vertical.upper(), linkListIDRef="0", linkListNextIDRef="0", textWidth="0", textHeight="0", hasTextRef="0", hasNumRef="0")
            sublist = list(tc)[0]
            cell_p = _sub(sublist, "hp", "p", id=str(next_id + 1), paraPrIDRef=str(align_map[cell.horizontal]), styleIDRef="0", pageBreak="0", columnBreak="0", merged="0")
            cell_run = _sub(cell_p, "hp", "run", charPrIDRef=str(char_map[(cell.font_bold, cell.font_rgb)]))
            text = _sub(cell_run, "hp", "t")
            text.text = cell.text
            _sub(tc, "hp", "cellAddr", colAddr=str(column - 1), rowAddr=str(row - 1))
            _sub(tc, "hp", "cellSpan", colSpan=str(column_span), rowSpan=str(row_span))
            _sub(tc, "hp", "cellSz", width=str(width), height=str(height))
            _sub(tc, "hp", "cellMargin", left="141", right="141", top="141", bottom="141")
            next_id += 2
    return _xml(root)


def package_parts(table: TableInput) -> dict[str, bytes]:
    preview = "\n".join("\t".join(table.cell_at(row, column).text for column in range(1, table.columns + 1)) for row in range(1, table.rows + 1))
    return {
        "version.xml": version_xml(),
        "settings.xml": settings_xml(),
        "META-INF/manifest.xml": manifest_xml(),
        "Contents/content.hpf": content_hpf_xml(),
        "Contents/header.xml": header_xml(table),
        "Contents/section0.xml": section_xml(table),
        "Preview/PrvText.txt": preview.encode("utf-8"),
    }
