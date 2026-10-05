"""생성 결과를 입력 생성기와 분리해 검사하는 런타임 검증기."""

from __future__ import annotations

from pathlib import Path, PurePosixPath
from collections import Counter
import zipfile
from xml.etree import ElementTree as ET

from .xmlparts import NS, q


REQUIRED = {
    "mimetype", "version.xml", "settings.xml", "META-INF/manifest.xml",
    "Contents/content.hpf", "Contents/header.xml", "Contents/section0.xml",
    "Preview/PrvText.txt",
}
FORBIDDEN_TOKENS = ("vba", "macro", "ole", "external")


def _failure(status: str, errors: list[str], missing: list[str] | None = None) -> dict[str, object]:
    return {"status": status, "errors": errors, "missing_references": missing or []}


def validate_hwpx(path: str | Path) -> dict[str, object]:
    target = Path(path)
    if not target.is_file():
        return _failure("INCOMPLETE", ["파일을 찾을 수 없습니다."])
    errors: list[str] = []
    missing_references: list[str] = []
    try:
        with zipfile.ZipFile(target) as archive:
            infos = archive.infolist()
            names = [info.filename for info in infos]
            if len(infos) > 128:
                errors.append("ZIP 항목 수 제한을 초과했습니다.")
            duplicates = sorted(name for name, count in Counter(names).items() if count > 1)
            if duplicates:
                errors.append(f"ZIP 항목 중복: {duplicates}")
            if not infos or infos[0].filename != "mimetype" or infos[0].compress_type != zipfile.ZIP_STORED:
                errors.append("mimetype은 첫 항목이며 무압축이어야 합니다.")
            elif archive.read("mimetype") != b"application/hwp+zip":
                errors.append("mimetype 내용이 잘못되었습니다.")
            missing = sorted(REQUIRED - set(names))
            if missing:
                return _failure("INCOMPLETE", [f"필수 항목 누락: {missing}"])
            for name in names:
                lower = name.lower()
                if any(token in lower for token in FORBIDDEN_TOKENS) or PurePosixPath(name).is_absolute() or ".." in PurePosixPath(name).parts:
                    errors.append(f"금지되거나 안전하지 않은 ZIP 항목: {name}")

            parsed: dict[str, ET.Element] = {}
            for name in ("version.xml", "settings.xml", "META-INF/manifest.xml", "Contents/content.hpf", "Contents/header.xml", "Contents/section0.xml"):
                try:
                    parsed[name] = ET.fromstring(archive.read(name))
                except ET.ParseError as exc:
                    errors.append(f"XML 구문 오류 {name}: {exc}")
            if errors and len(parsed) < 6:
                return _failure("FAIL", errors)

            content = parsed["Contents/content.hpf"]
            manifest_items = content.findall(f".//{{{NS['opf']}}}item")
            for item in manifest_items:
                href = item.attrib.get("href", "")
                if not href or "://" in href or PurePosixPath(href).is_absolute() or ".." in PurePosixPath(href).parts:
                    errors.append(f"안전하지 않은 content.hpf 참조: {href}")
                    continue
                resolved = href if href.startswith("Contents/") else f"Contents/{href}"
                if resolved not in names:
                    missing_references.append(resolved)
            if missing_references:
                errors.append("content.hpf 참조 대상이 없습니다.")
            package_manifest = parsed["META-INF/manifest.xml"]
            for entry in package_manifest:
                href = entry.attrib.get(f"{{{NS['odf']}}}full-path", "")
                if href == "/":
                    continue
                if not href or "://" in href or PurePosixPath(href).is_absolute() or ".." in PurePosixPath(href).parts:
                    errors.append(f"안전하지 않은 패키지 manifest 참조: {href}")
                elif href not in names:
                    missing_references.append(href)
            if missing_references and "패키지 manifest 참조 대상이 없습니다." not in errors:
                errors.append("패키지 manifest 참조 대상이 없습니다.")

            header = parsed["Contents/header.xml"]
            fontfaces = header.find(f".//{{{NS['hh']}}}fontfaces")
            groups = list(fontfaces) if fontfaces is not None else []
            if len(groups) != 7:
                errors.append("글꼴 슬롯은 7개여야 합니다.")
            if fontfaces is None or fontfaces.attrib.get("itemCnt") != str(len(groups)):
                errors.append("fontfaces.itemCnt가 실제 슬롯 수와 다릅니다.")
            for group in groups:
                fonts = group.findall(q("hh", "font"))
                if group.attrib.get("fontCnt") != str(len(fonts)):
                    errors.append("fontCnt가 실제 글꼴 수와 다릅니다.")
                if len(fonts) != 1 or fonts[0].attrib.get("face") != "중고딕" or fonts[0].attrib.get("type") != "HFT":
                    errors.append("모든 글꼴 슬롯은 중고딕/HFT여야 합니다.")
            for collection_tag, item_tag in (("borderFills", "borderFill"), ("charProperties", "charPr"), ("paraProperties", "paraPr"), ("tabProperties", "tabPr"), ("styles", "style")):
                collection = header.find(f".//{{{NS['hh']}}}{collection_tag}")
                if collection is None or collection.attrib.get("itemCnt") != str(len(collection.findall(q("hh", item_tag)))):
                    errors.append(f"{collection_tag}.itemCnt가 실제 항목 수와 다릅니다.")

            char_ids = {item.attrib.get("id") for item in header.findall(f".//{{{NS['hh']}}}charPr")}
            para_ids = {item.attrib.get("id") for item in header.findall(f".//{{{NS['hh']}}}paraPr")}
            border_ids = {item.attrib.get("id") for item in header.findall(f".//{{{NS['hh']}}}borderFill")}

            section = parsed["Contents/section0.xml"]
            table = section.find(f".//{{{NS['hp']}}}tbl")
            if table is None:
                errors.append("표를 찾을 수 없습니다.")
            else:
                rows = int(table.attrib.get("rowCnt", "0"))
                columns = int(table.attrib.get("colCnt", "0"))
                addresses: set[tuple[int, int]] = set()
                occupied: set[tuple[int, int]] = set()
                for cell in table.findall(f".//{{{NS['hp']}}}tc"):
                    address = cell.find(q("hp", "cellAddr"))
                    span = cell.find(q("hp", "cellSpan"))
                    size = cell.find(q("hp", "cellSz"))
                    if address is None or span is None or size is None:
                        errors.append("셀 주소/병합/크기 정보가 누락되었습니다.")
                        continue
                    column = int(address.attrib.get("colAddr", "-1"))
                    row = int(address.attrib.get("rowAddr", "-1"))
                    column_span = int(span.attrib.get("colSpan", "0"))
                    row_span = int(span.attrib.get("rowSpan", "0"))
                    if (row, column) in addresses or row < 0 or column < 0 or row + row_span > rows or column + column_span > columns or row_span < 1 or column_span < 1:
                        errors.append("셀 주소 또는 병합 범위가 잘못되었습니다.")
                    addresses.add((row, column))
                    area = {(r, c) for r in range(row, row + row_span) for c in range(column, column + column_span)}
                    if occupied & area:
                        errors.append("셀 병합 범위가 겹칩니다.")
                    occupied |= area
                    if int(size.attrib.get("width", "0")) <= 0 or int(size.attrib.get("height", "0")) <= 0:
                        errors.append("셀 크기는 양수여야 합니다.")
                    if cell.attrib.get("borderFillIDRef") not in border_ids:
                        errors.append("셀 borderFillIDRef가 header 정의와 일치하지 않습니다.")
                if occupied != {(r, c) for r in range(rows) for c in range(columns)}:
                    errors.append("표의 셀 영역이 완전하지 않습니다.")
                first_row = [
                    cell for cell in table.findall(f".//{{{NS['hp']}}}tc")
                    if cell.find(q("hp", "cellAddr")) is not None
                    and cell.find(q("hp", "cellAddr")).attrib.get("rowAddr") == "0"
                ]
                if sum(int(cell.find(q("hp", "cellSz")).attrib["width"]) for cell in first_row) != 42520:
                    errors.append("첫 행 셀 너비 합계가 본문 너비와 다릅니다.")
            for paragraph in section.findall(f".//{{{NS['hp']}}}p"):
                if paragraph.attrib.get("paraPrIDRef") not in para_ids:
                    errors.append("paraPrIDRef가 header 정의와 일치하지 않습니다.")
            for run in section.findall(f".//{{{NS['hp']}}}run"):
                if run.attrib.get("charPrIDRef") not in char_ids:
                    errors.append("charPrIDRef가 header 정의와 일치하지 않습니다.")
    except (zipfile.BadZipFile, OSError) as exc:
        return _failure("INCOMPLETE", [f"ZIP을 읽을 수 없습니다: {exc}"])
    except (KeyError, TypeError, ValueError) as exc:
        errors.append(f"구조 값 오류: {exc}")
    except Exception as exc:
        errors.append(f"예상하지 못한 검증 오류: {type(exc).__name__}: {exc}")
    return _failure("FAIL" if errors else "PASS", errors, missing_references)
