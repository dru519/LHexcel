#!/usr/bin/env python3
"""Import and normalize the user-owned v3.4 shortcut catalog for r53."""

from __future__ import annotations

import argparse
import hashlib
import json
import os
import re
import tempfile
from collections import Counter, defaultdict
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
DONOR = ROOT.parent / "v3.4" / "candidate" / "src" / "mdShortcutCatalog.bas"
SOURCE = ROOT / "src" / "resources" / "commands.ko-KR.json"
CONTRACT = ROOT / "contracts" / "command-contract.json"
PROVENANCE = ROOT / "provenance" / "v34-command-catalog.json"

ROW_PATTERN = re.compile(r'^\s*rows\((\d+)\)\s*=\s*"(.*)"\s*$', re.MULTILINE)
COMMAND_ID = re.compile(r"^NX-CMD-[A-Z0-9-]+$")
ARGUMENT_KEY = re.compile(r"^[A-Z0-9_]+$")

GROUPS = {
    "clipboard": {
        "order": 1, "label_ko": "복붙", "category_id": "NX-CAT-DATA",
        "handler": "NxCmdClipboard", "id_prefix": "CLIPBOARD",
        "indices": set(range(28, 38)), "input_context": "selection_or_clipboard",
        "synonyms": ["복사", "붙여넣기", "복붙"],
    },
    "print": {
        "order": 2, "label_ko": "인쇄", "category_id": "NX-CAT-FILE",
        "handler": "NxCmdPrint", "id_prefix": "PRINT",
        "indices": set(range(176, 183)), "input_context": "active_worksheet",
        "synonyms": ["인쇄", "출력", "페이지 설정"],
    },
    "style": {
        "order": 3, "label_ko": "스타일", "category_id": "NX-CAT-DRAW",
        "handler": "NxCmdStyle", "id_prefix": "STYLE",
        "indices": set(range(44, 57)) | {197, 200}, "input_context": "range_selection",
        "synonyms": ["스타일", "서식", "표 서식"],
    },
    "number_format": {
        "order": 4, "label_ko": "표시 형식", "category_id": "NX-CAT-DRAW",
        "handler": "NxCmdNumberFormat", "id_prefix": "NUMBER",
        "indices": set(range(63, 93)), "input_context": "range_selection",
        "synonyms": ["표시 형식", "숫자 서식", "날짜 서식"],
    },
    "font_color": {
        "order": 5, "label_ko": "글꼴·색상", "category_id": "NX-CAT-DRAW",
        "handler": "NxCmdFontColor", "id_prefix": "FONT",
        "indices": set(range(57, 63)), "input_context": "range_selection",
        "synonyms": ["글꼴", "폰트", "글자색", "색상"],
    },
    "selection_move": {
        "order": 6, "label_ko": "선택·이동", "category_id": "NX-CAT-DATA",
        "handler": "NxCmdSelection", "id_prefix": "SELECTION",
        "indices": set(range(93, 102)), "input_context": "active_worksheet",
        "synonyms": ["선택", "이동", "범위"],
    },
    "row_column": {
        "order": 7, "label_ko": "행·열 크기", "category_id": "NX-CAT-DATA",
        "handler": "NxCmdRowsColumns", "id_prefix": "SIZE",
        "indices": set(range(102, 113)), "input_context": "range_selection",
        "synonyms": ["행", "열", "너비", "높이", "크기"],
    },
    "alignment_merge": {
        "order": 8, "label_ko": "정렬·병합", "category_id": "NX-CAT-DRAW",
        "handler": "NxCmdAlignment", "id_prefix": "ALIGN",
        "indices": set(range(113, 121)), "input_context": "range_selection",
        "synonyms": ["정렬", "맞춤", "병합", "들여쓰기"],
    },
    "insert_delete_hide": {
        "order": 9, "label_ko": "삽입·삭제·숨기기", "category_id": "NX-CAT-DATA",
        "handler": "NxCmdInsertDelete", "id_prefix": "STRUCTURE",
        "indices": set(range(121, 128)), "input_context": "active_worksheet",
        "synonyms": ["삽입", "삭제", "숨기기", "그룹"],
    },
    "filter_sort": {
        "order": 10, "label_ko": "필터·정렬", "category_id": "NX-CAT-DATA",
        "handler": "NxCmdFilterSort", "id_prefix": "FILTER",
        "indices": set(range(128, 140)), "input_context": "active_data_region",
        "synonyms": ["필터", "정렬", "오름차순", "내림차순"],
    },
    "formula_reference": {
        "order": 11, "label_ko": "수식·참조", "category_id": "NX-CAT-DATA",
        "handler": "NxCmdFormula", "id_prefix": "FORMULA",
        "indices": set(range(140, 149)), "input_context": "range_selection",
        "synonyms": ["수식", "함수", "참조", "계산"],
    },
    "sheet": {
        "order": 12, "label_ko": "시트", "category_id": "NX-CAT-FILE",
        "handler": "NxCmdSheet", "id_prefix": "SHEET",
        "indices": set(range(149, 159)), "input_context": "active_workbook",
        "synonyms": ["시트", "워크시트", "탭"],
    },
    "view_window": {
        "order": 13, "label_ko": "보기·창", "category_id": "NX-CAT-FILE",
        "handler": "NxCmdView", "id_prefix": "VIEW",
        "indices": set(range(159, 176)), "input_context": "active_window",
        "synonyms": ["보기", "화면", "창", "배율", "스크롤"],
    },
    "info_diagnostics": {
        "order": 14, "label_ko": "정보·진단", "category_id": "NX-CAT-FILE",
        "handler": "NxCmdInfo", "id_prefix": "INFO",
        "indices": set(range(185, 191)), "input_context": "active_workbook",
        "synonyms": ["정보", "진단", "목록", "점검"],
    },
    "pivot": {
        "order": 15, "label_ko": "피벗", "category_id": "NX-CAT-DATA",
        "handler": "NxCmdPivot", "id_prefix": "PIVOT",
        "indices": set(range(183, 185)), "input_context": "active_workbook",
        "synonyms": ["피벗", "피벗테이블", "요약"],
    },
}

RIBBON_GROUP_BY_COMMAND_GROUP = {
    "clipboard": "NX-GRP-COPY",
    "print": "NX-GRP-PRINT",
    "style": "NX-GRP-STYLE",
    "number_format": "NX-GRP-STYLE",
    "font_color": "NX-GRP-STYLE",
    "selection_move": "NX-GRP-DATA",
    "row_column": "NX-GRP-CELL-FIT",
    "alignment_merge": "NX-GRP-STYLE",
    "insert_delete_hide": "NX-GRP-INSERT",
    "filter_sort": "NX-GRP-DATA",
    "formula_reference": "NX-GRP-FORMULA",
    "sheet": "NX-GRP-SHEET",
    "view_window": "NX-GRP-VIEW",
    "info_diagnostics": "NX-GRP-INFO",
    "pivot": "NX-GRP-DATA",
}
FLAT_NAVIGATION_GROUPS = {
    "clipboard",
    "print",
    "row_column",
    "formula_reference",
    "sheet",
    "view_window",
    "info_diagnostics",
}
OBJECT_SCOPE_BY_GROUP = {
    "clipboard": ["selection", "clipboard"],
    "print": ["worksheet"],
    "style": ["selection"],
    "number_format": ["selection"],
    "font_color": ["selection"],
    "selection_move": ["worksheet", "selection"],
    "row_column": ["selection"],
    "alignment_merge": ["selection"],
    "insert_delete_hide": ["worksheet", "selection"],
    "filter_sort": ["data_region"],
    "formula_reference": ["selection"],
    "sheet": ["workbook", "worksheet"],
    "view_window": ["window"],
    "info_diagnostics": ["workbook"],
    "pivot": ["workbook", "data_region"],
}
SELECTION_MODE_BY_CONTEXT = {
    "range_selection": "required",
    "active_worksheet": "optional",
    "active_workbook": "none",
    "active_window": "none",
    "selection_or_clipboard": "optional",
    "active_data_region": "required",
}
if set(RIBBON_GROUP_BY_COMMAND_GROUP) != set(GROUPS) or set(OBJECT_SCOPE_BY_GROUP) != set(GROUPS):
    raise RuntimeError("command group navigation map is incomplete")

ALLOWED_HANDLERS = [str(GROUPS[key]["handler"]) for key in GROUPS]
REQUIRED_FIELDS = [
    "id", "label_ko", "aliases_ko", "description_ko", "image_mso", "search_terms_ko",
    "category_id", "group_id", "group_order", "sort_order", "input_context",
    "active_when", "mutates_document", "recovery_policy", "ribbon_visible",
    "taskpane_visible", "favorite_eligible", "shortcut_eligible", "handler",
    "args", "legacy_ids", "legacy_entrypoints", "navigation",
]

# Office built-in images only.  The command catalog is part of v0.01 at
# runtime; these rules do not load v3.4 or any third-party image asset.
# Rules are ordered from the most specific operation to the group fallback so
# adjacent menu items remain visually distinguishable without inventing a
# custom icon pack.
COMMAND_ICON_DEFAULTS = {
    "clipboard": "Copy",
    "print": "DefinePrintStyles",
    "style": "TableAutoFormat",
    "number_format": "NumberingGallery",
    "font_color": "FontColorPicker",
    "selection_move": "SelectAll",
    "row_column": "RowHeight",
    "alignment_merge": "AlignCenter",
    "insert_delete_hide": "CellsInsertDialog",
    "filter_sort": "Filter",
    "formula_reference": "TableFormulaDialog",
    "sheet": "WorksheetInsert",
    "view_window": "ViewNormalViewExcel",
    "info_diagnostics": "FileProperties",
    "pivot": "PivotTableInsert",
}
COMMAND_ICON_RULES = {
    "clipboard": [
        (("PASTEVISIBLEFORMULAS", "PASTE_BYFORMULA"), "PasteFormulas"),
        (("PASTEVISIBLEFORMATS", "PASTE_BYSTYLE"), "PasteFormatting"),
        (("PASTEVISIBLEVALUES", "PASTE_BYVALE"), "PasteValues"),
        (("COPY_BYREFERENCE",), "NameDefine"),
    ],
    "print": [
        (("SETUP_REPEAT",), "TableHeaderRow"),
        (("SETUP_QUICK",), "PrintPreviewAndPrint"),
    ],
    "style": [
        (("RESETWORKBOOKSTYLES",), "DataFormDeleteRecord"),
        (("OPENOPTIONS",), "CellStylesGallery"),
        (("TABLEHEADER",), "TableHeaderRow"),
        (("EMPHASISCELL",), "CellFillColorPicker"),
        (("TOTALROW",), "AutoSum"),
        (("SUBTITLE",), "BorderBottom"),
        (("TITLE",), "FontColorPicker"),
        (("MEMO",), "ReviewShowAllComments"),
    ],
    "number_format": [
        (("DATE_", "CTRLSHIFT2", "CTRLSHIFT3"), "DateAndTimeInsert"),
        (("KRW", "USD", "EURO", "JPY"), "AccountingFormat"),
        (("PERCENTAGE", "UPDOWN"), "NumberFormatDialog"),
    ],
    "font_color": [
        (("SIZE_DOWN",), "ShapeDownArrow"),
        (("SIZE_UP",), "ShapeUpArrow"),
        (("TEXTFITTING",), "TextDirectionLeftToRight"),
    ],
    "selection_move": [
        (("EDITDIRECTION",), "TextDirectionLeftToRight"),
        (("TEXT_SPECIFIC",), "ViewAll"),
        (("MOVE_LEFT",), "ShapeLeftArrow"),
        (("MOVE_RIGHT",), "ShapeRightArrow"),
        (("MOVE_DOWN",), "ShapeDownArrow"),
        (("MOVE_UP",), "ShapeUpArrow"),
    ],
    "row_column": [
        (("AUTOFIT_COLUMN",), "TableAutoFormat"),
        (("CELL_RESIZE",), "PropertySheet"),
        (("SIZE_COLDOWN",), "ShapeLeftArrow"),
        (("SIZE_COLUP",), "ShapeRightArrow"),
        (("SIZE_ROWDOWN",), "ShapeDownArrow"),
        (("SIZE_ROWUP",), "ShapeUpArrow"),
    ],
    "alignment_merge": [
        (("WINDOWS_ALIGN_TILE",), "WindowsArrangeAll"),
        (("CELL_MERGE",), "MergeCenter"),
        (("ALIGN_LEFT", "INDENT_SUBSTRACT"), "ShapeLeftArrow"),
        (("ALIGN_RIGHT", "INDENT_ADD"), "ShapeRightArrow"),
    ],
    "insert_delete_hide": [
        (("MAKEGROUP",), "PropertySheet"),
        (("DELETE",), "DataFormDeleteRecord"),
        (("SHOW_COLUNHIDE",), "ViewAll"),
    ],
    "filter_sort": [
        (("SORT_ASCENDING",), "ShapeUpArrow"),
        (("SORT_DESCENDING",), "ShapeDownArrow"),
        (("AUTOFILTER_CANCEL",), "RefreshAll"),
        (("AUTOFILTER_SHOWALL",), "ViewAll"),
        (("DATA_SHOW_SOURCE",), "PivotTableInsert"),
        (("CONVERTTONUMBER",), "NumberingGallery"),
    ],
    "formula_reference": [
        (("PRECEDENTS_LIST",), "NameDefine"),
        (("FORMULA_SUM",), "AutoSum"),
        (("MINUS1", "MULTIPLY1"), "CalculateNow"),
        (("NOTE_",), "ReviewShowAllComments"),
    ],
    "sheet": [
        (("COPY_",), "Copy"),
        (("MOVE_END", "MOVE_NEXT"), "ShapeRightArrow"),
        (("MOVE_FIRST", "MOVE_PREVIOUS"), "ShapeLeftArrow"),
        (("SAVE_TO_FILE",), "FileSaveAs"),
        (("SELECT_",), "SelectAll"),
    ],
    "view_window": [
        (("FULLSCREEN",), "ViewAll"),
        (("SHOW_GRID",), "ViewGridlinesWord"),
        (("SHOW_ZERO",), "NumberingGallery"),
        (("COMPAREVIEW",), "WindowsArrangeAll"),
        (("SHOW_PAGEBREAKS",), "DefinePrintStyles"),
        (("FREEZE_PAN",), "PropertySheet"),
        (("ZOOM_IN",), "ZoomIn"),
        (("ZOOM_OUT",), "ZoomOut"),
        (("SCROLL_DOWN",), "ShapeDownArrow"),
        (("SCROLL_UP",), "ShapeUpArrow"),
        (("SCROLL_LEFT",), "ShapeLeftArrow"),
        (("SCROLL_RIGHT",), "ShapeRightArrow"),
        (("ZOOM_SELECTION", "ZOOM_USEDRANGE"), "SelectAll"),
    ],
    "info_diagnostics": [
        (("NAMEUNHIDE",), "NameDefine"),
        (("MEMO_LIST",), "ReviewShowAllComments"),
        (("SHEET_LIST", "SHOW_WORKSHEET_LIST"), "WorksheetInsert"),
        (("SHOW_WORKBOOK_LIST",), "FileOpen"),
    ],
}
if set(COMMAND_ICON_DEFAULTS) != set(GROUPS) or set(COMMAND_ICON_RULES) != set(GROUPS) - {"pivot"}:
    raise RuntimeError("command icon rules are incomplete")

STRIPPED_SUFFIXES = (
    "_CANTUNDO", "_KEYEVENT", "_BYKEY", "_ALL", "_BTN", "_B",
)


def sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def _selected_group(index: int) -> str | None:
    matches = [key for key, value in GROUPS.items() if index in value["indices"]]
    if len(matches) > 1:
        raise ValueError(f"donor row {index} belongs to multiple command groups")
    return matches[0] if matches else None


def _decode_vba_string(value: str) -> str:
    return value.replace('""', '"')


def parse_donor(source: str) -> list[dict[str, object]]:
    rows: list[dict[str, object]] = []
    seen_indexes: set[int] = set()
    for match in ROW_PATTERN.finditer(source):
        index = int(match.group(1))
        group_id = _selected_group(index)
        if group_id is None:
            continue
        if index in seen_indexes:
            raise ValueError(f"duplicate donor row index: {index}")
        fields = _decode_vba_string(match.group(2)).split("|")
        if len(fields) != 6:
            raise ValueError(f"donor row {index} must have exactly six fields")
        legacy_id, label, shortcut, category, description, entrypoints = fields
        if not legacy_id.strip() or not label.strip() or not category.strip() or not entrypoints.strip():
            raise ValueError(f"donor row {index} has an empty required value")
        rows.append(
            {
                "index": index,
                "group_id": group_id,
                "legacy_id": legacy_id.strip(),
                "label_ko": label.strip(),
                "shortcut": shortcut.strip(),
                "source_category_ko": category.strip(),
                "description_ko": description.strip(),
                "legacy_entrypoints": [item.strip() for item in entrypoints.split(",") if item.strip()],
            }
        )
        seen_indexes.add(index)
    expected_indexes = set().union(*(value["indices"] for value in GROUPS.values()))
    if seen_indexes != expected_indexes:
        missing = sorted(expected_indexes - seen_indexes)
        extra = sorted(seen_indexes - expected_indexes)
        raise ValueError(f"selected donor rows differ: missing={missing}, extra={extra}")
    counts = Counter(str(row["group_id"]) for row in rows)
    expected_counts = {key: len(value["indices"]) for key, value in GROUPS.items()}
    if dict(counts) != expected_counts:
        raise ValueError("selected donor group counts differ")
    return sorted(rows, key=lambda row: int(row["index"]))


def normalize_operation(entrypoint: str) -> str:
    value = entrypoint.strip().upper()
    changed = True
    while changed:
        changed = False
        for suffix in STRIPPED_SUFFIXES:
            if value.endswith(suffix):
                value = value[: -len(suffix)]
                changed = True
                break
    value = re.sub(r"[^A-Z0-9]+", "_", value).strip("_")
    if ARGUMENT_KEY.fullmatch(value) is None:
        raise ValueError(f"unsafe normalized operation: {entrypoint}")
    return value


def _command_id(group: dict[str, object], operation: str) -> str:
    suffix = operation.replace("_", "-")
    value = f"NX-CMD-{group['id_prefix']}-{suffix}"
    if COMMAND_ID.fullmatch(value) is None:
        raise ValueError(f"unsafe command id: {value}")
    return value


def _mutates_document(group_id: str, operation: str) -> bool:
    if group_id in {"selection_move", "view_window"}:
        return False
    if group_id == "clipboard":
        return "PASTE" in operation
    if group_id == "sheet" and ("SELECT_END" in operation or "SELECT_HOME" in operation):
        return False
    if group_id == "info_diagnostics" and "WORKBOOK_INFOMATION" in operation:
        return False
    return True


def _shortcut_eligible(operation: str) -> bool:
    blocked_tokens = (
        "MANAGE_DELETE", "SAVE_TO_FILE", "ADD1MMSPACER", "TEXT_SPECIFIC",
        "FILTERING_OPTIONAL", "SHOW_SOURCE", "PIVOT_BUILDUP",
    )
    return not any(token in operation for token in blocked_tokens)


def _search_terms(group: dict[str, object], labels: list[str]) -> list[str]:
    values = list(group["synonyms"])
    for label in labels:
        simplified = re.sub(r"[\[\]()_/·>.,:+-]+", " ", label)
        values.extend(part for part in simplified.split() if len(part) > 1)
    return sorted({str(value).strip() for value in values if str(value).strip()})


def _command_image_mso(group_id: str, operation: str) -> str:
    for needles, image_mso in COMMAND_ICON_RULES.get(group_id, []):
        if any(needle in operation for needle in needles):
            return image_mso
    return COMMAND_ICON_DEFAULTS[group_id]


def _apply_approved_command_policy(command: dict[str, object], operation: str) -> None:
    """Keep v0.01 product semantics authoritative over donor display text."""
    if operation == "RB_WORKBOOK_INFOMATION":
        command.update(
            {
                "label_ko": "내엑셀 버전 정보",
                "aliases_ko": ["현재 파일 정보 보기", "내엑셀 정보"],
                "description_ko": "현재 설치된 내엑셀의 버전, 완성본 프로필, 기능·명령 수와 빌드 ID를 확인합니다.",
                "image_mso": "FileProperties",
                "search_terms_ko": ["내엑셀", "버전", "빌드", "정보", "프로필"],
            }
        )
    elif operation == "RB_WINDOWS_VIEW_ZOOM_IN":
        command.update(
            {
                "label_ko": "보기 배율 크게 +5%",
                "description_ko": "현재 Excel 보기 배율을 5% 높입니다.",
                "image_mso": "ZoomIn",
                "search_terms_ko": ["5%", "배율", "보기", "스크롤", "증가", "창", "크게", "확대", "화면"],
            }
        )
    elif operation == "RB_WINDOWS_VIEW_ZOOM_OUT":
        command.update(
            {
                "label_ko": "보기 배율 작게 -5%",
                "description_ko": "현재 Excel 보기 배율을 5% 낮춥니다.",
                "image_mso": "ZoomOut",
                "search_terms_ko": ["5%", "감소", "배율", "보기", "스크롤", "작게", "창", "축소", "화면"],
            }
        )
    command["navigation"]["purpose_tags"] = list(command["search_terms_ko"])


def normalize_rows(rows: list[dict[str, object]]) -> list[dict[str, object]]:
    merged: dict[tuple[str, str, tuple[str, ...]], list[dict[str, object]]] = defaultdict(list)
    for row in rows:
        group_id = str(row["group_id"])
        group = GROUPS[group_id]
        operation = normalize_operation(str(row["legacy_entrypoints"][0]))
        key = (group_id, str(group["handler"]), (operation,))
        merged[key].append(row)

    commands: list[dict[str, object]] = []
    seen_ids: set[str] = set()
    for (group_id, handler, args), candidates in merged.items():
        group = GROUPS[group_id]
        labels = sorted({str(row["label_ko"]) for row in candidates}, key=lambda value: (len(value), value))
        label = labels[0]
        command_id = _command_id(group, args[0])
        if command_id in seen_ids:
            raise ValueError(f"duplicate normalized command id: {command_id}")
        descriptions = sorted(
            {str(row["description_ko"]) for row in candidates if str(row["description_ko"]).strip()},
            key=lambda value: (-len(value), value),
        )
        mutates = _mutates_document(group_id, args[0])
        command = {
                "id": command_id,
                "label_ko": label,
                "aliases_ko": sorted(value for value in labels[1:] if value != label),
                "description_ko": descriptions[0] if descriptions else label,
                "image_mso": _command_image_mso(group_id, args[0]),
                "search_terms_ko": _search_terms(group, labels),
                "category_id": group["category_id"],
                "group_id": f"NX-CMD-GRP-{group_id.upper().replace('_', '-')}",
                "group_order": group["order"],
                "sort_order": min(int(row["index"]) for row in candidates),
                "input_context": group["input_context"],
                "active_when": "excel_ready",
                "mutates_document": mutates,
                "recovery_policy": "command_scoped_guard" if mutates else "none",
                "ribbon_visible": False,
                "taskpane_visible": True,
                "favorite_eligible": True,
                "shortcut_eligible": _shortcut_eligible(args[0]),
                "handler": handler,
                "args": list(args),
                "legacy_ids": sorted({str(row["legacy_id"]) for row in candidates}),
                "legacy_entrypoints": sorted(
                    {str(value) for row in candidates for value in row["legacy_entrypoints"]}
                ),
                "navigation": {
                    "ribbon_group_id": RIBBON_GROUP_BY_COMMAND_GROUP[group_id],
                    "object_scope": list(OBJECT_SCOPE_BY_GROUP[group_id]),
                    "purpose_tags": _search_terms(group, labels),
                    "launch_surface": "direct",
                    "mutation_scope": "document" if mutates else "none",
                    "selection_mode": SELECTION_MODE_BY_CONTEXT[str(group["input_context"])],
                    "feedback_mode": "completion" if mutates else "inline",
                    "profile_availability": ["internal-xlam", "enhanced-dll"],
                },
            }
        _apply_approved_command_policy(command, args[0])
        commands.append(command)
        seen_ids.add(command_id)
    commands.sort(key=lambda row: (int(row["group_order"]), int(row["sort_order"]), str(row["id"])))
    return commands


def render_document(commands: list[dict[str, object]]) -> bytes:
    groups = []
    for key, value in GROUPS.items():
        group = {
            "id": f"NX-CMD-GRP-{key.upper().replace('_', '-')}",
            "key": key,
            "label_ko": value["label_ko"],
            "order": value["order"],
            "category_id": value["category_id"],
            "handler": value["handler"],
            "raw_count": len(value["indices"]),
            "ribbon_group_id": RIBBON_GROUP_BY_COMMAND_GROUP[key],
        }
        if key in FLAT_NAVIGATION_GROUPS:
            group["navigation_mode"] = "flat"
        groups.append(group)
    document = {
        "schema_version": 2,
        "required_fields": REQUIRED_FIELDS,
        "allowed_handlers": ALLOWED_HANDLERS,
        "raw_command_count": 159,
        "normalized_command_count": len(commands),
        "groups": groups,
        "commands": commands,
    }
    return (json.dumps(document, ensure_ascii=False, indent=2) + "\n").encode("utf-8")


def render_provenance(raw_rows: list[dict[str, object]], commands: list[dict[str, object]]) -> bytes:
    raw_counts = Counter(str(row["group_id"]) for row in raw_rows)
    normalized_counts = Counter(
        str(row["group_id"]).removeprefix("NX-CMD-GRP-").lower().replace("-", "_")
        for row in commands
    )
    document = {
        "schema_version": 1,
        "source": {
            "path": "../v3.4/candidate/src/mdShortcutCatalog.bas",
            "sha256": sha256(DONOR),
            "reuse_basis": "user-owned-v3.4-lineage",
        },
        "selected_index_ranges": ["28-37", "44-190", "197", "200"],
        "selected_raw_count": len(raw_rows),
        "normalized_command_count": len(commands),
        "raw_group_counts": {key: raw_counts[key] for key in GROUPS},
        "normalized_group_counts": {key: normalized_counts[key] for key in GROUPS},
        "duplicate_groups_merged": len(raw_rows) - len(commands),
        "normalization": {
            "identity": "group + allowlisted handler + normalized legacy operation argument",
            "removed_entrypoint_suffixes": list(STRIPPED_SUFFIXES),
            "canonical_label": "shortest Korean label then ordinal label",
            "aliases": "all remaining Korean labels",
        },
        "third_party_assets": [],
        "kutools_policy": "UI, classification, and interaction patterns only; no Kutools code or binary included",
    }
    return (json.dumps(document, ensure_ascii=False, indent=2) + "\n").encode("utf-8")


def atomic_write(path: Path, content: bytes) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    descriptor, temporary = tempfile.mkstemp(prefix=f".{path.name}.", dir=path.parent)
    try:
        with os.fdopen(descriptor, "wb") as handle:
            handle.write(content)
            handle.flush()
            os.fsync(handle.fileno())
        os.replace(temporary, path)
    finally:
        if os.path.exists(temporary):
            os.unlink(temporary)


def outputs() -> dict[Path, bytes]:
    raw_rows = parse_donor(DONOR.read_text(encoding="utf-8"))
    commands = normalize_rows(raw_rows)
    if len(raw_rows) != 159 or len(commands) != 147:
        raise ValueError("v3.4 command inventory count rejected")
    document = render_document(commands)
    return {
        SOURCE: document,
        CONTRACT: document,
        PROVENANCE: render_provenance(raw_rows, commands),
    }


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    mode = parser.add_mutually_exclusive_group(required=True)
    mode.add_argument("--write", action="store_true")
    mode.add_argument("--check", action="store_true")
    args = parser.parse_args()
    rendered = outputs()
    if args.check:
        stale = [path for path, expected in rendered.items() if not path.is_file() or path.read_bytes() != expected]
        if stale:
            raise SystemExit("stale v3.4 command catalog: " + ", ".join(str(path) for path in stale))
        return 0
    for path, content in rendered.items():
        atomic_write(path, content)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
