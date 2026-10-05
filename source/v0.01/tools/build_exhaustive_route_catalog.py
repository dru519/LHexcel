from __future__ import annotations

import argparse
import json
import uuid
from pathlib import Path
from typing import Any


MODAL_COMMAND_KEYS = {
    "RB_PRINT_SETUP_QUICK",
    "RB_SELECTION_CELL_TEXT_SPECIFIC",
    "RB_EDIT_CELL_RESIZE",
    "RB_FILTER_AUTOFILTER_FILTERING_OPTIONAL",
    "RB_WORKBOOK_INFOMATION",
}

POPUP_COMMAND_KEYS = {
    "NX_NUMBER_EMPHASIS", "NX_VIEW_PRESET_SETTINGS", "RB_EDIT_MEMO_ADD_LHEXCELFORMULA",
    "RB_EDIT_ALIGN_CENTER_OVERCELLS",
    "RB_LHESTYLE_OPENOPTIONS",
    "NX_DATA_DATE_CONVERT",
    "NX_FUNCTION_ROUND",
    "NX_FUNCTION_IFERROR",
}


FORBIDDEN_WINDOW_TITLES = ["오류", "Error", "Microsoft Visual Basic", "Microsoft Excel"]


FEATURE_FILE_OUTPUT_ORACLES: dict[str, dict[str, Any]] = {
    "NX-FILE-SHEET-COPY-SAVE": {"type": "file-output", "format": "xlsx", "expected_files": 1},
    "NX-FILE-RANGE-COPY-SAVE": {"type": "file-output", "format": "xlsx", "expected_files": 1},
    "NX-FILE-RANGE-PNG": {"type": "file-output", "format": "png", "expected_files": 1},
    "NX-FILE-CHART-PNG": {"type": "file-output", "format": "png", "expected_files": 1},
    "NX-FILE-PDF-CURRENT-SHEET": {"type": "file-output", "format": "pdf", "expected_files": 1},
    "NX-FILE-PDF-EACH-SHEET": {"type": "file-output", "format": "pdf", "expected_files": 3},
    "NX-FILE-PDF-SELECTED-COMBINED": {"type": "file-output", "format": "pdf", "expected_files": 1},
    "NX-FILE-PDF-ALL-COMBINED": {"type": "file-output", "format": "pdf", "expected_files": 1},
    "NX-FILE-PDF-SETTINGS": {"type": "file-output", "format": "settings", "expected_files": 1},
}


COMMAND_FILE_OUTPUT_ORACLES: dict[str, dict[str, Any]] = {
    "RB_SHEET_SAVE_TO_FILE": {"type": "file-output", "format": "xlsx", "expected_files": 1},
}


DIRECT_COMMAND_ORACLES: dict[str, dict[str, Any]] = {
    "RB_CLIPBOARD_COPY_BYFORMULA": {
        "type": "property", "property": "clipboard.text", "operator": "equals", "expected": "=B2*2"
    },
    "RB_CLIPBOARD_COPY_BYREFERENCE": {
        "type": "property", "property": "clipboard.text", "operator": "fixture-reference", "expected": "B2"
    },
    "RB_CLIPBOARD_COPY_BYTEXT": {
        "type": "property", "property": "clipboard.text", "operator": "equals",
        "expected": "alpha\tbeta\talpha\tgamma",
    },
    "RB_LHECLIPBOARD_COPYVISIBLE": {
        "type": "property", "property": "clipboard.text", "operator": "fixture-visible-text", "expected": "A1:J6"
    },
    "RB_APP_OPTION_EDITDIRECTION_CHANGE": {
        "type": "property", "property": "application.move_after_return_direction", "operator": "equals", "expected": -4161
    },
    "RB_SELECTION_CELL_EXPAND_COL": {
        "type": "property", "property": "selection.address", "operator": "equals", "expected": "$B:$D"
    },
    "RB_SELECTION_CELL_EXPAND_ROW": {
        "type": "property", "property": "selection.address", "operator": "equals", "expected": "$2:$4"
    },
    "RB_SELECTION_CELL_MOVE_DOWN": {
        "type": "property", "property": "selection.address", "operator": "equals", "expected": "$B$5"
    },
    "RB_SELECTION_CELL_MOVE_LEFT": {
        "type": "property", "property": "selection.address", "operator": "equals", "expected": "$A$4"
    },
    "RB_SELECTION_CELL_MOVE_RIGHT": {
        "type": "property", "property": "selection.address", "operator": "equals", "expected": "$E$4"
    },
    "RB_SELECTION_CELL_MOVE_UP": {
        "type": "property", "property": "selection.address", "operator": "equals", "expected": "$B$1"
    },
    "RB_SELECTION_CELL_USEDRANGE": {
        "type": "property", "property": "selection.address", "operator": "equals", "expected": "$A$1:$Z$20"
    },
    "RB_SHEET_SELECT_END": {
        "type": "property", "property": "active_sheet.name", "operator": "equals", "expected": "Last"
    },
    "RB_SHEET_SELECT_HOME": {
        "type": "property", "property": "active_sheet.name", "operator": "equals", "expected": "Data"
    },
    "RB_WINDOWS_VIEW_FULLSCREEN": {
        "type": "property", "property": "application.fullscreen", "operator": "equals", "expected": True
    },
    "RB_WINDOWS_VIEW_SCROLL_DOWN": {
        "type": "property", "property": "window.scroll_row", "operator": "equals", "expected": 11
    },
    "RB_WINDOWS_VIEW_SCROLL_LEFT": {
        "type": "property", "property": "window.scroll_column", "operator": "equals", "expected": 5
    },
    "RB_WINDOWS_VIEW_SCROLL_RIGHT": {
        "type": "property", "property": "window.scroll_column", "operator": "equals", "expected": 6
    },
    "RB_WINDOWS_VIEW_SCROLL_UP": {
        "type": "property", "property": "window.scroll_row", "operator": "equals", "expected": 10
    },
}


MUTATING_COMMAND_ORACLES: dict[str, dict[str, Any]] = {
    "RB_FILTER_AUTOFILTER_CANCEL": {
        "type": "property", "property": "worksheet.filter_mode", "operator": "equals", "expected": False
    },
    "RB_FILTER_AUTOFILTER_SHOWALL": {
        "type": "property", "property": "worksheet.filter_mode", "operator": "equals", "expected": False
    },
    "RB_LHEPRINT_ADD1MMSPACER": {
        "type": "property", "property": "worksheet.left_margin_points", "operator": "equals", "expected": 74.88
    },
    "RB_PRINT_SETUP_REPEAT": {
        "type": "property", "property": "worksheet.print_title_rows", "operator": "equals", "expected": "$2:$4"
    },
    "RB_EDIT_CELL_MAKEGROUP_HIDDEN_ROWCOLUMN": {
        "type": "property", "property": "selection.first_column_outline_level", "operator": "equals", "expected": 2
    },
    "RB_EDIT_MEMO_ADD_LHEXCELFORMULA": {
        "type": "property", "property": "selection.formula_comments", "operator": "equals", "expected": True
    },
}


for _suffix, _property, _expected in [
    *((f"FONT_SIZE_{size}", "selection.font_size", size) for size in (9, 10, 12, 15)),
    ("FONT_COLOR_BLACK", "selection.font_color", 0),
    ("FONT_COLOR_RED", "selection.font_color", 255),
    ("FONT_COLOR_BLUE", "selection.font_color", 16711680),
    ("FONT_COLOR_GREEN", "selection.font_color", 32768),
    ("FILL_GRAY", "selection.fill_color", 14277081),
    ("FILL_LIGHT_RED", "selection.fill_color", 13551615),
    ("FILL_LIGHT_YELLOW", "selection.fill_color", 10284031),
    ("FILL_LIGHT_GREEN", "selection.fill_color", 13561798),
]:
    MUTATING_COMMAND_ORACLES["NX_STYLE_" + _suffix] = {
        "type": "property", "property": _property, "operator": "equals", "expected": _expected,
    }


def owned_window(
    title: str,
    required_any: list[str],
    *,
    required_all: list[str] | None = None,
    expected_classes: list[str] | None = None,
    forbidden_titles: list[str] | None = None,
) -> dict[str, Any]:
    return {
        "type": "owned-window",
        "expected_titles": [title],
        "expected_classes": expected_classes or ["Thunder"],
        "required_controls_any": required_any,
        "required_controls_all": required_all or [],
        "forbidden_titles": FORBIDDEN_WINDOW_TITLES + (forbidden_titles or []),
    }


FEATURE_POPUP_ORACLES: dict[str, dict[str, Any]] = {
    "NX-AI-SUMMARY": owned_window("내엑셀 - AI 모드", ["복사 후 AI 열기", "취소"], required_all=["요약/분석"]),
    "NX-AI-CLEAN": owned_window("내엑셀 - AI 모드", ["복사 후 AI 열기", "취소"], required_all=["데이터 정리"]),
    "NX-AI-FORMULA": owned_window("내엑셀 - AI 모드", ["복사 후 AI 열기", "취소"], required_all=["수식"]),
    "NX-AI-WRITE": owned_window("내엑셀 - AI 모드", ["복사 후 AI 열기", "취소"], required_all=["텍스트"]),
    "NX-AI-IMAGE": owned_window("내엑셀 - AI 모드", ["복사 후 AI 열기", "취소"], required_all=["이미지"]),
    "NX-DATA-NORMALIZE": owned_window("내엑셀 - 데이터 정규화", ["범위 다시 선택", "실행", "취소"], required_all=["기본 정리"]),
    "NX-DATA-UNIQUE-COUNT": owned_window("내엑셀 - 고유값 개수", ["미리보기", "실행", "취소"], required_all=["고유값 개수 · 값 또는 행/열 조합별 빈도 집계"]),
    "NX-DATA-DUPLICATE-LIST": owned_window("내엑셀 - 중복 목록", ["미리보기", "실행", "취소"], required_all=["중복 목록 · 기준 목록과 대조 목록 비교"]),
    "NX-DATA-AGE": owned_window("내엑셀 - 만나이 계산", ["미리보기", "실행", "취소"], required_all=["생년월일·YYMMDD·주민번호형 입력을 기준일의 만 나이로 계산합니다."]),
    "NX-DATA-KOREAN-MONEY": owned_window("내엑셀 - 한글 금액 변환", ["미리보기", "실행", "취소"], required_all=["정수 금액을 한글로 변환합니다. 음수·소수·인식 실패 값은 error로 표시합니다."]),
    "NX-HANGUL-PICTURE-SEND": owned_window(
        "내엑셀 - 아래한글 그림 전송", ["그림 파일 선택", "전송", "취소"],
        required_all=["그림 셀 크기 (가로 × 세로, cm)", "그림 아래 제목행 추가 (파일명)"],
    ),
    "NX-DATA-PRIVACY-MASK": owned_window("내엑셀 - 개인정보 마스킹", ["미리보기", "실행", "취소"], required_all=["종류별로 값을 가립니다. 인식 실패는 error. 개인정보 유출 점검과 별도 기능입니다."]),
    "NX-DRAW-TITLE-TABLE": owned_window("내엑셀 - 표 그리기", ["범위 선택", "적용", "취소"], required_all=["스타일: 제목"]),
    "NX-DRAW-BUSINESS-TABLE": owned_window("내엑셀 - 표 그리기", ["범위 선택", "적용", "취소"], required_all=["스타일: 닫힌 표"]),
    "NX-DRAW-ROLE-STYLE": owned_window("셀스타일", ["범위 다시 선택", "적용", "취소"], required_all=["제목", "단색"]),
    "NX-DRAW-FIT-PICTURE": owned_window("내엑셀 - 그림 삽입", ["범위 선택", "미리보기", "취소"], required_all=["선택 그림 가져오기", "그림 맞춤"]),
    "NX-DRAW-INSERT-PICTURE": owned_window("내엑셀 - 그림 삽입", ["찾아보기", "미리보기", "취소"], required_all=["그림 삽입 / 선택 그림 맞춤", "그림 삽입"]),
    "NX-FILE-CONSOLIDATE": owned_window("내엑셀 - 파일 통합", ["파일 추가", "통합", "취소"], required_all=["결과 위치"]),
    "NX-FILE-RANGE-PNG": owned_window("내엑셀 - 저장", ["미리보기", "실행", "취소"], required_all=["선택 범위를 PNG 이미지로 저장합니다."]),
    "NX-FILE-CHART-PNG": owned_window("내엑셀 - 저장", ["미리보기", "실행", "취소"], required_all=["선택한 차트를 PNG 이미지로 저장합니다."]),
    "NX-FILE-WORKBOOK-COMPARE": owned_window("내엑셀 - 범위 비교", ["범위 선택", "비교 실행", "취소"]),
    "NX-FILE-SHEET-COMPARE": owned_window("내엑셀 - 시트 비교", ["파일 열기", "비교 실행", "취소"]),
    "NX-FILE-FILE-COMPARE": owned_window(
        "내엑셀 - 파일 비교", ["찾아보기", "시트 목록", "비교", "취소"],
        required_all=["기준 파일", "비교 파일", "결과 위치"],
    ),
    "NX-FILE-BATCH-RENAME": owned_window("내엑셀 - 파일 이름 변경 작업대", ["파일 추가", "미리보기", "취소"]),
    "NX-FILE-SHEET-BATCH-RENAME": owned_window("내엑셀 - 시트 이름 변경 작업대", ["미리보기", "취소"]),
    "NX-FILE-PDF-CURRENT-SHEET": owned_window("내엑셀 - PDF", ["찾아보기", "실행", "취소"], required_all=["현재 시트를 PDF 한 파일로 저장합니다."]),
    "NX-FILE-PDF-EACH-SHEET": owned_window("내엑셀 - PDF", ["찾아보기", "실행", "취소"], required_all=["표시된 각 시트를 개별 PDF로 저장합니다."]),
    "NX-FILE-PDF-SELECTED-COMBINED": owned_window("내엑셀 - PDF", ["찾아보기", "실행", "취소"], required_all=["선택한 시트를 PDF 한 파일로 합칩니다."]),
    "NX-FILE-PDF-ALL-COMBINED": owned_window("내엑셀 - PDF", ["찾아보기", "실행", "취소"], required_all=["통합문서 전체를 PDF 한 파일로 저장합니다."]),
    "NX-FILE-PDF-SETTINGS": owned_window("내엑셀 - PDF", ["찾아보기", "실행", "취소"], required_all=["PDF 기본 저장 폴더·이름·품질을 저장합니다."]),
    "NX-FILE-FOLDER-CREATE": owned_window("내엑셀 - 폴더 일괄 생성", ["찾아보기", "생성", "취소"], required_all=["생성할 하위 폴더(한 줄에 하나)"]),
    "NX-TPL-REGISTER-SHEET": owned_window("내엑셀 - 현재시트 등록", ["미리보기", "등록", "취소"], required_all=["등록할 원본을 확인하고 템플릿 이름을 입력하세요."]),
    "NX-TPL-LIST": owned_window("내엑셀 템플릿 관리", ["닫기"], required_all=["보관된 템플릿", "현재시트 등록", "특정파일 등록"]),
    "NX-TPL-LOAD": owned_window("내엑셀 - 템플릿 사용", ["불러오기", "취소"], required_all=["새 시트로 불러올 템플릿을 선택하세요."]),
    "NX-TPL-RENAME": owned_window("내엑셀 - 템플릿 이름·설명 변경", ["이름 변경", "취소"], required_all=["이름을 바꿀 템플릿을 선택하세요."]),
    "NX-TPL-DELETE": owned_window("내엑셀 - 템플릿 제거", ["격리 보관", "취소"], required_all=["격리 보관할 템플릿을 선택하세요."]),
    "NX-UTIL-CALCULATOR": owned_window("내엑셀 - 계산기", ["=", "계산결과 복사", "계산결과 붙여넣기"], required_all=["선택 셀값 불러오기"]),
    "NX-UTIL-SYMBOLS": owned_window("내엑셀 - 기호표", ["복사", "셀에 입력", "닫기"], required_all=["실무 기호", "유니코드", "최근 사용"]),
    "NX-UTIL-NAVIGATOR": owned_window(
        "내엑셀 Navigator", ["검색", "TextBox", "ListBox"], required_all=["전체"],
        expected_classes=["WindowsForms10", "CustomTaskPane", "NetUIHWND"], forbidden_titles=["내엑셀 - 탐색창"],
    ),
    "NX-UTIL-DOCUMENT-NAVIGATOR": owned_window(
        "내엑셀 - 문서·시트 탐색", ["새로고침", "이동", "닫기"], required_all=["열린 문서", "시트"],
    ),
}


COMMAND_POPUP_ORACLES: dict[str, dict[str, Any]] = {
    "NX_NUMBER_EMPHASIS": owned_window("내엑셀 - 숫자 강조(▲▼)", ["적용", "취소"], required_all=["백분율로 표시"]),
    "NX_VIEW_PRESET_SETTINGS": owned_window("내엑셀 - 화면 프리셋 설정", ["저장", "취소"], required_all=["화면 프리셋 1", "화면 프리셋 5"]),
    "RB_EDIT_MEMO_ADD_LHEXCELFORMULA": owned_window("내엑셀 - 셀수식 메모 기록", ["예", "아니요", "취소"], expected_classes=["#32770"]),
    "NX_DATA_DATE_CONVERT": owned_window("내엑셀 - 날짜/생년월일 변환", ["실행", "취소"], required_all=["결과 위치", "날짜 표시 형식"]),
    "NX_FUNCTION_ROUND": owned_window("내엑셀 - 함수 감싸기", ["적용", "취소"], required_all=["자릿수"]),
    "NX_FUNCTION_IFERROR": owned_window("내엑셀 - 함수 감싸기", ["적용", "취소"]),
    "RB_EDIT_ALIGN_CENTER_OVERCELLS": owned_window("내엑셀 - 모형병합 실행/취소", ["범위 선택", "적용", "원상복구"], required_all=["가운데"]),
    "RB_LHESTYLE_OPENOPTIONS": owned_window("셀스타일", ["범위 다시 선택", "적용", "취소"], required_all=["제목", "단색"]),
    "RB_SELECTION_CELL_TEXT_SPECIFIC": owned_window(
        "LHexcel", ["Text to find", "Edit", "OK", "Cancel"], expected_classes=["#32770", "bosa_sdm"]
    ),
    "RB_WORKBOOK_INFOMATION": owned_window("내엑셀 버전", ["내엑셀", "닫기"]),
}


for _index, _zoom in enumerate((75, 100, 125, 150, 200), 1):
    DIRECT_COMMAND_ORACLES[f"NX_VIEW_PRESET_{_index}"] = {"type": "property", "property": "window.zoom", "operator": "equals", "expected": _zoom}


def load(path: Path) -> dict[str, Any]:
    return json.loads(path.read_text(encoding="utf-8"))


def fixture_for_owner(owner: str) -> str:
    return {
        "ai": "table",
        "template": "table",
        "data": "table",
        "draw": "drawing",
        "file": "file",
        "hangul": "table",
        "calculator": "formula",
        "symbols": "blank",
        "navigator": "blank",
    }.get(owner, "blank")


def fixture_for_handler(handler: str) -> str:
    return {
        "NxCmdClipboard": "table",
        "NxCmdPrint": "table",
        "NxCmdStyle": "table",
        "NxCmdNumberFormat": "table",
        "NxCmdSelection": "table",
        "NxCmdRowsColumns": "table",
        "NxCmdAlignment": "table",
        "NxCmdInsertDelete": "sheets",
        "NxCmdFilterSort": "table",
        "NxCmdFormula": "formula",
        "NxCmdSheet": "sheets",
        "NxCmdView": "sheets",
        "NxCmdInfo": "sheets",
        "NxCmdPivot": "table",
    }[handler]


def feature_case(row: dict[str, Any]) -> dict[str, Any]:
    route_id = row["id"]
    launch = row["launch_surface"]
    owner = row["owner"]
    if route_id == "NX-HANGUL-TABLE-SEND":
        mode, expected, timeout = "external", "external-open", 30000
        oracle = {"type": "external-open"}
    elif route_id in FEATURE_FILE_OUTPUT_ORACLES:
        mode, expected, timeout = "destructive", "file-output", 30000
        oracle = FEATURE_FILE_OUTPUT_ORACLES[route_id]
    elif launch in {"dialog", "workbench"}:
        mode, expected, timeout = "popup", "window-open", 15000
        oracle = FEATURE_POPUP_ORACLES[route_id]
    elif route_id in {"NX-UTIL-CALCULATOR", "NX-UTIL-SYMBOLS", "NX-UTIL-NAVIGATOR", "NX-UTIL-DOCUMENT-NAVIGATOR"}:
        mode, expected, timeout = "popup", "window-open", 15000
        oracle = FEATURE_POPUP_ORACLES[route_id]
    else:
        mode, expected, timeout = "range", "state-delta", 15000
        oracle = {"type": "state-delta"}
    fixture = "drawing" if route_id == "NX-AI-IMAGE" else fixture_for_owner(owner)
    return {
        "route_id": route_id,
        "kind": "feature",
        "mode": mode,
        "fixture": fixture,
        "expected": expected,
        "timeout_ms": timeout,
        "oracle": oracle,
    }


def command_case(row: dict[str, Any]) -> dict[str, Any]:
    key = row["args"][0]
    if key in COMMAND_FILE_OUTPUT_ORACLES:
        mode, expected, timeout = "destructive", "file-output", 30000
        oracle = COMMAND_FILE_OUTPUT_ORACLES[key]
    elif key in POPUP_COMMAND_KEYS:
        mode, expected, timeout = "popup", "window-open", 15000
        oracle = COMMAND_POPUP_ORACLES[key]
    elif key in MODAL_COMMAND_KEYS:
        mode = "popup"
        expected = "oracle" if key in MUTATING_COMMAND_ORACLES else (
            "state-delta" if row["mutates_document"] else "window-open"
        )
        timeout = 15000
        oracle = MUTATING_COMMAND_ORACLES.get(key)
        if oracle is None:
            oracle = {"type": "state-delta"} if row["mutates_document"] else COMMAND_POPUP_ORACLES[key]
    elif key in MUTATING_COMMAND_ORACLES:
        mode = "range" if row["navigation"]["selection_mode"] == "required" else "destructive"
        expected, timeout = "oracle", 15000
        oracle = MUTATING_COMMAND_ORACLES[key]
    elif row["mutates_document"]:
        mode = "range" if row["navigation"]["selection_mode"] == "required" else "destructive"
        expected, timeout = "state-delta", 15000
        oracle = {"type": "state-delta"}
    else:
        mode, expected, timeout = "direct", "oracle", 15000
        oracle = DIRECT_COMMAND_ORACLES[key]
    return {
        "route_id": row["id"],
        "kind": "command",
        "mode": mode,
        "fixture": fixture_for_handler(row["handler"]),
        "expected": expected,
        "timeout_ms": timeout,
        "oracle": oracle,
    }


def build(root: Path) -> dict[str, Any]:
    features = load(root / "contracts" / "feature-contract.json")
    commands = load(root / "contracts" / "command-contract.json")
    declared_release_ids = features["lineage"]["release_feature_ids"]
    release_ids = set(declared_release_ids)
    feature_rows = [row for row in features["features"] if row["id"] in release_ids]
    if (len(declared_release_ids) != len(release_ids)
            or {row["id"] for row in feature_rows} != release_ids):
        raise ValueError("exhaustive release feature declarations must match unique feature rows")
    cases = [feature_case(row) for row in feature_rows]
    cases.extend(command_case(row) for row in commands["commands"])
    cases.sort(key=lambda row: (row["kind"], row["route_id"]))
    if (len(feature_rows) != len(declared_release_ids)
            or len(commands["commands"]) != commands["normalized_command_count"]
            or len(cases) != len(declared_release_ids) + commands["normalized_command_count"]):
        raise ValueError("exhaustive route inventory must match the current feature and command contracts")
    if len({row["route_id"] for row in cases}) != len(cases):
        raise ValueError("exhaustive route IDs must be unique")
    return {
        "schema_version": 1,
        "release_id": "r101",
        "feature_count": len(feature_rows),
        "command_count": len(commands["commands"]),
        "route_count": len(cases),
        "cases": cases,
    }


def write_atomic(path: Path, document: dict[str, Any]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    temporary = path.with_name(f".{path.name}.{uuid.uuid4().hex}.tmp")
    temporary.write_text(
        json.dumps(document, ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
        newline="\n",
    )
    temporary.replace(path)


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--root", default=".")
    parser.add_argument("--output", default="contracts/exhaustive-native-contract.json")
    parser.add_argument("--check", action="store_true")
    args = parser.parse_args()
    root = Path(args.root).resolve()
    output = (root / args.output).resolve()
    output.relative_to(root)
    generated = build(root)
    if args.check:
        if not output.is_file() or load(output) != generated:
            raise SystemExit("exhaustive native contract is stale")
        print(
            "exhaustive native contract PASS: "
            f"{generated['feature_count']} features and {generated['command_count']} commands"
        )
        return 0
    write_atomic(output, generated)
    print(f"wrote {output}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
