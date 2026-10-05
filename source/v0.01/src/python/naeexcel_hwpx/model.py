"""VBA와 Python 사이의 중립 표 데이터 계약."""

from __future__ import annotations

from dataclasses import dataclass
import json
import re
from pathlib import Path
from typing import Any


_RGB = re.compile(r"^[0-9A-Fa-f]{6}$")
_ROOT_KEYS = {
    "schema_version", "rows", "columns", "title_row", "source_mode",
    "cells", "merges", "column_widths", "row_heights",
}
_CELL_KEYS = {
    "row", "column", "text", "fill_rgb", "font_bold", "font_rgb",
    "horizontal", "vertical", "borders",
}
_MERGE_KEYS = {"row", "column", "row_span", "column_span"}


@dataclass(frozen=True)
class Cell:
    row: int
    column: int
    text: str
    fill_rgb: str
    font_bold: bool
    font_rgb: str
    horizontal: str
    vertical: str
    borders: str


@dataclass(frozen=True)
class Merge:
    row: int
    column: int
    row_span: int
    column_span: int


@dataclass(frozen=True)
class TableInput:
    rows: int
    columns: int
    title_row: bool
    source_mode: str
    cells: tuple[Cell, ...]
    merges: tuple[Merge, ...]
    column_widths: tuple[float, ...]
    row_heights: tuple[float, ...]

    def cell_at(self, row: int, column: int) -> Cell:
        return self.cells[(row - 1) * self.columns + column - 1]


def _exact_keys(value: dict[str, Any], expected: set[str], label: str) -> None:
    missing = expected - value.keys()
    unknown = value.keys() - expected
    if missing or unknown:
        raise ValueError(f"{label} 키 오류: missing={sorted(missing)}, unknown={sorted(unknown)}")


def _positive_int(value: Any, label: str) -> int:
    if isinstance(value, bool) or not isinstance(value, int) or value <= 0:
        raise ValueError(f"{label}은 양의 정수여야 합니다.")
    return value


def _rgb(value: Any, label: str) -> str:
    if not isinstance(value, str) or not _RGB.fullmatch(value):
        raise ValueError(f"{label}은 6자리 RGB 값이어야 합니다.")
    return value.upper()


def _number_list(value: Any, count: int, label: str) -> tuple[float, ...]:
    if not isinstance(value, list) or len(value) != count:
        raise ValueError(f"{label}의 개수는 {count}개여야 합니다.")
    result: list[float] = []
    for item in value:
        if isinstance(item, bool) or not isinstance(item, (int, float)) or item <= 0:
            raise ValueError(f"{label}에는 양수만 사용할 수 있습니다.")
        result.append(float(item))
    return tuple(result)


def load_table_input(path: str | Path) -> TableInput:
    payload = json.loads(Path(path).read_text(encoding="utf-8-sig"))
    if not isinstance(payload, dict):
        raise ValueError("입력 JSON의 최상위 값은 객체여야 합니다.")
    _exact_keys(payload, _ROOT_KEYS, "root")
    if payload["schema_version"] != 1:
        raise ValueError("지원하지 않는 schema_version입니다.")

    rows = _positive_int(payload["rows"], "rows")
    columns = _positive_int(payload["columns"], "columns")
    if rows > 5000 or columns > 100 or rows * columns > 50000:
        raise ValueError("표 크기 제한(5000행, 100열, 50000셀)을 초과했습니다.")
    if type(payload["title_row"]) is not bool:
        raise ValueError("title_row는 boolean이어야 합니다.")
    if payload["source_mode"] not in {"display", "value"}:
        raise ValueError("source_mode는 display 또는 value여야 합니다.")

    raw_cells = payload["cells"]
    if not isinstance(raw_cells, list) or len(raw_cells) != rows * columns:
        raise ValueError("cells는 모든 셀을 빠짐없이 한 번씩 포함해야 합니다.")
    cells_by_position: dict[tuple[int, int], Cell] = {}
    for index, raw in enumerate(raw_cells):
        if not isinstance(raw, dict):
            raise ValueError(f"cells[{index}]는 객체여야 합니다.")
        _exact_keys(raw, _CELL_KEYS, f"cells[{index}]")
        row = _positive_int(raw["row"], f"cells[{index}].row")
        column = _positive_int(raw["column"], f"cells[{index}].column")
        if row > rows or column > columns or (row, column) in cells_by_position:
            raise ValueError(f"cells[{index}]의 좌표가 범위를 벗어나거나 중복되었습니다.")
        if not isinstance(raw["text"], str) or len(raw["text"]) > 32767:
            raise ValueError(f"cells[{index}].text가 문자열 제한을 벗어났습니다.")
        if type(raw["font_bold"]) is not bool:
            raise ValueError(f"cells[{index}].font_bold는 boolean이어야 합니다.")
        if raw["horizontal"] not in {"left", "center", "right", "justify"}:
            raise ValueError(f"cells[{index}].horizontal 값이 잘못되었습니다.")
        if raw["vertical"] not in {"top", "center", "bottom"}:
            raise ValueError(f"cells[{index}].vertical 값이 잘못되었습니다.")
        if raw["borders"] not in {"none", "all"}:
            raise ValueError(f"cells[{index}].borders 값이 잘못되었습니다.")
        cells_by_position[(row, column)] = Cell(
            row=row, column=column, text=raw["text"],
            fill_rgb=_rgb(raw["fill_rgb"], f"cells[{index}].fill_rgb"),
            font_bold=raw["font_bold"],
            font_rgb=_rgb(raw["font_rgb"], f"cells[{index}].font_rgb"),
            horizontal=raw["horizontal"], vertical=raw["vertical"],
            borders=raw["borders"],
        )
    expected_positions = {(r, c) for r in range(1, rows + 1) for c in range(1, columns + 1)}
    if cells_by_position.keys() != expected_positions:
        raise ValueError("cells 좌표가 조밀한 직사각형을 이루지 않습니다.")

    merges: list[Merge] = []
    occupied: set[tuple[int, int]] = set()
    raw_merges = payload["merges"]
    if not isinstance(raw_merges, list):
        raise ValueError("merges는 배열이어야 합니다.")
    for index, raw in enumerate(raw_merges):
        if not isinstance(raw, dict):
            raise ValueError(f"merges[{index}]는 객체여야 합니다.")
        _exact_keys(raw, _MERGE_KEYS, f"merges[{index}]")
        row = _positive_int(raw["row"], f"merges[{index}].row")
        column = _positive_int(raw["column"], f"merges[{index}].column")
        row_span = _positive_int(raw["row_span"], f"merges[{index}].row_span")
        column_span = _positive_int(raw["column_span"], f"merges[{index}].column_span")
        if row_span == 1 and column_span == 1:
            raise ValueError("1x1 병합은 허용하지 않습니다.")
        positions = {
            (r, c)
            for r in range(row, row + row_span)
            for c in range(column, column + column_span)
        }
        if max(r for r, _ in positions) > rows or max(c for _, c in positions) > columns:
            raise ValueError("병합 범위가 표를 벗어났습니다.")
        if occupied & positions:
            raise ValueError("병합 범위가 서로 겹칩니다.")
        occupied |= positions
        merges.append(Merge(row, column, row_span, column_span))

    ordered_cells = tuple(cells_by_position[(r, c)] for r in range(1, rows + 1) for c in range(1, columns + 1))
    return TableInput(
        rows=rows, columns=columns, title_row=payload["title_row"],
        source_mode=payload["source_mode"], cells=ordered_cells,
        merges=tuple(merges),
        column_widths=_number_list(payload["column_widths"], columns, "column_widths"),
        row_heights=_number_list(payload["row_heights"], rows, "row_heights"),
    )

