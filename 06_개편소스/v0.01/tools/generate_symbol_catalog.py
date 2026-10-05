#!/usr/bin/env python3
"""Validate the Korean symbol authority and generate ASCII-only VBA rows."""
from __future__ import annotations

import argparse
import json
import os
import re
import tempfile
from collections import Counter
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
SOURCE = ROOT / "src" / "resources" / "symbols.ko-KR.json"
OUTPUT = ROOT / "src" / "vba" / "features" / "symbols" / "NxGeneratedSymbolCatalog.bas"
CATEGORY_IDS = [
    "punctuation",
    "brackets",
    "math",
    "unit_currency",
    "arrows",
    "shapes_checks",
    "numbers",
    "greek_latin",
    "other",
]
EXPECTED_CATEGORY_COUNTS = {
    "punctuation": 16,
    "brackets": 16,
    "math": 16,
    "unit_currency": 34,
    "arrows": 16,
    "shapes_checks": 16,
    "numbers": 44,
    "greek_latin": 16,
    "other": 16,
}
BLOCKS = [
    ("latin1_supplement", 0x0080, 0x00FF),
    ("greek_coptic", 0x0370, 0x03FF),
    ("general_punctuation", 0x2000, 0x206F),
    ("currency_symbols", 0x20A0, 0x20CF),
    ("letterlike_symbols", 0x2100, 0x214F),
    ("number_forms", 0x2150, 0x218F),
    ("arrows", 0x2190, 0x21FF),
    ("math_operators", 0x2200, 0x22FF),
    ("enclosed_alphanumerics", 0x2460, 0x24FF),
    ("box_drawing", 0x2500, 0x257F),
    ("geometric_shapes", 0x25A0, 0x25FF),
    ("misc_symbols", 0x2600, 0x26FF),
    ("dingbats", 0x2700, 0x27BF),
    ("cjk_compatibility", 0x3300, 0x33FF),
]
ID_PATTERN = re.compile(r"^[a-z][a-z0-9_]*$")


def exact_keys(value: dict, expected: set[str], label: str) -> None:
    if not isinstance(value, dict) or set(value) != expected:
        raise ValueError(f"{label} fields must be exact: {sorted(expected)}")


def allowed_scalar(codepoint: int) -> bool:
    if not isinstance(codepoint, int) or isinstance(codepoint, bool) or not 0 <= codepoint <= 0x10FFFF:
        return False
    if codepoint <= 0x1F or 0x7F <= codepoint <= 0x9F or 0xD800 <= codepoint <= 0xDFFF:
        return False
    if 0xFDD0 <= codepoint <= 0xFDEF or codepoint & 0xFFFF in {0xFFFE, 0xFFFF}:
        return False
    return True


def require_label(value: object, label: str) -> str:
    if not isinstance(value, str) or not value.strip():
        raise ValueError(f"{label} must be a non-empty string")
    if "|" in value:
        raise ValueError(f"{label} contains the row delimiter")
    return value


def require_order(value: object, label: str) -> int:
    if not isinstance(value, int) or isinstance(value, bool) or value <= 0:
        raise ValueError(f"{label} must be a positive integer")
    return value


def validate() -> dict:
    catalog = json.loads(SOURCE.read_text(encoding="utf-8"))
    exact_keys(catalog, {"schema_version", "locale", "categories", "blocks", "symbols"}, "catalog")
    if catalog["schema_version"] != 1 or catalog["locale"] != "ko-KR":
        raise ValueError("symbol catalog version or locale rejected")
    if not isinstance(catalog["categories"], list) or len(catalog["categories"]) != 9:
        raise ValueError("symbol catalog requires exact 9 categories")
    if not isinstance(catalog["blocks"], list) or len(catalog["blocks"]) != 14:
        raise ValueError("symbol catalog requires exact 14 blocks")
    if not isinstance(catalog["symbols"], list) or len(catalog["symbols"]) != 190:
        raise ValueError("symbol catalog requires exact 190 symbols")

    category_orders: set[int] = set()
    category_map: dict[str, dict] = {}
    for index, category in enumerate(catalog["categories"]):
        exact_keys(category, {"id", "label", "order"}, "category")
        category_id = category["id"]
        if not isinstance(category_id, str) or not ID_PATTERN.fullmatch(category_id):
            raise ValueError("category id rejected")
        if category_id != CATEGORY_IDS[index] or category_id in category_map:
            raise ValueError("category order or duplicate rejected")
        require_label(category["label"], "category label")
        order = require_order(category["order"], "category order")
        if order in category_orders:
            raise ValueError("duplicate category order")
        category_orders.add(order)
        category_map[category_id] = category

    block_orders: set[int] = set()
    block_map: dict[str, dict] = {}
    for index, block in enumerate(catalog["blocks"]):
        exact_keys(block, {"id", "label", "order", "start", "end"}, "block")
        expected_id, expected_start, expected_end = BLOCKS[index]
        if block["id"] != expected_id or block["start"] != expected_start or block["end"] != expected_end:
            raise ValueError("Unicode block authority changed")
        require_label(block["label"], "block label")
        order = require_order(block["order"], "block order")
        if order in block_orders:
            raise ValueError("duplicate block order")
        block_orders.add(order)
        block_map[block["id"]] = block

    codepoints: set[int] = set()
    orders_by_category: dict[str, set[int]] = {category_id: set() for category_id in CATEGORY_IDS}
    for symbol in catalog["symbols"]:
        exact_keys(
            symbol,
            {"codepoint", "character", "name_ko", "aliases_ko", "category_id", "block_id", "order"},
            "symbol",
        )
        codepoint = symbol["codepoint"]
        if not allowed_scalar(codepoint) or codepoint in codepoints:
            raise ValueError("invalid or duplicate symbol codepoint")
        if not isinstance(symbol["character"], str) or len(symbol["character"]) != 1 or ord(symbol["character"]) != codepoint:
            raise ValueError("symbol character/codepoint mismatch")
        require_label(symbol["name_ko"], "symbol Korean name")
        aliases = symbol["aliases_ko"]
        if not isinstance(aliases, list) or not aliases:
            raise ValueError("symbol aliases must be a non-empty list")
        for alias in aliases:
            require_label(alias, "symbol Korean alias")
        category_id = symbol["category_id"]
        block_id = symbol["block_id"]
        if category_id not in category_map or block_id not in block_map:
            raise ValueError("unknown symbol category or block")
        block = block_map[block_id]
        if not block["start"] <= codepoint <= block["end"]:
            raise ValueError("symbol codepoint is outside its Unicode block")
        order = require_order(symbol["order"], "symbol order")
        if order in orders_by_category[category_id]:
            raise ValueError("duplicate symbol order inside category")
        orders_by_category[category_id].add(order)
        codepoints.add(codepoint)
    counts = Counter(symbol["category_id"] for symbol in catalog["symbols"])
    if counts != Counter(EXPECTED_CATEGORY_COUNTS):
        raise ValueError("practical symbol category counts changed")
    return catalog


def utf16_hex(value: str) -> str:
    raw = value.encode("utf-16-le")
    return ".".join(f"{int.from_bytes(raw[index:index + 2], 'little'):04X}" for index in range(0, len(raw), 2))


def render(catalog: dict) -> str:
    lines = [
        'Attribute VB_Name = "NxGeneratedSymbolCatalog"',
        "Option Explicit",
        "' GENERATED FILE - edit src/resources/symbols.ko-KR.json",
        "Public Function NxGeneratedSymbolCategoryRows() As Collection",
        "    Dim categories As New Collection",
    ]
    for category in catalog["categories"]:
        row = f'{category["id"]}|{category["order"]}|{utf16_hex(category["label"])}'
        lines.append(f'    categories.Add "{row}"')
    lines += ["    Set NxGeneratedSymbolCategoryRows = categories", "End Function", "", "Public Function NxGeneratedSymbolBlockRows() As Collection", "    Dim blocks As New Collection"]
    for block in catalog["blocks"]:
        row = f'{block["id"]}|{block["order"]}|{block["start"]:X}|{block["end"]:X}|{utf16_hex(block["label"])}'
        lines.append(f'    blocks.Add "{row}"')
    lines += ["    Set NxGeneratedSymbolBlockRows = blocks", "End Function", "", "Public Function NxGeneratedSymbolRows() As Collection", "    Dim rows As New Collection"]
    for symbol in catalog["symbols"]:
        aliases = ";".join(symbol["aliases_ko"])
        row = "|".join(
            (
                symbol["category_id"],
                symbol["block_id"],
                str(symbol["order"]),
                f'{symbol["codepoint"]:X}',
                utf16_hex(symbol["name_ko"]),
                utf16_hex(aliases),
            )
        )
        lines.append(f'    rows.Add "{row}"')
    lines += ["    Set NxGeneratedSymbolRows = rows", "End Function", ""]
    result = "\n".join(lines)
    result.encode("ascii")
    return result


def atomic_write(path: Path, content: str) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    handle, raw = tempfile.mkstemp(prefix=path.name + ".", suffix=".tmp", dir=path.parent)
    temp = Path(raw)
    try:
        with os.fdopen(handle, "w", encoding="ascii", newline="\n") as stream:
            stream.write(content)
        os.replace(temp, path)
    finally:
        if temp.exists():
            temp.unlink()


def main() -> int:
    parser = argparse.ArgumentParser()
    mode = parser.add_mutually_exclusive_group(required=True)
    mode.add_argument("--write", action="store_true")
    mode.add_argument("--check", action="store_true")
    args = parser.parse_args()
    expected = render(validate())
    if args.check:
        if not OUTPUT.exists() or OUTPUT.read_bytes() != expected.encode("ascii"):
            raise SystemExit(f"stale generated symbol catalog: {OUTPUT}")
    else:
        atomic_write(OUTPUT, expected)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
