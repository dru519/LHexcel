#!/usr/bin/env python3
"""Generate the deterministic, ASCII-only VBA string module for v0.01."""

from __future__ import annotations

import argparse
import json
import re
import sys
from pathlib import Path
from typing import Any


ROOT = Path(__file__).resolve().parents[1]
DEFAULT_INPUT = ROOT / "src" / "resources" / "strings.ko-KR.json"
DEFAULT_OUTPUT = ROOT / "out" / "generated" / "vba" / "NxStrings.g.bas"
MODULE_NAME = "NxStrings_g"
KEY_SEGMENT_RE = re.compile(r"^[A-Za-z][A-Za-z0-9_]*$")


def reject_duplicate_object_keys(pairs: list[tuple[str, Any]]) -> dict[str, Any]:
    """Reject duplicate JSON keys at every object depth during parsing."""
    value: dict[str, Any] = {}
    for key, item in pairs:
        if key in value:
            raise ValueError(f"duplicate JSON object key: {key!r}")
        value[key] = item
    return value


def validate_key_segment(key: Any) -> str:
    if not isinstance(key, str) or not KEY_SEGMENT_RE.fullmatch(key):
        raise ValueError("resource key segments must match ^[A-Za-z][A-Za-z0-9_]*$")
    return key


def flatten_strings(value: Any, prefix: str = "") -> list[tuple[str, str]]:
    """Return sorted dotted keys while rejecting ambiguous resource values."""
    if isinstance(value, str):
        if not prefix:
            raise ValueError("a string resource requires a key")
        return [(prefix, value)]
    if not isinstance(value, dict):
        raise ValueError(f"resource {prefix or '<root>'} must be an object or string")

    flattened: list[tuple[str, str]] = []
    seen_keys: set[str] = set()

    def visit(node: Any, current_prefix: str) -> None:
        if isinstance(node, str):
            if not current_prefix:
                raise ValueError("a string resource requires a key")
            if current_prefix in seen_keys:
                raise ValueError(f"duplicate flattened resource key: {current_prefix!r}")
            seen_keys.add(current_prefix)
            flattened.append((current_prefix, node))
            return
        if not isinstance(node, dict):
            raise ValueError(f"resource {current_prefix or '<root>'} must be an object or string")
        keys = [validate_key_segment(raw_key) for raw_key in node]
        for key in sorted(keys):
            child_prefix = f"{current_prefix}.{key}" if current_prefix else key
            visit(node[key], child_prefix)

    visit(value, prefix)
    return flattened


def vba_expression(text: str) -> list[str]:
    """Encode Unicode text as UTF-16 code units without source-file Unicode."""
    try:
        utf16 = text.encode("utf-16-le")
    except UnicodeEncodeError as error:
        raise ValueError("resource values must not contain lone surrogates") from error
    terms = [f"ChrW$(&H{int.from_bytes(utf16[index:index + 2], 'little'):04X})" for index in range(0, len(utf16), 2)]
    if not terms:
        return ["vbNullString"]
    return terms


def render_assignment(key: str, value: str) -> list[str]:
    terms = vba_expression(value)
    if terms == ["vbNullString"]:
        return [f'        Case "{key}"', "            NxString = vbNullString"]

    lines = [f'        Case "{key}"']
    chunks = [terms[index:index + 5] for index in range(0, len(terms), 5)]
    for index, chunk in enumerate(chunks):
        prefix = "            NxString = " if index == 0 else "                "
        suffix = " & _" if index < len(chunks) - 1 else ""
        lines.append(prefix + " & ".join(chunk) + suffix)
    return lines


def render_module(strings: dict[str, Any]) -> bytes:
    entries = flatten_strings({key: value for key, value in strings.items() if key != "schema_version"})
    lines = [
        f'Attribute VB_Name = "{MODULE_NAME}"',
        "Option Explicit",
        "",
        "Public Function NxString(ByVal key As String) As String",
        "    Select Case key",
    ]
    for key, value in entries:
        lines.extend(render_assignment(key, value))
    lines.extend([
        "        Case Else",
        "            NxString = vbNullString",
        "    End Select",
        "End Function",
        "",
    ])
    output = "\n".join(lines).encode("ascii")
    return output


def load_strings(path: Path) -> dict[str, Any]:
    with path.open("r", encoding="utf-8") as stream:
        strings = json.load(stream, object_pairs_hook=reject_duplicate_object_keys)
    if not isinstance(strings, dict) or type(strings.get("schema_version")) is not int or strings["schema_version"] != 1:
        raise ValueError("strings resource must be a schema_version 1 object")
    return strings


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--input", type=Path, default=DEFAULT_INPUT, help="UTF-8 strings JSON input")
    parser.add_argument("--output", type=Path, default=DEFAULT_OUTPUT, help="generated VBA output path")
    parser.add_argument("--stdout", action="store_true", help="write generated VBA to standard output")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    output = render_module(load_strings(args.input))
    if args.stdout:
        sys.stdout.buffer.write(output)
        return 0
    args.output.parent.mkdir(parents=True, exist_ok=True)
    args.output.write_bytes(output)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
