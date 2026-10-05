#!/usr/bin/env python3
"""Generate the closed r53 command contract and VBA registry."""

from __future__ import annotations

import argparse
import json
import os
import re
import tempfile
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
SOURCE = ROOT / "src" / "resources" / "commands.ko-KR.json"
CONTRACT = ROOT / "contracts" / "command-contract.json"
GENERATED = ROOT / "src" / "vba" / "ui" / "NxGeneratedCommandRegistry.bas"
REQUIRED_FIELDS = [
    "id", "label_ko", "aliases_ko", "description_ko", "image_mso", "search_terms_ko",
    "category_id", "group_id", "group_order", "sort_order", "input_context",
    "active_when", "mutates_document", "recovery_policy", "ribbon_visible",
    "taskpane_visible", "favorite_eligible", "shortcut_eligible", "handler",
    "args", "legacy_ids", "legacy_entrypoints", "navigation",
]
ALLOWED_HANDLERS = [
    "NxCmdClipboard", "NxCmdPrint", "NxCmdStyle", "NxCmdNumberFormat",
    "NxCmdSelection", "NxCmdRowsColumns", "NxCmdAlignment",
    "NxCmdInsertDelete", "NxCmdFilterSort", "NxCmdFormula", "NxCmdSheet",
    "NxCmdView", "NxCmdInfo", "NxCmdPivot",
]
COMMAND_ID = re.compile(r"^NX-CMD-[A-Z0-9-]+$")
CATEGORY_ID = re.compile(r"^NX-CAT-[A-Z0-9-]+$")
GROUP_ID = re.compile(r"^NX-CMD-GRP-[A-Z0-9-]+$")
ARGUMENT_KEY = re.compile(r"^[A-Z0-9_]+$")
IMAGE_MSO = re.compile(r"^[A-Za-z][A-Za-z0-9]+$")
RIBBON_GROUP_ID = re.compile(r"^NX-GRP-[A-Z0-9-]+$")
NAVIGATION_FIELDS = {
    "ribbon_group_id", "object_scope", "purpose_tags", "launch_surface",
    "mutation_scope", "selection_mode", "feedback_mode", "profile_availability",
}
INPUT_CONTEXTS = {
    "range_selection", "active_worksheet", "active_workbook", "active_window",
    "selection_or_clipboard", "active_data_region",
}


def vba(value: str) -> str:
    return value.replace('"', '""')


def validate_commands(commands: object) -> list[dict[str, object]]:
    if not isinstance(commands, list):
        raise ValueError("commands must be a list")
    validated: list[dict[str, object]] = []
    seen: set[str] = set()
    seen_order: set[tuple[str, int]] = set()
    for index, raw in enumerate(commands):
        if not isinstance(raw, dict):
            raise ValueError(f"command {index} must be an object")
        missing = [field for field in REQUIRED_FIELDS if field not in raw]
        if missing:
            raise ValueError(f"command {index} missing fields: {missing}")
        command_id = raw["id"]
        label = raw["label_ko"]
        category_id = raw["category_id"]
        group_id = raw["group_id"]
        args = raw["args"]
        handler = raw["handler"]
        if not isinstance(command_id, str) or COMMAND_ID.fullmatch(command_id) is None:
            raise ValueError(f"command {index} id is invalid")
        if command_id in seen:
            raise ValueError(f"duplicate command id: {command_id}")
        if not isinstance(label, str) or not label.strip():
            raise ValueError(f"command {command_id} label is empty")
        for field in ("aliases_ko", "search_terms_ko", "legacy_ids", "legacy_entrypoints"):
            values = raw[field]
            if not isinstance(values, list) or any(not isinstance(value, str) for value in values):
                raise ValueError(f"command {command_id} {field} must be a list of strings")
        if not isinstance(raw["description_ko"], str) or not raw["description_ko"].strip():
            raise ValueError(f"command {command_id} description is empty")
        image_mso = raw["image_mso"]
        if not isinstance(image_mso, str) or IMAGE_MSO.fullmatch(image_mso) is None:
            raise ValueError(f"command {command_id} image_mso is invalid")
        if image_mso in {"MacroPlay", "ViewMacros"}:
            raise ValueError(f"command {command_id} image_mso is generic")
        if not isinstance(category_id, str) or CATEGORY_ID.fullmatch(category_id) is None:
            raise ValueError(f"command {command_id} category is invalid")
        if not isinstance(group_id, str) or GROUP_ID.fullmatch(group_id) is None:
            raise ValueError(f"command {command_id} group is invalid")
        if isinstance(raw["group_order"], bool) or not isinstance(raw["group_order"], int) or raw["group_order"] <= 0:
            raise ValueError(f"command {command_id} group_order must be a positive integer")
        if isinstance(raw["sort_order"], bool) or not isinstance(raw["sort_order"], int) or raw["sort_order"] <= 0:
            raise ValueError(f"command {command_id} sort_order must be a positive integer")
        if raw["input_context"] not in INPUT_CONTEXTS:
            raise ValueError(f"command {command_id} input_context is invalid")
        if raw["active_when"] != "excel_ready":
            raise ValueError(f"command {command_id} active_when is invalid")
        for field in ("ribbon_visible", "taskpane_visible", "favorite_eligible", "shortcut_eligible"):
            if not isinstance(raw[field], bool):
                raise ValueError(f"command {command_id} {field} must be boolean")
        if (
            not isinstance(args, list)
            or len(args) != 1
            or not isinstance(args[0], str)
            or ARGUMENT_KEY.fullmatch(args[0]) is None
        ):
            raise ValueError(f"command {command_id} args must contain one safe argument key")
        if not isinstance(raw["mutates_document"], bool):
            raise ValueError(f"command {command_id} mutates_document must be boolean")
        recovery_policy = raw["recovery_policy"]
        if not isinstance(recovery_policy, str) or not recovery_policy.strip():
            raise ValueError(f"command {command_id} recovery_policy must be a non-empty string")
        if raw["mutates_document"] and recovery_policy == "none":
            raise ValueError(f"command {command_id} mutation cannot use recovery_policy=none")
        if handler not in ALLOWED_HANDLERS:
            raise ValueError(f"command {command_id} handler is not allowed: {handler}")
        navigation = raw["navigation"]
        if not isinstance(navigation, dict) or set(navigation) != NAVIGATION_FIELDS:
            raise ValueError(f"command {command_id} navigation fields are invalid")
        if RIBBON_GROUP_ID.fullmatch(str(navigation["ribbon_group_id"])) is None:
            raise ValueError(f"command {command_id} navigation ribbon group is invalid")
        for field in ("object_scope", "purpose_tags"):
            values = navigation[field]
            if not isinstance(values, list) or not values or any(not isinstance(value, str) or not value for value in values):
                raise ValueError(f"command {command_id} navigation {field} is invalid")
        if navigation["launch_surface"] not in {"direct", "dialog"}:
            raise ValueError(f"command {command_id} navigation launch surface is invalid")
        if navigation["mutation_scope"] != ("document" if raw["mutates_document"] else "none"):
            raise ValueError(f"command {command_id} navigation mutation scope is invalid")
        if navigation["selection_mode"] not in {"none", "optional", "required"}:
            raise ValueError(f"command {command_id} navigation selection mode is invalid")
        if navigation["feedback_mode"] not in {"inline", "completion"}:
            raise ValueError(f"command {command_id} navigation feedback mode is invalid")
        if navigation["profile_availability"] != ["internal-xlam", "enhanced-dll"]:
            raise ValueError(f"command {command_id} navigation profile availability is invalid")
        seen.add(command_id)
        order_key = (group_id, int(raw["sort_order"]))
        if order_key in seen_order:
            raise ValueError(f"duplicate group/sort order: {group_id}/{raw['sort_order']}")
        seen_order.add(order_key)
        validated.append(dict(raw))
    return validated


def load_source() -> tuple[dict[str, object], list[dict[str, object]]]:
    document = json.loads(SOURCE.read_text(encoding="utf-8"))
    if document.get("schema_version") != 2:
        raise ValueError("command source schema version rejected")
    if document.get("allowed_handlers") != ALLOWED_HANDLERS:
        raise ValueError("command source handler set rejected")
    if document.get("required_fields") != REQUIRED_FIELDS:
        raise ValueError("command source required field set rejected")
    return document, validate_commands(document.get("commands"))


def render_contract(document: dict[str, object]) -> bytes:
    return (json.dumps(document, ensure_ascii=False, indent=2) + "\n").encode("utf-8")


def render_vba(commands: list[dict[str, object]]) -> str:
    lines = [
        'Attribute VB_Name = "NxGeneratedCommandRegistry"',
        "Option Explicit",
        "",
        "Public Sub NxRegisterGeneratedCommands(ByVal registry As CNxCommandRegistry)",
        "    ' Each generated row carries the sealed shortcutEligible decision.",
    ]
    for row in commands:
        lines.append(
            '    registry.Add "{id}", "{label}", "{category}", "{group}", "{handler}", "{argument}", {mutates}, {shortcut}, "{recovery}"'.format(
                id=row["id"],
                label=vba(str(row["label_ko"])),
                category=row["category_id"],
                group=row["group_id"],
                handler=row["handler"],
                argument=row["args"][0],
                mutates="True" if row["mutates_document"] else "False",
                shortcut="True" if row["shortcut_eligible"] else "False",
                recovery=vba(str(row["recovery_policy"])),
            )
        )
    lines.extend(
        [
            "End Sub",
            "",
            "Public Function NxGeneratedCommandCount() As Long",
            f"    NxGeneratedCommandCount = {len(commands)}",
            "End Function",
            "",
        ]
    )
    lines += ['Public Function NxRibbonOnlyCommandMetadata(ByVal routeKey As String) As String', '    Select Case routeKey']
    for row in commands:
        if not row['taskpane_visible'] and (row['args'][0].startswith('NX_STYLE_') or row['args'][0] == 'RB_EDIT_NUMBERFORMAT_NUMBER'):
            nav = row['navigation']
            value = '|'.join([row['input_context'], nav['selection_mode'], nav['launch_surface'], nav['mutation_scope']])
            lines.append(f'        Case "command:{row["id"]}": NxRibbonOnlyCommandMetadata = "{vba(value)}"')
    lines += ['    End Select', 'End Function', '']
    return "\n".join(lines)


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


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    mode = parser.add_mutually_exclusive_group(required=True)
    mode.add_argument("--write", action="store_true")
    mode.add_argument("--check", action="store_true")
    args = parser.parse_args()
    document, commands = load_source()
    outputs = {
        CONTRACT: render_contract(document),
        GENERATED: render_vba(commands).encode("utf-8"),
    }
    if args.check:
        stale = [path for path, expected in outputs.items() if not path.is_file() or path.read_bytes() != expected]
        if stale:
            raise SystemExit("stale command surface: " + ", ".join(str(path) for path in stale))
        return 0
    for path, content in outputs.items():
        atomic_write(path, content)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
