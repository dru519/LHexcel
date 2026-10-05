#!/usr/bin/env python3
"""Apply the approved r47 information-architecture contract migration."""

from __future__ import annotations

import json
from pathlib import Path


GROUPS = [
    "NX-GRP-MYEXCEL", "NX-GRP-SAVE", "NX-GRP-PRINT", "NX-GRP-COPY",
    "NX-GRP-INSERT", "NX-GRP-FILE", "NX-GRP-TEMPLATE", "NX-GRP-DATA",
    "NX-GRP-CELL-FIT", "NX-GRP-FORMULA", "NX-GRP-SHEET", "NX-GRP-VIEW",
    "NX-GRP-STYLE", "NX-GRP-UTIL", "NX-GRP-INFO",
]


def load(path: Path) -> dict:
    return json.loads(path.read_text(encoding="utf-8-sig"))


def dump(path: Path, value: dict) -> None:
    path.write_text(json.dumps(value, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")


def migrate_features(path: Path) -> None:
    contract = load(path)
    contract["ribbon_metadata"]["groups"] = GROUPS
    for section in ("required_feature_ids", "release_feature_ids"):
        values = contract.get("lineage", {}).get(section, [])
        contract.setdefault("lineage", {})[section] = [value for value in values if value != "NX-TPL-REGISTER-RANGE"]
    contract["features"] = [
        feature for feature in contract["features"] if feature.get("id") != "NX-TPL-REGISTER-RANGE"
    ]
    for feature in contract["features"]:
        if feature.get("id") == "NX-DATA-DUPLICATE-LIST":
            feature.setdefault("ribbon", {})["favorite_eligible"] = False
        if feature.get("id") == "NX-HANGUL-TABLE-SEND":
            feature.setdefault("privacy", {})["local_file_creation"] = True
            feature.setdefault("navigation", {})["transport"] = "embedded-HWPX"
    dump(path, contract)


def mutate_group(group: dict) -> None:
    group_id = group.get("id")
    if group_id == "NX-GRP-COPY":
        for item in group.get("controls", []):
            for nested in item.get("menu", []) if isinstance(item, dict) else []:
                if nested.get("kind") == "command_group":
                    nested["label"] = "선택 복붙"
    elif group_id == "NX-GRP-INSERT":
        for item in group.get("controls", []):
            for nested in item.get("menu", []) if isinstance(item, dict) else []:
                if nested.get("kind") == "command_group":
                    nested["label"] = "셀 작업"
    elif group_id == "NX-GRP-TEMPLATE":
        for item in group.get("controls", []):
            item["menu"] = [
                child for child in item.get("menu", [])
                if child.get("feature_id") != "NX-TPL-REGISTER-RANGE"
            ]
    elif group_id == "NX-GRP-SHEET":
        group["controls"] = [
            item for item in group.get("controls", [])
            if item.get("command_group_id") != "NX-CMD-GRP-ROW-COLUMN"
        ]
        for item in group.get("controls", []):
            if item.get("kind") == "command_group" and item.get("command_group_id") == "NX-CMD-GRP-SHEET":
                item["label"] = "관리·정리"
        group["label"] = "시트"
    elif group_id == "NX-GRP-VIEW":
        group["label"] = "보기창"
        for item in group.get("controls", []):
            if item.get("kind") == "command_group":
                item["label"] = "창 관리"
    elif group_id == "NX-GRP-UTIL":
        group["controls"] = [
            item for item in group.get("controls", [])
            if item.get("id") != "NX-CMD-MENU-INFO"
        ]
    elif group_id == "NX-GRP-STYLE":
        for item in group.get("controls", []):
            if item.get("id") == "NX-STYLE-HANGUL-MENU":
                menu = item.setdefault("menu", [])
                if not any(child.get("id") == "NX-STYLE-HANGUL-SETTINGS" for child in menu):
                    menu.append({
                        "kind": "button",
                        "id": "NX-STYLE-HANGUL-SETTINGS",
                        "target": "NX-ENTRY-HANGUL-SETTINGS",
                        "tag": "nx1|entry|NX-ENTRY-HANGUL-SETTINGS",
                        "label": "한글 표 설정",
                        "image_mso": "ControlProperties",
                        "callback": "NxRibbonExecute",
                    })


def migrate_ribbon(path: Path) -> None:
    contract = load(path)
    product = contract["surfaces"]["product"]
    groups = product["groups"]
    for group in groups:
        mutate_group(group)
    # Move the row/column group to a dedicated Cell Fit group.
    groups = [group for group in groups if group.get("id") not in {"NX-GRP-CELL-FIT", "NX-GRP-INFO"}]
    cell_fit = {
        "id": "NX-GRP-CELL-FIT",
        "label": "셀맞춤",
        "controls": [
            {
                "kind": "command_group",
                "id": "NX-CMD-MENU-ROW-COLUMN",
                "command_group_id": "NX-CMD-GRP-ROW-COLUMN",
                "label": "행·열 맞춤",
                "image_mso": "RowHeight",
            }
        ],
    }
    info = {
        "id": "NX-GRP-INFO",
        "label": "정보진단",
        "controls": [
            {
                "kind": "button",
                "id": "NX-INFO-ABOUT",
                "target": "NX-MGMT-FILE-INFO",
                "tag": "nx1|entry|NX-MGMT-FILE-INFO",
                "label": "내엑셀 버전",
                "image_mso": "Info",
                "callback": "NxRibbonExecute",
            },
            {
                "kind": "command_group",
                "id": "NX-CMD-MENU-INFO",
                "command_group_id": "NX-CMD-GRP-INFO-DIAGNOSTICS",
                "label": "정보·진단",
                "image_mso": "Diagnostics",
            },
        ],
    }
    insert_at = next(index for index, group in enumerate(groups) if group.get("id") == "NX-GRP-FORMULA")
    groups.insert(insert_at, cell_fit)
    groups.append(info)
    product["groups"] = groups
    dump(path, contract)


def migrate_commands(path: Path) -> None:
    contract = load(path)
    for group in contract.get("groups", []):
        if group.get("id") == "NX-CMD-GRP-ROW-COLUMN":
            group["ribbon_group_id"] = "NX-GRP-CELL-FIT"
            group["label_ko"] = "셀맞춤"
        elif group.get("id") == "NX-CMD-GRP-SHEET":
            group["label_ko"] = "시트 작업"
        elif group.get("id") == "NX-CMD-GRP-INFO-DIAGNOSTICS":
            group["ribbon_group_id"] = "NX-GRP-INFO"
            group["label_ko"] = "정보·진단"
        elif group.get("id") == "NX-CMD-GRP-CLIPBOARD":
            group["label_ko"] = "선택 복붙"
        elif group.get("id") == "NX-CMD-GRP-INSERT-DELETE-HIDE":
            group["label_ko"] = "셀 작업"
    for command in contract.get("commands", []):
        navigation = command.get("navigation", {})
        if command.get("group_id") == "NX-CMD-GRP-ROW-COLUMN":
            navigation["ribbon_group_id"] = "NX-GRP-CELL-FIT"
        elif command.get("group_id") == "NX-CMD-GRP-INFO-DIAGNOSTICS":
            navigation["ribbon_group_id"] = "NX-GRP-INFO"
        command["navigation"] = navigation
    dump(path, contract)


def main() -> int:
    root = Path(__file__).resolve().parents[1]
    migrate_features(root / "contracts/feature-contract.json")
    migrate_ribbon(root / "contracts/ribbon-shell-contract.json")
    migrate_commands(root / "contracts/command-contract.json")
    # Keep the generated resource aligned with the authoritative command contract.
    dump(root / "src/resources/commands.ko-KR.json", load(root / "contracts/command-contract.json"))
    print("r47 contract migration complete")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
