#!/usr/bin/env python3
"""Validate explicit MSForms client geometry and non-overlapping controls."""

from __future__ import annotations

import argparse
import json
from pathlib import Path


def rectangle(row: dict) -> tuple[float, float, float, float]:
    left = row["left"]
    top = row["top"]
    return left, top, left + row["width"], top + row["height"]


def overlaps(first: dict, second: dict) -> bool:
    a_left, a_top, a_right, a_bottom = rectangle(first)
    b_left, b_top, b_right, b_bottom = rectangle(second)
    return a_left < b_right and b_left < a_right and a_top < b_bottom and b_top < a_bottom


def mutually_exclusive(first: dict, second: dict) -> bool:
    group = first.get("visibility_group")
    return (
        isinstance(group, str)
        and bool(group)
        and group == second.get("visibility_group")
        and first.get("visibility_state") != second.get("visibility_state")
    )


def validate_form(path: Path) -> list[str]:
    errors: list[str] = []
    try:
        form = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as exc:
        return [f"{path}: invalid JSON: {exc}"]
    client_width = form.get("client_width")
    client_height = form.get("client_height")
    if not isinstance(client_width, int) or client_width <= 0:
        errors.append(f"{path}: client_width must be a positive integer")
    if not isinstance(client_height, int) or client_height <= 0:
        errors.append(f"{path}: client_height must be a positive integer")
    if errors:
        return errors

    controls = form.get("controls")
    if not isinstance(controls, list):
        return [f"{path}: controls must be a list"]
    names: set[str] = set()
    for row in controls:
        name = row.get("name")
        if not isinstance(name, str) or not name:
            errors.append(f"{path}: control name is missing")
            continue
        if name in names:
            errors.append(f"{path}: duplicate control name: {name}")
        names.add(name)
        try:
            left, top, right, bottom = rectangle(row)
        except (KeyError, TypeError):
            errors.append(f"{path}: incomplete control geometry: {name}")
            continue
        if left < 0 or top < 0 or right > client_width or bottom > client_height:
            errors.append(
                f"{path}: control outside client area: {name} "
                f"({left},{top},{right},{bottom}) > {client_width}x{client_height}"
            )

    solid = [row for row in controls if row.get("type") != "Label" and not row.get("decorative", False)]
    for index, first in enumerate(solid):
        for second in solid[index + 1 :]:
            if overlaps(first, second) and not mutually_exclusive(first, second):
                errors.append(f"{path}: controls overlap: {first.get('name')} / {second.get('name')}")

    regions = form.get("dynamic_regions", [])
    if not isinstance(regions, list):
        errors.append(f"{path}: dynamic_regions must be a list")
        return errors
    for region in regions:
        try:
            left, top, right, bottom = rectangle(region)
        except (KeyError, TypeError):
            errors.append(f"{path}: incomplete dynamic region geometry")
            continue
        if left < 0 or top < 0 or right > client_width or bottom > client_height:
            errors.append(f"{path}: dynamic region outside client area: {region.get('name')}")
    return errors


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--root", type=Path, required=True)
    args = parser.parse_args()
    paths = sorted(args.root.rglob("*.form.json"))
    if not paths:
        raise SystemExit("no form layouts found")
    errors = [error for path in paths for error in validate_form(path)]
    if errors:
        print(f"FAIL form geometry: {len(errors)} error(s)")
        for error in errors:
            print(f"- {error}")
        return 1
    print(f"PASS form geometry: {len(paths)} forms")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
