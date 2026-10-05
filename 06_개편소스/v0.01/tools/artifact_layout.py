#!/usr/bin/env python3
"""Validate and inventory the canonical 내엑셀 release artifact layout."""

from __future__ import annotations

import argparse
import hashlib
import json
import re
from pathlib import Path
from typing import Any


SCHEMA_VERSION = 1
FILE_ATTRIBUTE_HIDDEN = 0x2
COMPLETE_ROOT_NAME = "01_완성본(고도화)"
DISTRIBUTION_ROOT_NAME = "02_배포본(고도화)"
ROOT_NAMES = (COMPLETE_ROOT_NAME, DISTRIBUTION_ROOT_NAME)
IGNORED_NAMES = {".DS_Store"}

RELEASE_PREFIX = r"내엑셀 v0\.01_r(?P<revision>\d+)(?P<qualifier>_[A-Za-z0-9]+)?"
COMPLETE_RE = re.compile(rf"^{RELEASE_PREFIX} 완성본$")
INTERNAL_RE = re.compile(rf"^{RELEASE_PREFIX} 내부망 배포본$")
ENHANCED_RE = re.compile(rf"^{RELEASE_PREFIX} DLL 배포본$")
PACKAGE_DIRECTORY = "패키지"

ENHANCED_REQUIRED = {
    PACKAGE_DIRECTORY: "directory",
    f"{PACKAGE_DIRECTORY}/payload": "directory",
    f"{PACKAGE_DIRECTORY}/deployment": "directory",
    f"{PACKAGE_DIRECTORY}/PACKAGE_SHA256SUMS": "file",
    f"{PACKAGE_DIRECTORY}/README.md": "file",
    f"{PACKAGE_DIRECTORY}/payload/Product.xlam": "file",
    f"{PACKAGE_DIRECTORY}/payload/NxHost32.dll": "file",
    f"{PACKAGE_DIRECTORY}/payload/NxHost64.dll": "file",
    f"{PACKAGE_DIRECTORY}/payload/NxCore32.dll": "file",
    f"{PACKAGE_DIRECTORY}/payload/NxCore64.dll": "file",
    f"{PACKAGE_DIRECTORY}/payload/SHA256SUMS": "file",
}


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _relative(path: Path, project_root: Path) -> str:
    return path.relative_to(project_root).as_posix()


def _file_attributes(path: Path) -> int | None:
    try:
        return int(path.stat(follow_symlinks=False).st_file_attributes)
    except AttributeError:
        return None


def inventory(project_root: Path) -> list[dict[str, Any]]:
    project_root = project_root.resolve()
    records: list[dict[str, Any]] = []
    for root_name in ROOT_NAMES:
        root = project_root / root_name
        if not root.exists():
            continue
        paths = [root, *root.rglob("*")]
        for path in sorted(paths, key=lambda item: _relative(item, project_root)):
            is_file = path.is_file()
            record: dict[str, Any] = {
                "path": _relative(path, project_root),
                "type": "file" if is_file else "directory",
                "attributes": _file_attributes(path),
            }
            if is_file:
                record["size"] = path.stat().st_size
                record["sha256"] = sha256_file(path)
            records.append(record)
    return records


def _add_error(
    errors: list[dict[str, str]],
    code: str,
    path: Path,
    project_root: Path,
    message: str,
) -> None:
    errors.append(
        {
            "code": code,
            "path": _relative(path, project_root),
            "message": message,
        }
    )


def _inspect_complete(
    folder: Path, errors: list[dict[str, str]], project_root: Path
) -> None:
    xlams = sorted(
        path.name
        for path in folder.iterdir()
        if path.is_file() and path.suffix.casefold() == ".xlam"
    )
    if xlams != ["Product.xlam"]:
        _add_error(
            errors,
            "COMPLETE_XLAM",
            folder,
            project_root,
            "완성본에는 Product.xlam이 정확히 1개 있어야 합니다.",
        )


def _release_stem(match: re.Match[str]) -> str:
    qualifier = match.group("qualifier") or ""
    return f"내엑셀 v0.01_r{int(match.group('revision'))}{qualifier}"


def _inspect_internal(
    folder: Path,
    match: re.Match[str],
    errors: list[dict[str, str]],
    project_root: Path,
) -> None:
    expected = f"{_release_stem(match)} 내부망 배포본.xlam"
    members = sorted(
        path.name for path in folder.iterdir() if path.name not in IGNORED_NAMES
    )
    if members != [expected] or not (folder / expected).is_file():
        _add_error(
            errors,
            "INTERNAL_XLAM",
            folder,
            project_root,
            f"내부망 배포본에는 {expected}만 있어야 합니다.",
        )


def _inspect_enhanced(
    folder: Path,
    match: re.Match[str],
    errors: list[dict[str, str]],
    project_root: Path,
) -> None:
    expected_zip = f"{_release_stem(match)} DLL 배포본.zip"
    zips = sorted(
        path.name
        for path in folder.iterdir()
        if path.is_file() and path.suffix.casefold() == ".zip"
    )
    if zips != [expected_zip]:
        _add_error(
            errors,
            "ENHANCED_ZIP",
            folder,
            project_root,
            f"DLL 배포본에는 {expected_zip}이 정확히 1개 있어야 합니다.",
        )

    for relative_name, expected_type in ENHANCED_REQUIRED.items():
        member = folder / Path(relative_name)
        valid = member.is_dir() if expected_type == "directory" else member.is_file()
        if not valid:
            _add_error(
                errors,
                "ENHANCED_MEMBER",
                member,
                project_root,
                f"DLL 배포본 필수 {expected_type} 항목이 없습니다.",
            )


def inspect_layout(project_root: Path) -> dict[str, Any]:
    project_root = project_root.resolve()
    errors: list[dict[str, str]] = []
    folder_counts = {"complete": 0, "internal": 0, "enhanced": 0}

    for root_name in ROOT_NAMES:
        root = project_root / root_name
        if not root.is_dir():
            errors.append(
                {
                    "code": "ROOT_MISSING",
                    "path": root_name,
                    "message": "산출물 루트 폴더가 없습니다.",
                }
            )
            continue

        for child in sorted(root.iterdir(), key=lambda item: item.name.casefold()):
            if child.name in IGNORED_NAMES:
                continue
            if not child.is_dir():
                _add_error(
                    errors,
                    "ROOT_FILE",
                    child,
                    project_root,
                    "산출물 루트 바로 아래에는 일반 파일을 둘 수 없습니다.",
                )
                continue

            if root_name == COMPLETE_ROOT_NAME:
                match = COMPLETE_RE.fullmatch(child.name)
                if match is None:
                    _add_error(
                        errors,
                        "FOLDER_NAME",
                        child,
                        project_root,
                        "완성본 폴더 이름이 표준 형식과 다릅니다.",
                    )
                    continue
                folder_counts["complete"] += 1
                _inspect_complete(child, errors, project_root)
                continue

            internal_match = INTERNAL_RE.fullmatch(child.name)
            if internal_match is not None:
                folder_counts["internal"] += 1
                _inspect_internal(child, internal_match, errors, project_root)
                continue

            enhanced_match = ENHANCED_RE.fullmatch(child.name)
            if enhanced_match is not None:
                folder_counts["enhanced"] += 1
                _inspect_enhanced(child, enhanced_match, errors, project_root)
                continue

            _add_error(
                errors,
                "FOLDER_NAME",
                child,
                project_root,
                "배포본 폴더 이름이 표준 형식과 다릅니다.",
            )

    inventory_records = inventory(project_root)
    for record in inventory_records:
        attributes = record["attributes"]
        if (
            attributes is not None
            and attributes & FILE_ATTRIBUTE_HIDDEN
            and Path(record["path"]).name not in IGNORED_NAMES
        ):
            errors.append(
                {
                    "code": "HIDDEN_ARTIFACT",
                    "path": record["path"],
                    "message": "숨김 산출물은 표준 배치에서 허용되지 않습니다.",
                }
            )

    return {
        "schema_version": SCHEMA_VERSION,
        "status": "PASS" if not errors else "FAIL",
        "project_root": str(project_root),
        "folder_counts": folder_counts,
        "errors": errors,
        "inventory": inventory_records,
    }


def write_json(path: Path, payload: dict[str, Any]) -> None:
    path = path.resolve()
    path.parent.mkdir(parents=True, exist_ok=True)
    temporary = path.with_name(f"{path.name}.tmp")
    temporary.write_text(
        json.dumps(payload, ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )
    temporary.replace(path)


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Validate the canonical 내엑셀 artifact folder layout."
    )
    parser.add_argument("--project-root", type=Path, required=True)
    parser.add_argument("--output", type=Path)
    args = parser.parse_args()

    report = inspect_layout(args.project_root)
    if args.output is not None:
        write_json(args.output, report)
    else:
        print(json.dumps(report, ensure_ascii=False, indent=2))
    return 0 if report["status"] == "PASS" else 1


if __name__ == "__main__":
    raise SystemExit(main())
