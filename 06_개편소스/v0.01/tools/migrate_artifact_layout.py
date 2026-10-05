#!/usr/bin/env python3
"""Transactionally normalize historical 내엑셀 artifact paths."""

from __future__ import annotations

import argparse
import ctypes
import json
import os
import re
from collections import Counter
from pathlib import Path
from typing import Any, Callable

from artifact_layout import (
    COMPLETE_RE,
    COMPLETE_ROOT_NAME,
    DISTRIBUTION_ROOT_NAME,
    ENHANCED_RE,
    IGNORED_NAMES,
    INTERNAL_RE,
    PACKAGE_DIRECTORY,
    inspect_layout,
    inventory,
    sha256_file,
    write_json,
)


FILE_ATTRIBUTE_HIDDEN = 0x2
FILE_ATTRIBUTE_REPARSE_POINT = 0x400

LEGACY_COMPLETE_RULES: tuple[
    tuple[re.Pattern[str], Callable[[re.Match[str]], str]], ...
] = (
    (
        re.compile(r"^내엑셀 v0\.01-r(?P<revision>\d+) 고도화$"),
        lambda match: f"내엑셀 v0.01_r{match.group('revision')} 완성본",
    ),
    (
        re.compile(r"^내엑셀 v0\.01_r(?P<revision>\d+)$"),
        lambda match: f"내엑셀 v0.01_r{match.group('revision')} 완성본",
    ),
    (
        re.compile(
            r"^내엑셀 v0\.01_r(?P<revision>\d+)_P0 포커스셀 수정본$"
        ),
        lambda match: f"내엑셀 v0.01_r{match.group('revision')}_P0 완성본",
    ),
)
LOOSE_INTERNAL_RE = re.compile(
    r"^내엑셀 v0\.01_r(?P<revision>\d+)(?P<qualifier>_[A-Za-z0-9]+)?_배포본\.xlam$"
)
LEGACY_DLL_DIR_RE = re.compile(
    r"^내엑셀 v0\.01_r(?P<revision>\d+)(?P<qualifier>_[A-Za-z0-9]+)?_DLL 배포본$"
)
LEGACY_DLL_ZIP_RE = re.compile(
    r"^내엑셀 v0\.01_r(?P<revision>\d+)(?P<qualifier>_[A-Za-z0-9]+)?_DLL 배포본\.zip$"
)
ENHANCED_PACKAGE_MEMBERS = (
    "payload",
    "deployment",
    "PACKAGE_SHA256SUMS",
    "README.md",
)


class MigrationHold(RuntimeError):
    """Raised before mutation when the current tree cannot be migrated safely."""


def _relative(path: Path, project_root: Path) -> str:
    return path.relative_to(project_root).as_posix()


def _path_is_within(path: Path, root: Path) -> bool:
    try:
        path.resolve(strict=False).relative_to(root.resolve(strict=True))
    except (FileNotFoundError, ValueError):
        return False
    return True


def _attributes(path: Path) -> int | None:
    try:
        return int(path.stat(follow_symlinks=False).st_file_attributes)
    except AttributeError:
        return None


def _is_reparse_point(path: Path) -> bool:
    attributes = _attributes(path)
    return path.is_symlink() or bool(
        attributes is not None and attributes & FILE_ATTRIBUTE_REPARSE_POINT
    )


def _assert_safe_roots(project_root: Path) -> tuple[Path, Path]:
    roots = (
        project_root / COMPLETE_ROOT_NAME,
        project_root / DISTRIBUTION_ROOT_NAME,
    )
    for root in roots:
        if not root.is_dir():
            raise MigrationHold(f"산출물 루트가 없습니다: {root}")
        if _is_reparse_point(root):
            raise MigrationHold(f"산출물 루트가 재분석 지점입니다: {root}")
        for current, directories, files in os.walk(root, followlinks=False):
            for name in [*directories, *files]:
                candidate = Path(current) / name
                if _is_reparse_point(candidate):
                    raise MigrationHold(
                        f"산출물 트리에 재분석 지점이 있습니다: {candidate}"
                    )
    return roots


def _operation(
    source: Path,
    destination: Path,
    kind: str,
    project_root: Path,
    *,
    create_parent: bool = False,
) -> dict[str, Any]:
    return {
        "source": _relative(source, project_root),
        "destination": _relative(destination, project_root),
        "kind": kind,
        "create_parent": create_parent,
    }


def _release_stem(match: re.Match[str]) -> str:
    qualifier = match.groupdict().get("qualifier") or ""
    return f"내엑셀 v0.01_r{int(match.group('revision'))}{qualifier}"


def _match_complete_name(name: str) -> str | None:
    for pattern, renderer in LEGACY_COMPLETE_RULES:
        match = pattern.fullmatch(name)
        if match is not None:
            return renderer(match)
    return None


def _validate_plan(project_root: Path, operations: list[dict[str, Any]]) -> None:
    complete_root = project_root / COMPLETE_ROOT_NAME
    distribution_root = project_root / DISTRIBUTION_ROOT_NAME
    allowed_roots = (complete_root, distribution_root)
    sources: set[str] = set()
    destinations: set[str] = set()
    planned_directory_destinations = {
        item["destination"]
        for item in operations
        if item["kind"] in {"complete-folder", "dll-folder"}
    }

    for item in operations:
        source_text = item["source"]
        destination_text = item["destination"]
        source = project_root / Path(source_text)
        destination = project_root / Path(destination_text)
        if source_text in sources:
            raise MigrationHold(f"중복 원본 경로입니다: {source_text}")
        if destination_text in destinations:
            raise MigrationHold(f"중복 목적 경로입니다: {destination_text}")
        sources.add(source_text)
        destinations.add(destination_text)

        if not any(_path_is_within(source, root) for root in allowed_roots):
            raise MigrationHold(f"허용 루트 밖 원본입니다: {source_text}")
        if not any(_path_is_within(destination, root) for root in allowed_roots):
            raise MigrationHold(f"허용 루트 밖 목적지입니다: {destination_text}")
        if not source.exists():
            raise MigrationHold(f"원본 경로가 없습니다: {source_text}")
        if destination.exists():
            raise MigrationHold(f"목적 경로가 이미 있습니다: {destination_text}")
        if item["create_parent"] and destination.parent.exists():
            raise MigrationHold(
                "도구가 생성해야 할 목적 폴더가 이미 있습니다: "
                + _relative(destination.parent, project_root)
            )
        if not item["create_parent"] and not destination.parent.exists():
            parent_text = _relative(destination.parent, project_root)
            if parent_text not in planned_directory_destinations:
                raise MigrationHold(
                    f"목적 부모 폴더가 없고 생성 계획도 없습니다: {parent_text}"
                )


def build_plan(project_root: Path) -> list[dict[str, Any]]:
    project_root = project_root.resolve()
    complete_root, distribution_root = _assert_safe_roots(project_root)

    complete_file_operations: list[dict[str, Any]] = []
    complete_folder_operations: list[dict[str, Any]] = []
    dll_package_operations: list[dict[str, Any]] = []
    dll_folder_operations: list[dict[str, Any]] = []
    internal_operations: list[dict[str, Any]] = []
    dll_zip_operations: list[dict[str, Any]] = []

    for child in sorted(complete_root.iterdir(), key=lambda item: item.name.casefold()):
        if child.name in IGNORED_NAMES:
            continue
        if not child.is_dir():
            raise MigrationHold(f"완성본 루트의 예상하지 못한 파일입니다: {child}")
        if COMPLETE_RE.fullmatch(child.name):
            continue
        destination_name = _match_complete_name(child.name)
        if destination_name is None:
            raise MigrationHold(f"알 수 없는 완성본 폴더 이름입니다: {child.name}")

        plain_match = re.fullmatch(r"내엑셀 v0\.01_r(?P<revision>\d+)", child.name)
        if plain_match is not None:
            old_xlam = child / f"내엑셀 v0.01_r{plain_match.group('revision')}.xlam"
            if old_xlam.exists():
                complete_file_operations.append(
                    _operation(
                        old_xlam,
                        child / "Product.xlam",
                        "complete-xlam",
                        project_root,
                    )
                )

        complete_folder_operations.append(
            _operation(
                child,
                complete_root / destination_name,
                "complete-folder",
                project_root,
            )
        )

    for child in sorted(
        distribution_root.iterdir(), key=lambda item: item.name.casefold()
    ):
        if child.name in IGNORED_NAMES:
            continue
        if child.is_dir():
            if INTERNAL_RE.fullmatch(child.name):
                continue
            standard_match = ENHANCED_RE.fullmatch(child.name)
            legacy_match = LEGACY_DLL_DIR_RE.fullmatch(child.name)
            if standard_match is None and legacy_match is None:
                raise MigrationHold(f"알 수 없는 배포본 폴더 이름입니다: {child.name}")

            package = child / PACKAGE_DIRECTORY
            direct_members = [child / name for name in ENHANCED_PACKAGE_MEMBERS]
            if package.exists():
                if any(member.exists() for member in direct_members):
                    raise MigrationHold(
                        f"DLL 패키지 계층이 중복되어 있습니다: {child.name}"
                    )
                nested_members = [package / name for name in ENHANCED_PACKAGE_MEMBERS]
                if not all(member.exists() for member in nested_members):
                    raise MigrationHold(
                        f"DLL 패키지 필수 항목이 부족합니다: {child.name}"
                    )
            else:
                if not all(member.exists() for member in direct_members):
                    raise MigrationHold(
                        f"DLL 패키지 필수 항목이 부족합니다: {child.name}"
                    )
                for member in direct_members:
                    dll_package_operations.append(
                        _operation(
                            member,
                            package / member.name,
                            "dll-package-member",
                            project_root,
                            create_parent=True,
                        )
                    )

            if legacy_match is not None:
                stem = _release_stem(legacy_match)
                dll_folder_operations.append(
                    _operation(
                        child,
                        distribution_root / f"{stem} DLL 배포본",
                        "dll-folder",
                        project_root,
                    )
                )
            continue

        internal_match = LOOSE_INTERNAL_RE.fullmatch(child.name)
        if internal_match is not None:
            stem = _release_stem(internal_match)
            folder = distribution_root / f"{stem} 내부망 배포본"
            internal_operations.append(
                _operation(
                    child,
                    folder / f"{stem} 내부망 배포본.xlam",
                    "internal-xlam",
                    project_root,
                    create_parent=True,
                )
            )
            continue

        zip_match = LEGACY_DLL_ZIP_RE.fullmatch(child.name)
        if zip_match is not None:
            stem = _release_stem(zip_match)
            folder = distribution_root / f"{stem} DLL 배포본"
            dll_zip_operations.append(
                _operation(
                    child,
                    folder / f"{stem} DLL 배포본.zip",
                    "dll-zip",
                    project_root,
                )
            )
            continue

        raise MigrationHold(f"배포본 루트의 예상하지 못한 파일입니다: {child}")

    operations = [
        *complete_file_operations,
        *complete_folder_operations,
        *dll_package_operations,
        *dll_folder_operations,
        *internal_operations,
        *dll_zip_operations,
    ]
    _validate_plan(project_root, operations)
    return operations


def _fingerprint(path: Path) -> dict[str, Any]:
    if path.is_file():
        return {"type": "file", "files": {".": sha256_file(path)}}
    files = {
        item.relative_to(path).as_posix(): sha256_file(item)
        for item in sorted(path.rglob("*"))
        if item.is_file()
    }
    return {"type": "directory", "files": files}


def _hash_multiset(records: list[dict[str, Any]]) -> Counter[str]:
    return Counter(
        record["sha256"] for record in records if record["type"] == "file"
    )


def _set_attributes(path: Path, attributes: int) -> None:
    if os.name != "nt":
        return
    set_attributes = ctypes.windll.kernel32.SetFileAttributesW
    set_attributes.argtypes = (ctypes.c_wchar_p, ctypes.c_uint32)
    set_attributes.restype = ctypes.c_int
    if not set_attributes(str(path), attributes):
        raise ctypes.WinError()


def _clear_hidden_artifacts(project_root: Path) -> list[dict[str, Any]]:
    changes: list[dict[str, Any]] = []
    if os.name != "nt":
        return changes
    for root_name in (COMPLETE_ROOT_NAME, DISTRIBUTION_ROOT_NAME):
        root = project_root / root_name
        for artifact in sorted(
            (path for path in root.rglob("*") if path.name not in IGNORED_NAMES),
            key=lambda item: item.name.casefold(),
        ):
            attributes = _attributes(artifact)
            if attributes is None or not attributes & FILE_ATTRIBUTE_HIDDEN:
                continue
            new_attributes = attributes & ~FILE_ATTRIBUTE_HIDDEN
            _set_attributes(artifact, new_attributes)
            changes.append(
                {
                    "path": _relative(artifact, project_root),
                    "before": attributes,
                    "after": new_attributes,
                }
            )
    return changes


def _restore_hidden_changes(
    project_root: Path, changes: list[dict[str, Any]]
) -> None:
    for change in reversed(changes):
        path = project_root / Path(change["path"])
        if path.exists():
            _set_attributes(path, int(change["before"]))


def _write_receipt(receipt: Path | None, report: dict[str, Any]) -> None:
    if receipt is not None:
        write_json(receipt, report)


def migrate(
    project_root: Path,
    *,
    apply: bool,
    receipt: Path | None = None,
    fail_after: int | None = None,
) -> dict[str, Any]:
    project_root = project_root.resolve()
    try:
        _assert_safe_roots(project_root)
        before_inventory = inventory(project_root)
        operations = build_plan(project_root)
    except MigrationHold as exc:
        report = {
            "schema_version": 1,
            "status": "HOLD",
            "project_root": str(project_root),
            "reason": str(exc),
            "operations": [],
        }
        _write_receipt(receipt, report)
        return report

    if not apply:
        report = {
            "schema_version": 1,
            "status": "DRY_RUN",
            "project_root": str(project_root),
            "operation_count": len(operations),
            "operations": operations,
            "before_inventory": before_inventory,
        }
        _write_receipt(receipt, report)
        return report

    applied: list[dict[str, Any]] = []
    created_directories: list[Path] = []
    hidden_changes: list[dict[str, Any]] = []
    operation_receipts: list[dict[str, Any]] = []
    try:
        for item in operations:
            source = project_root / Path(item["source"])
            destination = project_root / Path(item["destination"])
            if item["create_parent"] and not destination.parent.exists():
                destination.parent.mkdir()
                created_directories.append(destination.parent)
            if not destination.parent.is_dir():
                raise RuntimeError(
                    f"목적 부모 폴더를 사용할 수 없습니다: {destination.parent}"
                )

            before_fingerprint = _fingerprint(source)
            source.rename(destination)
            applied.append(item)
            after_fingerprint = _fingerprint(destination)
            preserved = before_fingerprint == after_fingerprint
            operation_receipts.append(
                {
                    **item,
                    "before_fingerprint": before_fingerprint,
                    "after_fingerprint": after_fingerprint,
                    "hashes_preserved": preserved,
                }
            )
            if not preserved:
                raise RuntimeError(f"이동 뒤 파일 지문이 다릅니다: {item['destination']}")
            if fail_after is not None and len(applied) >= fail_after:
                raise RuntimeError(f"시험용 실패 주입: {fail_after}")

        hidden_changes = _clear_hidden_artifacts(project_root)
        layout_report = inspect_layout(project_root)
        if layout_report["status"] != "PASS":
            raise RuntimeError(
                "이관 뒤 레이아웃 검사 실패: "
                + json.dumps(layout_report["errors"], ensure_ascii=False)
            )
        after_inventory = inventory(project_root)
        hashes_preserved = _hash_multiset(before_inventory) == _hash_multiset(
            after_inventory
        )
        if not hashes_preserved:
            raise RuntimeError("이관 전후 전체 파일 SHA-256 집합이 다릅니다.")

        report = {
            "schema_version": 1,
            "status": "PASS",
            "project_root": str(project_root),
            "operation_count": len(operations),
            "operations": operation_receipts,
            "hidden_attribute_changes": hidden_changes,
            "hashes_preserved": True,
            "before_inventory": before_inventory,
            "after_inventory": after_inventory,
            "layout": layout_report,
        }
        _write_receipt(receipt, report)
        return report
    except Exception as exc:  # rollback is deliberately broader than product logic
        _restore_hidden_changes(project_root, hidden_changes)
        rollback_errors: list[str] = []
        for item in reversed(applied):
            source = project_root / Path(item["source"])
            destination = project_root / Path(item["destination"])
            try:
                if not destination.exists():
                    raise RuntimeError(f"롤백 원본이 없습니다: {destination}")
                if source.exists():
                    raise RuntimeError(f"롤백 목적지가 이미 있습니다: {source}")
                destination.rename(source)
            except Exception as rollback_exc:
                rollback_errors.append(str(rollback_exc))
        for directory in reversed(created_directories):
            try:
                if directory.exists():
                    directory.rmdir()
            except Exception as rollback_exc:
                rollback_errors.append(str(rollback_exc))

        report = {
            "schema_version": 1,
            "status": "ROLLBACK_FAILED" if rollback_errors else "ROLLED_BACK",
            "project_root": str(project_root),
            "reason": str(exc),
            "operation_count": len(operations),
            "operations": operation_receipts,
            "rollback_errors": rollback_errors,
            "before_inventory": before_inventory,
            "after_rollback_inventory": inventory(project_root),
        }
        _write_receipt(receipt, report)
        return report


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Normalize historical 내엑셀 artifact paths transactionally."
    )
    parser.add_argument("--project-root", type=Path, required=True)
    parser.add_argument("--receipt", type=Path)
    parser.add_argument("--apply", action="store_true")
    args = parser.parse_args()

    report = migrate(
        args.project_root,
        apply=args.apply,
        receipt=args.receipt,
    )
    summary = {
        key: report.get(key)
        for key in ("status", "operation_count", "reason", "hashes_preserved")
        if key in report
    }
    print(json.dumps(summary, ensure_ascii=False, indent=2))
    return 0 if report["status"] in {"PASS", "DRY_RUN"} else 2


if __name__ == "__main__":
    raise SystemExit(main())
