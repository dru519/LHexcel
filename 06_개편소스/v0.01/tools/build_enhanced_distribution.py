#!/usr/bin/env python3
"""Assemble and atomically publish the r105 enhanced-DLL distribution."""
from __future__ import annotations

import argparse
import ctypes
import hashlib
import json
import os
import shutil
import subprocess
import tempfile
import zipfile
from pathlib import Path, PurePosixPath

from windows_powershell import windows_powershell_environment
from distribution_notice import NOTICE_TEXT


ROOT = Path(__file__).resolve().parents[1]
DEPLOYMENT_FILES = (
    "Install-NxEnhanced.ps1",
    "Register-NxHost.ps1",
    "Uninstall-NxEnhanced.ps1",
    "Verify-NxEnhancedPackage.ps1",
)
PAYLOAD_FILES = (
    "Product.xlam",
    "NxHost32.dll",
    "NxHost64.dll",
    "NxCore32.dll",
    "NxCore64.dll",
    "SHA256SUMS",
    "docs/r105-build.md",
)
PACKAGE_FILES = (
    "README.md",
    "PACKAGE_SHA256SUMS",
    "payload/Product.xlam",
    "payload/NxHost32.dll",
    "payload/NxHost64.dll",
    "payload/NxCore32.dll",
    "payload/NxCore64.dll",
    "payload/SHA256SUMS",
    "payload/docs/r105-build.md",
    "deployment/Install-NxEnhanced.ps1",
    "deployment/Register-NxHost.ps1",
    "deployment/Uninstall-NxEnhanced.ps1",
    "deployment/Verify-NxEnhancedPackage.ps1",
)


def file_attributes(path: Path) -> int | None:
    """Read fresh Win32 attributes; a non-Windows check is not native evidence."""
    if os.name != "nt":
        return None
    get_attributes = ctypes.WinDLL("kernel32", use_last_error=True).GetFileAttributesW
    get_attributes.argtypes = (ctypes.c_wchar_p,)
    get_attributes.restype = ctypes.c_uint32
    attributes = get_attributes(str(path))
    if attributes == 0xFFFFFFFF:
        raise ctypes.WinError(ctypes.get_last_error())
    return int(attributes)


def public_paths(root: Path):
    """Walk only the caller-owned publication, rejecting links before descent."""
    attributes = file_attributes(root)
    if root.is_symlink() or (attributes is not None and attributes & 0x400):
        raise RuntimeError(f"refusing reparse publication path: {root}")
    yield root, attributes
    if root.is_dir():
        for child in sorted(root.iterdir(), key=lambda path: (not path.is_dir(), path.name)):
            if child.name != ".DS_Store":
                yield from public_paths(child)


def clear_hidden_attribute(path: Path) -> None:
    # Exact-path attrib preserves other bits and survives the Parallels SMB
    # attribute propagation observed after copy/rename. Never use /S or /D.
    completed = subprocess.run(
        ["attrib.exe", "-H", str(path)], capture_output=True, text=False
    )
    if completed.returncode:
        raise RuntimeError(f"could not clear hidden publication attribute: {path}")


def verify_public_visibility(root: Path) -> dict[str, object]:
    checked = 0
    unavailable = False
    for path, attributes in public_paths(root):
        if attributes is None:
            unavailable = True
        else:
            checked += 1
            if attributes & 2:
                raise RuntimeError(f"hidden publication path: {path}")
    return {"status": "NOT_RUN" if unavailable else "PASS", "checked": checked, "hidden": 0 if not unavailable else None}


def make_public_visible(root: Path) -> None:
    # Preflight the entire owned tree before changing any attribute.
    paths = list(public_paths(root))
    for path, attributes in paths:
        if attributes is not None and attributes & 2:
            clear_hidden_attribute(path)
    verify_public_visibility(root)


def sha256_file(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def iter_files(root: Path, excluded_name: str | None = None) -> list[Path]:
    return sorted(
        (
            path
            for path in root.rglob("*")
            if path.is_file()
            and path.name != ".DS_Store"
            and (excluded_name is None or path.relative_to(root).as_posix() != excluded_name)
        ),
        key=lambda path: path.relative_to(root).as_posix().encode("utf-8"),
    )


def write_manifest(root: Path, manifest_name: str) -> None:
    rows = [
        f"{sha256_file(path)}  {path.relative_to(root).as_posix()}"
        for path in iter_files(root, manifest_name)
    ]
    (root / manifest_name).write_text("\n".join(rows) + "\n", encoding="utf-8", newline="\n")


def read_manifest(root: Path, manifest_name: str) -> dict[str, str]:
    manifest = root / manifest_name
    if not manifest.is_file():
        raise RuntimeError(f"missing hash manifest: {manifest}")
    records: dict[str, str] = {}
    previous: bytes | None = None
    for line in manifest.read_text(encoding="utf-8").splitlines():
        try:
            digest, relative = line.split("  ", 1)
        except ValueError as exc:
            raise RuntimeError(f"malformed hash line: {line}") from exc
        relative_path = PurePosixPath(relative)
        if (
            len(digest) != 64
            or any(character not in "0123456789abcdef" for character in digest)
            or relative_path.is_absolute()
            or ".." in relative_path.parts
            or "\\" in relative
            or relative in records
        ):
            raise RuntimeError(f"unsafe or malformed hash line: {line}")
        sort_key = relative.encode("utf-8")
        if previous is not None and previous >= sort_key:
            raise RuntimeError(f"hash manifest is not strictly sorted: {manifest}")
        records[relative] = digest
        previous = sort_key
    return records


def verify_manifest(root: Path, manifest_name: str) -> None:
    records = read_manifest(root, manifest_name)
    actual = [path.relative_to(root).as_posix() for path in iter_files(root, manifest_name)]
    if list(records) != actual:
        raise RuntimeError(f"{manifest_name} file inventory mismatch")
    for relative, digest in records.items():
        if sha256_file(root / Path(*PurePosixPath(relative).parts)) != digest:
            raise RuntimeError(f"SHA-256 mismatch: {relative}")


def assert_exact_inventory(root: Path, expected: tuple[str, ...]) -> None:
    actual = sorted(
        (path.relative_to(root).as_posix() for path in iter_files(root)),
        key=lambda value: value.encode("utf-8"),
    )
    wanted = sorted(expected, key=lambda value: value.encode("utf-8"))
    if actual != wanted:
        raise RuntimeError(f"package inventory mismatch: expected={wanted!r}, actual={actual!r}")


def assemble(complete_root: Path, source_root: Path, staging: Path, *, development_candidate: bool = False) -> None:
    if staging.exists() and any(staging.iterdir()):
        raise RuntimeError(f"refusing non-empty enhanced staging directory: {staging}")
    staging.mkdir(parents=True, exist_ok=True)
    assert_exact_inventory(complete_root, PAYLOAD_FILES)
    verify_manifest(complete_root, "SHA256SUMS")
    if not development_candidate:
        from distribution_source_profile import verify_enhanced
        verify_enhanced(complete_root / "Product.xlam")

    payload = staging / "payload"
    deployment = staging / "deployment"
    payload.mkdir()
    deployment.mkdir()
    for relative in PAYLOAD_FILES:
        source = complete_root / Path(*PurePosixPath(relative).parts)
        destination = payload / Path(*PurePosixPath(relative).parts)
        destination.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(source, destination)
    for name in DEPLOYMENT_FILES:
        source = source_root / "build" / name
        if not source.is_file():
            raise RuntimeError(f"missing enhanced deployment script: {source}")
        shutil.copy2(source, deployment / name)

    (staging / "README.md").write_text(
        ("# 내엑셀 v0.01_r105 DLL 개발 검증본 — 정식 배포 금지\n\n"
         "출시 전 시험용입니다. Product.xlam은 개발 완성본이며 배포 보호를 적용하지 않았습니다. "
         "서명·정식 출시·실사용 수용 검증 완료를 뜻하지 않습니다. 승인된 검증 환경에서만 사용하세요.\n\n"
         if development_candidate else "# 내엑셀 v0.01_r105 DLL 배포본\n\n")
        +
        "이 배포본은 Product.xlam과 Excel 비트수별 NxHost DLL을 포함합니다. "
        "실행 파일, PYD, Python 런타임은 포함하지 않습니다. 내부망 단일 파일 반입은 별도의 XLAM 배포본을 사용하세요.\n\n"
        "문서 비교와 아래한글 표·그림 전송은 DLL 연결이 필요합니다. 고도화판에는 해당 VBA 비교 알고리즘과 HWPX PowerShell 대체 실행을 포함하지 않습니다. DLL 연결이 안 되면 설치 안내를 확인하거나 내부망판을 사용하세요.\n\n"
        + (Path(__file__).resolve().parents[1] / "docs/dll-install-guide.md").read_text(encoding="utf-8")
        + "\n\n## 공개 소스 및 자료 취급 안내\n\n" + NOTICE_TEXT + "\n\n"
        + (Path(__file__).resolve().parents[1] / "docs/public-use-terms.md").read_text(encoding="utf-8"),
        encoding="utf-8",
        newline="\n",
    )
    write_manifest(staging, "PACKAGE_SHA256SUMS")
    assert_exact_inventory(staging, PACKAGE_FILES)
    verify_manifest(staging, "PACKAGE_SHA256SUMS")
    make_public_visible(staging)


def run_package_verifier(package_root: Path) -> None:
    script = package_root / "deployment" / "Verify-NxEnhancedPackage.ps1"
    command = [
        "powershell",
        "-NoProfile",
        "-NonInteractive",
        "-ExecutionPolicy",
        "Bypass",
        "-File",
        str(script),
        "-PackageRoot",
        str(package_root),
        "-Quiet",
    ]
    completed = subprocess.run(
        command,
        cwd=package_root,
        env=windows_powershell_environment(),
        capture_output=True,
        text=False,
    )
    if completed.returncode:
        raw = completed.stderr or completed.stdout or b"enhanced package verification failed"
        raise RuntimeError(raw.decode("utf-8", errors="replace").strip())


def write_archive(package_root: Path, destination_folder: Path, archive: Path) -> None:
    with zipfile.ZipFile(archive, "w", compression=zipfile.ZIP_DEFLATED, compresslevel=9) as handle:
        for path in iter_files(package_root):
            relative = path.relative_to(package_root).as_posix()
            info = zipfile.ZipInfo(f"{destination_folder.name}/{relative}", date_time=(1980, 1, 1, 0, 0, 0))
            info.compress_type = zipfile.ZIP_DEFLATED
            info.external_attr = 0o100644 << 16
            handle.writestr(info, path.read_bytes())


def extract_archive_safely(archive: Path, extraction_root: Path) -> Path:
    with zipfile.ZipFile(archive, "r") as handle:
        for info in handle.infolist():
            member = PurePosixPath(info.filename)
            if member.is_absolute() or ".." in member.parts or "\\" in info.filename:
                raise RuntimeError(f"unsafe archive member: {info.filename}")
        handle.extractall(extraction_root)
        roots = {PurePosixPath(info.filename).parts[0] for info in handle.infolist() if info.filename}
    if len(roots) != 1:
        raise RuntimeError("enhanced archive must contain exactly one package root")
    return extraction_root / next(iter(roots))


def publish(staging: Path, destination_folder: Path, destination_zip: Path) -> dict[str, object]:
    if destination_folder.exists() or destination_zip.exists():
        raise RuntimeError("refusing existing enhanced destination")
    if staging.parent.resolve() != destination_folder.parent.resolve():
        raise RuntimeError("enhanced staging and destination must use the same parent for atomic publication")
    destination_folder.parent.mkdir(parents=True, exist_ok=True)
    destination_zip.parent.mkdir(parents=True, exist_ok=True)
    for parent in {destination_folder.parent, destination_zip.parent}:
        attributes = file_attributes(parent)
        if attributes is not None and attributes & (2 | 0x400):
            # The existing parent is not an owned staging tree: do not clear
            # unrelated release folders or traverse a publication junction.
            raise RuntimeError(f"hidden or reparse publication parent: {parent}")
    make_public_visible(staging)
    run_package_verifier(staging)

    with tempfile.NamedTemporaryFile(
        prefix=f"staging-{destination_zip.stem}-", suffix=".tmp", dir=destination_zip.parent, delete=False
    ) as temporary:
        temporary_archive = Path(temporary.name)
    temporary_archive.unlink()
    try:
        write_archive(staging, destination_folder, temporary_archive)
        with tempfile.TemporaryDirectory(prefix="naeexcel-r105-archive-", dir=destination_zip.parent) as temporary:
            extracted = extract_archive_safely(temporary_archive, Path(temporary))
            run_package_verifier(extracted)
        staging.replace(destination_folder)
        archive_published = False
        try:
            temporary_archive.replace(destination_zip)
            archive_published = True
            # A rename may acquire Hidden on shared folders; verify the final
            # names, not just the pre-publication staging names or file hashes.
            make_public_visible(destination_folder)
            make_public_visible(destination_zip)
            visibility = {
                "folder": verify_public_visibility(destination_folder),
                "archive": verify_public_visibility(destination_zip),
            }
            run_package_verifier(destination_folder)
        except (OSError, RuntimeError) as failure:
            rollback_errors = []
            owned_publications = [(destination_folder, staging)]
            if archive_published:
                owned_publications.insert(0, (destination_zip, temporary_archive))
            for published, original in owned_publications:
                try:
                    published.replace(original)
                except OSError as rollback_error:
                    rollback_errors.append(f"{published}: {rollback_error}")
            if rollback_errors:
                raise RuntimeError("publication failed; rollback incomplete: " + "; ".join(rollback_errors)) from failure
            raise
        return {
            "status": "pass",
            "folder": str(destination_folder),
            "archive": str(destination_zip),
            "archive_sha256": sha256_file(destination_zip),
            "package_manifest_sha256": sha256_file(destination_folder / "PACKAGE_SHA256SUMS"),
            "published_atomically": True,
            "visibility": visibility,
        }
    finally:
        if temporary_archive.exists():
            temporary_archive.unlink()


def resolve_path(path: Path) -> Path:
    return path.resolve() if path.is_absolute() else (ROOT / path).resolve()


def main() -> int:
    parser = argparse.ArgumentParser(description="Build and atomically publish the r105 enhanced-DLL package")
    parser.add_argument("--complete-root", type=Path, required=True)
    parser.add_argument("--destination-root", type=Path, required=True)
    parser.add_argument("--destination-zip", type=Path, required=True)
    parser.add_argument("--development-candidate", action="store_true", help="Mark the hashed README as an unprotected development-only test package")
    args = parser.parse_args()
    complete_root = resolve_path(args.complete_root)
    destination_root = resolve_path(args.destination_root)
    destination_zip = resolve_path(args.destination_zip)
    if destination_root.parent.resolve() != destination_zip.parent.resolve():
        parser.error("enhanced folder and ZIP destinations must share one parent")
    if destination_root.exists() or destination_zip.exists():
        parser.error("refusing existing enhanced destination")
    destination_root.parent.mkdir(parents=True, exist_ok=True)
    staging = Path(tempfile.mkdtemp(prefix=f"staging-{destination_root.name}-", dir=destination_root.parent))
    try:
        assemble(complete_root, ROOT, staging, development_candidate=args.development_candidate)
        receipt = publish(staging, destination_root, destination_zip)
        print(json.dumps(receipt, ensure_ascii=False, indent=2))
        return 0
    except (OSError, RuntimeError, zipfile.BadZipFile) as exc:
        parser.error(str(exc))
    finally:
        if staging.exists():
            shutil.rmtree(staging, ignore_errors=True)


if __name__ == "__main__":
    raise SystemExit(main())
