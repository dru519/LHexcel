#!/usr/bin/env python3
"""Classify and, when explicitly requested, preserve a Windows result lane."""

from __future__ import annotations

import argparse
import hashlib
import json
import os
import shutil
import sys
import tempfile
from pathlib import Path
from typing import Any

try:
    from classify_windows_result import classify_lane
except ImportError:  # pragma: no cover - allows direct execution from this directory
    from .classify_windows_result import classify_lane


ARCHIVABLE_VERDICTS = {"PASS", "DIAGNOSTIC"}


def _sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as source:
        for block in iter(lambda: source.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def _snapshot_identity(verdict: dict[str, Any]) -> tuple[Any, ...]:
    return tuple(
        verdict.get(key)
        for key in ("verdict", "suite", "run_id", "exit", "return_zip", "zip_sha256")
    )


def archive_lane(lane: str | Path, destination: str | Path) -> dict[str, Any]:
    """Classify *lane* and preserve its completed evidence in a new directory.

    The destination must not exist.  Only completed, contract-valid PASS or
    DIAGNOSTIC lanes are copied; the source lane is never modified.
    """

    root = Path(lane)
    dest = Path(destination)
    verdict = classify_lane(root)
    if verdict.get("verdict") not in ARCHIVABLE_VERDICTS:
        raise ValueError(f"refusing to archive {verdict.get('verdict')} lane")

    status_path = root / "status" / "run-status.txt"
    zip_name = verdict.get("return_zip")
    if not isinstance(zip_name, str) or zip_name not in {"PASS.zip", "DIAGNOSTIC.zip"}:
        raise ValueError("classified lane has no supported return ZIP")
    zip_path = root / "return" / zip_name

    # lexists also rejects a dangling symlink, preserving the no-overwrite rule.
    if os.path.lexists(dest):
        raise FileExistsError(f"destination already exists: {dest}")
    parent = dest.parent
    parent.mkdir(parents=True, exist_ok=True)
    snapshot_root = Path(tempfile.mkdtemp(prefix=f".{dest.name}.snapshot-", dir=parent))
    destination_created = False
    try:
        snapshot_status = snapshot_root / "status"
        snapshot_return = snapshot_root / "return"
        snapshot_status.mkdir()
        snapshot_return.mkdir()
        shutil.copy2(status_path, snapshot_status / status_path.name)
        launcher_output = root / "status" / "launcher-output.txt"
        if launcher_output.is_file():
            shutil.copy2(launcher_output, snapshot_status / launcher_output.name)
        snapshot_zip = snapshot_return / zip_path.name
        shutil.copy2(zip_path, snapshot_zip)

        expected_digest = verdict.get("zip_sha256")
        if not isinstance(expected_digest, str) or _sha256(snapshot_zip) != expected_digest:
            raise ValueError("lane changed during snapshot: return ZIP digest mismatch")
        snapshot_verdict = classify_lane(snapshot_root)
        if snapshot_verdict.get("verdict") not in ARCHIVABLE_VERDICTS:
            raise ValueError(
                f"lane changed during snapshot: {snapshot_verdict.get('verdict')}"
            )
        if _snapshot_identity(snapshot_verdict) != _snapshot_identity(verdict):
            raise ValueError("lane changed during snapshot: classification identity mismatch")

        # mkdir is the no-overwrite publication reservation.  It is attempted
        # only after the isolated snapshot has been digest-checked and fully
        # reclassified, and it fails if any late destination appeared.
        dest.mkdir()
        destination_created = True
        shutil.copy2(snapshot_status / status_path.name, dest / status_path.name)
        snapshot_output = snapshot_status / "launcher-output.txt"
        if snapshot_output.is_file():
            shutil.copy2(snapshot_output, dest / snapshot_output.name)
        published_zip = dest / snapshot_zip.name
        shutil.copy2(snapshot_zip, published_zip)
        if _sha256(published_zip) != expected_digest:
            raise ValueError("published return ZIP digest mismatch")

        judgment = dest / "판정.json"
        fd, temporary = tempfile.mkstemp(prefix=".판정-", suffix=".tmp", dir=dest)
        try:
            with os.fdopen(fd, "w", encoding="utf-8", newline="\n") as stream:
                json.dump(snapshot_verdict, stream, ensure_ascii=False, indent=2, sort_keys=True)
                stream.write("\n")
                stream.flush()
                os.fsync(stream.fileno())
            os.replace(temporary, judgment)
        finally:
            if os.path.exists(temporary):
                os.unlink(temporary)
    except Exception:
        if destination_created:
            shutil.rmtree(dest, ignore_errors=True)
        raise
    finally:
        shutil.rmtree(snapshot_root, ignore_errors=True)
    return snapshot_verdict


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("lane", type=Path)
    parser.add_argument("--destination", type=Path)
    args = parser.parse_args(argv)

    if args.destination is None:
        print(json.dumps(classify_lane(args.lane), ensure_ascii=False, sort_keys=True))
        return 0
    try:
        result = archive_lane(args.lane, args.destination)
    except (FileExistsError, OSError, ValueError) as exc:
        print(str(exc), file=sys.stderr)
        return 2
    print(json.dumps(result, ensure_ascii=False, sort_keys=True))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
