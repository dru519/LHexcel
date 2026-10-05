"""내엑셀 HWPX 임시 파일의 명시적 24시간 보존 정책."""

from __future__ import annotations

from pathlib import Path
import time


OWNED_SUFFIXES = {".json", ".hwpx", ".tmp"}


def cleanup_expired(directory: str | Path, hours: float = 24.0) -> dict[str, object]:
    if hours < 0:
        raise ValueError("hours는 0 이상이어야 합니다.")
    root = Path(directory)
    cutoff = time.time() - hours * 3600
    removed = 0
    retained = 0
    if root.is_dir():
        for candidate in root.iterdir():
            if not candidate.is_file() or candidate.suffix.lower() not in OWNED_SUFFIXES:
                continue
            if candidate.stat().st_mtime >= cutoff:
                retained += 1
                continue
            try:
                candidate.unlink()
                removed += 1
            except OSError:
                retained += 1
    return {"status": "PASS", "removed": removed, "retained": retained, "directory": str(root)}

