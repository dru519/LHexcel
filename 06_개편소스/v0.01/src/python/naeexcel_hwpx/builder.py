"""검증을 통과한 HWPX만 최종 경로로 승격한다."""

from __future__ import annotations

import hashlib
import os
from pathlib import Path
import tempfile
import zipfile

from .model import load_table_input
from .validator import validate_hwpx
from .xmlparts import package_parts


def build_hwpx(input_path: str | Path, output_path: str | Path) -> dict[str, object]:
    table = load_table_input(input_path)
    output = Path(output_path)
    output.parent.mkdir(parents=True, exist_ok=True)
    descriptor, temporary_name = tempfile.mkstemp(prefix=f".{output.stem}-", suffix=".hwpx", dir=output.parent)
    os.close(descriptor)
    temporary = Path(temporary_name)
    try:
        with zipfile.ZipFile(temporary, "w") as archive:
            archive.writestr("mimetype", b"application/hwp+zip", compress_type=zipfile.ZIP_STORED)
            for name, content in package_parts(table).items():
                archive.writestr(name, content, compress_type=zipfile.ZIP_DEFLATED)
        validation = validate_hwpx(temporary)
        if validation["status"] != "PASS":
            raise ValueError(f"생성된 HWPX 검증 실패: {validation}")
        os.replace(temporary, output)
        return {
            "status": "PASS", "path": str(output.resolve()),
            "sha256": hashlib.sha256(output.read_bytes()).hexdigest(),
            "rows": table.rows, "columns": table.columns, "merges": len(table.merges),
        }
    finally:
        temporary.unlink(missing_ok=True)

