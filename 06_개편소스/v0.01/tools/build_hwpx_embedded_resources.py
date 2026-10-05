"""Build the deterministic VBA module that carries HWPX runtime resources."""

from __future__ import annotations

import argparse
import base64
import hashlib
from pathlib import Path


SCRIPT_RELATIVE = Path("tools/lhexcel_hwpx_table_export.ps1")
TEMPLATE_RELATIVE = Path("templates/표.hwpx")
MODULE_RELATIVE = Path("src/vba/features/hangul/hwpx/adapters/NxHwpxEmbeddedResources.bas")


def _record(path: Path) -> tuple[bytes, int, str]:
    payload = path.read_bytes()
    return payload, len(payload), hashlib.sha256(payload).hexdigest()


def _string_function(name: str, value: str) -> str:
    return f'Public Function {name}() As String\n    {name} = "{value}"\nEnd Function'


def _long_function(name: str, value: int) -> str:
    return f"Public Function {name}() As Long\n    {name} = {value}\nEnd Function"


def _payload_function(name: str, payload: bytes) -> str:
    encoded = base64.b64encode(payload).decode("ascii")
    chunks = [encoded[index : index + 76] for index in range(0, len(encoded), 76)]
    lines = [f"Public Function {name}() As String", "    Dim text As String"]
    lines.extend(f'    text = text & "{chunk}"' for chunk in chunks)
    lines.extend([f"    {name} = text", "End Function"])
    return "\n".join(lines)


def render(root: Path) -> str:
    script, script_size, script_sha = _record(root / SCRIPT_RELATIVE)
    template, template_size, template_sha = _record(root / TEMPLATE_RELATIVE)
    version_seed = hashlib.sha256((script_sha + template_sha).encode("ascii")).hexdigest()[:12]
    sections = [
        'Attribute VB_Name = "NxHwpxEmbeddedResources"',
        "Option Explicit",
        _string_function("NxHwpxEmbeddedResourceVersion", f"v2026_08_22_r86_{version_seed}"),
        _long_function("NxHwpxEmbeddedScriptSize", script_size),
        _string_function("NxHwpxEmbeddedScriptSha256", script_sha),
        _payload_function("NxHwpxEmbeddedScriptBase64", script),
        _long_function("NxHwpxEmbeddedTemplateSize", template_size),
        _string_function("NxHwpxEmbeddedTemplateSha256", template_sha),
        _payload_function("NxHwpxEmbeddedTemplateBase64", template),
    ]
    return "\n\n".join(sections) + "\n"


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--root", type=Path, required=True)
    parser.add_argument("--check", action="store_true")
    args = parser.parse_args()
    root = args.root.resolve()
    destination = root / MODULE_RELATIVE
    expected = render(root).encode("utf-8")
    if args.check:
        if not destination.is_file() or destination.read_bytes() != expected:
            raise SystemExit("generated HWPX embedded resource module is stale")
        return 0
    destination.parent.mkdir(parents=True, exist_ok=True)
    destination.write_bytes(expected)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
