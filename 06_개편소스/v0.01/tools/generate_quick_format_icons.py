#!/usr/bin/env python3
"""Build original, font-independent 32/64 px quick-format Ribbon PNGs."""
from __future__ import annotations

import argparse
import struct
import zlib
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
# These are the only custom image resources accepted by the Ribbon packager.
ICON_NAMES = (
    "font-size-9", "font-size-10", "font-size-12", "font-size-15",
    "font-color-black", "font-color-red", "font-color-blue", "font-color-green",
    "fill-gray", "fill-light-red", "fill-light-yellow", "fill-light-green",
    "number-percent", "number-comma", "number-decimal",
)
QUICK_FORMAT_IMAGES = {
    "nx-quick-" + name: "images/nx-quick-" + name + ".png" for name in ICON_NAMES
}
FONT_COLORS = {
    "black": (0, 0, 0), "red": (255, 0, 0),
    "blue": (0, 0, 255), "green": (0, 128, 0),
}
FILL_COLORS = {
    "gray": (219, 219, 219), "light-red": (255, 202, 208),
    "light-yellow": (255, 236, 161), "light-green": (201, 240, 208),
}
# Original compact glyphs: no system fonts, Office assets or new dependencies.
GLYPHS = {
    "0": ("01110", "11011", "11011", "11011", "11011", "11011", "01110"),
    "1": ("00110", "01110", "00110", "00110", "00110", "00110", "01111"),
    "2": ("01110", "11011", "00011", "00110", "01100", "11000", "11111"),
    "5": ("11111", "11000", "11110", "00011", "00011", "11011", "01110"),
    "9": ("01110", "11011", "11011", "01111", "00011", "00110", "11100"),
    "A": ("01110", "11011", "11011", "11111", "11011", "11011", "11011"),
    "%": ("11001", "11010", "00010", "00100", "01000", "01011", "10011"),
    ",": ("00000", "00000", "00000", "00000", "00110", "00110", "01100"),
    ".": ("00000", "00000", "00000", "00000", "00000", "00110", "00110"),
}


def png_chunk(kind: bytes, data: bytes) -> bytes:
    return struct.pack(">I", len(data)) + kind + data + struct.pack(">I", zlib.crc32(kind + data))


def render_icon(image_id: str, size: int = 32) -> bytes:
    if image_id not in QUICK_FORMAT_IMAGES or size not in (32, 64):
        raise ValueError("quick-format image ID or size rejected")
    scale = size // 32
    pixels = bytearray(size * size * 4)

    def rectangle(x: int, y: int, width: int, height: int, color: tuple[int, int, int]) -> None:
        rgba = bytes((*color, 255))
        for row in range(y * scale, (y + height) * scale):
            start = (row * size + x * scale) * 4
            pixels[start:start + width * scale * 4] = rgba * (width * scale)

    def text(value: str, y: int, pixel: int, color: tuple[int, int, int], vertical: int | None = None) -> None:
        width = (len(value) * 6 - 1) * pixel
        left = (32 - width) // 2
        height = pixel if vertical is None else vertical
        for index, character in enumerate(value):
            for row, bits in enumerate(GLYPHS[character]):
                for col, bit in enumerate(bits):
                    if bit == "1":
                        rectangle(left + (index * 6 + col) * pixel, y + row * height, pixel, height, color)

    name = image_id.removeprefix("nx-quick-")
    if name.startswith("font-size-"):
        text(name.removeprefix("font-size-"), 8, 2, (28, 35, 45))
    elif name.startswith("font-color-"):
        color = FONT_COLORS[name.removeprefix("font-color-")]
        text("A", 3, 3, color)
        rectangle(4, 27, 24, 3, color)
    elif name.startswith("number-"):
        mark = {"number-percent": "%", "number-comma": ",", "number-decimal": "0.00"}[name]
        text(mark, 9 if mark == "0.00" else 5, 1 if mark == "0.00" else 3, (28, 35, 45), 2 if mark == "0.00" else 3)
    else:
        color = FILL_COLORS[name.removeprefix("fill-")]
        rectangle(3, 3, 26, 26, (88, 99, 112))
        rectangle(5, 5, 22, 22, color)
    raw = b"".join(b"\x00" + bytes(pixels[row * size * 4:(row + 1) * size * 4]) for row in range(size))
    header = struct.pack(">IIBBBBB", size, size, 8, 6, 0, 0, 0)
    return b"\x89PNG\r\n\x1a\n" + png_chunk(b"IHDR", header) + png_chunk(b"IDAT", zlib.compress(raw, 9)) + png_chunk(b"IEND", b"")


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    mode = parser.add_mutually_exclusive_group(required=True)
    mode.add_argument("--write", action="store_true")
    mode.add_argument("--check", action="store_true")
    args = parser.parse_args()
    stale = []
    for image_id, resource in QUICK_FORMAT_IMAGES.items():
        for size in (32, 64):
            path = ROOT / "src/ribbon" / resource
            if size == 64:
                path = path.with_stem(path.stem + "-64")
            expected = render_icon(image_id, size)
            if args.write:
                path.parent.mkdir(parents=True, exist_ok=True)
                path.write_bytes(expected)
            elif not path.is_file() or path.read_bytes() != expected:
                stale.append(str(path))
    if stale:
        raise SystemExit("stale quick-format icons: " + ", ".join(stale))
    print(f"PASS: {len(QUICK_FORMAT_IMAGES) * 2} original quick-format PNGs (32/64 px)")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
