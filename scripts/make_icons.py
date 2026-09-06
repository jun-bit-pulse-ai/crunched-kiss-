#!/usr/bin/env python3
"""Write solid copper PNG icons for the Office manifest."""

from __future__ import annotations

import struct
import zlib
from pathlib import Path

RGB = (196, 132, 90)


def png(width: int, height: int, rgb: tuple[int, int, int]) -> bytes:
    def chunk(tag: bytes, data: bytes) -> bytes:
        return struct.pack(">I", len(data)) + tag + data + struct.pack(">I", zlib.crc32(tag + data) & 0xFFFFFFFF)

    raw = b"".join(b"\x00" + bytes(rgb) * width for _ in range(height))
    return b"".join(
        [
            b"\x89PNG\r\n\x1a\n",
            chunk(b"IHDR", struct.pack(">IIBBBBB", width, height, 8, 2, 0, 0, 0)),
            chunk(b"IDAT", zlib.compress(raw, 9)),
            chunk(b"IEND", b""),
        ]
    )


def main() -> None:
    assets = Path(__file__).resolve().parents[1] / "frontend" / "assets"
    assets.mkdir(parents=True, exist_ok=True)
    for size in (16, 32, 64, 80):
        (assets / f"icon-{size}.png").write_bytes(png(size, size, RGB))


if __name__ == "__main__":
    main()
