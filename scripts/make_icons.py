#!/usr/bin/env python3
"""Resize the Crunched brandmark into the icon sizes the Office manifest needs.

Run after changing frontend/assets/logo-source.png:

    python3 scripts/make_icons.py

Uses macOS `sips`, so there is no image dependency to install.
"""

from __future__ import annotations

import subprocess
import sys
from pathlib import Path

SIZES = (16, 32, 64, 80)


def main() -> int:
    assets = Path(__file__).resolve().parents[1] / "frontend" / "assets"
    source = assets / "logo-source.png"
    if not source.exists():
        print(f"missing {source}", file=sys.stderr)
        return 1

    for size in SIZES:
        target = assets / f"icon-{size}.png"
        subprocess.run(
            ["sips", "-s", "format", "png", "-z", str(size), str(size), str(source), "--out", str(target)],
            check=True,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
        )
        print(f"wrote {target.relative_to(assets.parents[1])}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
