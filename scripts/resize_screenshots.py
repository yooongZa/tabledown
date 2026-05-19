#!/usr/bin/env python3
"""Convert raw macOS screenshots into Mac App Store 2880x1800 frames.

Mac App Store accepts screenshots at one of four exact sizes
(1280x800, 1440x900, 2560x1600, 2880x1800). Raw macOS captures rarely
match one of these aspect ratios, so we pad each input on a 2880x1800
canvas while preserving the menu-bar position (top-aligned).

Background colour is sampled from a corner pixel that should still
show the desktop, so the padded strip blends with the capture itself.

Usage:
    scripts/resize_screenshots.py <output-dir> <input.png> [<input.png> ...]

Example:
    scripts/resize_screenshots.py outputs/app-store-screenshots/v2 \\
        ~/Desktop/tabledown-menu-open.png \\
        ~/Desktop/tabledown-menu-bar.png \\
        ~/Desktop/tabledown-conversion.png
"""
from __future__ import annotations

import sys
from pathlib import Path

from PIL import Image

TARGET_W, TARGET_H = 2880, 1800
TARGET_RATIO = TARGET_W / TARGET_H


def sample_background(img: Image.Image) -> tuple[int, int, int]:
    """Pick a desktop pixel. Bottom-left corner is rarely occluded by app windows."""
    rgb = img.convert("RGB")
    return rgb.getpixel((10, img.height - 10))


def fit_to_canvas(src: Image.Image) -> Image.Image:
    bg = sample_background(src)
    src_ratio = src.width / src.height

    if src_ratio >= TARGET_RATIO:
        # Source is wider than 16:10 → fit by width, pad below the capture.
        new_w = TARGET_W
        new_h = round(src.height * (TARGET_W / src.width))
        resized = src.resize((new_w, new_h), Image.LANCZOS)
        canvas = Image.new("RGB", (TARGET_W, TARGET_H), bg)
        canvas.paste(resized, (0, 0))  # menu bar stays at the very top
    else:
        # Source is narrower → fit by height, pad left/right symmetrically.
        new_h = TARGET_H
        new_w = round(src.width * (TARGET_H / src.height))
        resized = src.resize((new_w, new_h), Image.LANCZOS)
        canvas = Image.new("RGB", (TARGET_W, TARGET_H), bg)
        canvas.paste(resized, ((TARGET_W - new_w) // 2, 0))

    return canvas


def main() -> int:
    if len(sys.argv) < 3:
        print(__doc__)
        return 1

    out_dir = Path(sys.argv[1]).expanduser().resolve()
    out_dir.mkdir(parents=True, exist_ok=True)

    for index, raw_path in enumerate(sys.argv[2:], start=1):
        src_path = Path(raw_path).expanduser().resolve()
        with Image.open(src_path) as src:
            framed = fit_to_canvas(src)
        out_path = out_dir / f"{index:02d}-{src_path.stem}.png"
        framed.save(out_path, "PNG", optimize=True)
        print(f"  {src_path.name}  →  {out_path}  ({framed.width}x{framed.height})")

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
