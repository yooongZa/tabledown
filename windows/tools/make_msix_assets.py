#!/usr/bin/env python3
"""Generate the MSIX tile/logo PNG assets from the master app icon.

MSIX requires a set of square/wide logo PNGs referenced by AppxManifest.xml.
This renders them from the 1024px master icon (or a drawn fallback), centering
the artwork on a branded background so non-square tiles look intentional.

Usage:
    python make_msix_assets.py --source <master.png> --output-dir <Assets/>
"""
from __future__ import annotations

import argparse
from pathlib import Path

from PIL import Image, ImageDraw


BACKGROUND = (18, 24, 38, 255)  # #121826 — matches manifest BackgroundColor

# name -> (width, height). Names must match AppxManifest.xml references.
TARGETS = {
    "StoreLogo": (50, 50),
    "Square44x44Logo": (44, 44),
    "Square71x71Logo": (71, 71),
    "Square150x150Logo": (150, 150),
    "Square310x310Logo": (310, 310),
    "Wide310x150Logo": (310, 150),
}


def load_master(source: Path) -> Image.Image:
    if source.exists():
        return Image.open(source).convert("RGBA")

    # Fallback: draw the 3x2 table glyph (same motif as the menu-bar icon).
    image = Image.new("RGBA", (1024, 1024), (0, 0, 0, 0))
    draw = ImageDraw.Draw(image)
    draw.rounded_rectangle((128, 128, 896, 896), radius=160, fill=BACKGROUND)
    for x in (384, 640):
        draw.line((x, 192, x, 832), fill=(255, 255, 255, 235), width=40)
    for y in (384, 640):
        draw.line((192, y, 832, y), fill=(255, 255, 255, 235), width=40)
    return image


def render_tile(master: Image.Image, size: tuple[int, int]) -> Image.Image:
    width, height = size
    tile = Image.new("RGBA", size, BACKGROUND)

    # Fit the glyph into ~70% of the smaller dimension, centered.
    glyph = min(width, height)
    target = max(1, int(glyph * 0.7))
    art = master.resize((target, target), Image.LANCZOS)

    offset = ((width - target) // 2, (height - target) // 2)
    tile.alpha_composite(art, offset)
    return tile


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--source", type=Path, required=True)
    parser.add_argument("--output-dir", type=Path, required=True)
    args = parser.parse_args()

    args.output_dir.mkdir(parents=True, exist_ok=True)
    master = load_master(args.source)

    for name, size in TARGETS.items():
        tile = render_tile(master, size)
        out = args.output_dir / f"{name}.png"
        tile.save(out)
        print(f"wrote {out} ({size[0]}x{size[1]})")

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
