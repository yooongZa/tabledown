#!/usr/bin/env python3
"""Create a Windows .ico file for the PyInstaller build."""
from __future__ import annotations

import argparse
from pathlib import Path

from PIL import Image, ImageDraw


def load_or_create_icon(source: Path) -> Image.Image:
    if source.exists():
        return Image.open(source).convert("RGBA")

    image = Image.new("RGBA", (256, 256), (0, 0, 0, 0))
    draw = ImageDraw.Draw(image)
    draw.rounded_rectangle((32, 32, 224, 224), radius=40, fill=(18, 24, 38, 255))
    for x in (96, 160):
        draw.line((x, 48, x, 208), fill=(255, 255, 255, 230), width=10)
    for y in (96, 160):
        draw.line((48, y, 208, y), fill=(255, 255, 255, 230), width=10)
    return image


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--source", type=Path, required=True)
    parser.add_argument("--output", type=Path, required=True)
    args = parser.parse_args()

    args.output.parent.mkdir(parents=True, exist_ok=True)
    image = load_or_create_icon(args.source)
    image.save(args.output, sizes=[(16, 16), (24, 24), (32, 32), (48, 48), (64, 64), (128, 128), (256, 256)])
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
