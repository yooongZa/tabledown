"""Generate derived menu bar icon PNGs from the base table icon.

Currently produces tablemark_menu_40_check.png — the 1-second "success flash"
variant shown after 'Copy table as XML' (see TabledownApp._flash_icon_success).

Why PIL over rendering assets/tablemark_menu_check.svg directly: these PNGs are
macOS *template* images, where color is ignored and only the alpha channel
draws. The off/check SVGs express the badge as a white "punch" stroke under a
black mark, but white pixels in a template PNG would render solid — the punch
must become *transparency*. So we take the committed base PNG and (1) erase a
wide band along the badge path (alpha -> 0), then (2) draw the narrow black
mark on top. tablemark_menu_40_off.png was produced the same way (it contains
zero white pixels), keeping every variant visually consistent.

Run:
    .venv/bin/python scripts/make_menu_icons.py
"""
from __future__ import annotations

from pathlib import Path

from PIL import Image, ImageChops, ImageDraw

ASSETS = Path(__file__).resolve().parent.parent / "assets" / "generated"
BASE = ASSETS / "tablemark_menu_40.png"
CHECK = ASSETS / "tablemark_menu_40_check.png"

# Checkmark polyline, bottom-right of the 40x40 canvas (grid spans x7-33, y9-31).
# Mirrors assets/tablemark_menu_check.svg: round caps. The punch is wide enough
# (10px) that the band crossing the grid's right border leaves no orphaned
# border fragment beside the check's elbow.
CHECK_POINTS = [(18, 25), (25, 32), (36, 15)]
PUNCH_WIDTH = 10
MARK_WIDTH = 4


def _draw_round_polyline(draw: ImageDraw.ImageDraw, points, width: int, fill) -> None:
    """Polyline with round joints and round end caps (PIL has no cap option)."""
    draw.line(points, fill=fill, width=width, joint="curve")
    radius = width / 2
    for x, y in (points[0], points[-1]):
        draw.ellipse((x - radius, y - radius, x + radius, y + radius), fill=fill)


def _remove_specks(img: Image.Image, min_pixels: int = 12) -> None:
    """Zero the alpha of tiny disconnected fragments left behind by the punch.

    The punch band crossing the grid can strand a few border pixels between
    itself and the canvas edge; those orphans read as dirt at menu bar size.
    """
    width, height = img.size
    alpha = img.getchannel("A").load()
    seen = [[False] * height for _ in range(width)]
    px = img.load()
    for sx in range(width):
        for sy in range(height):
            if seen[sx][sy] or alpha[sx, sy] == 0:
                continue
            component, stack = [], [(sx, sy)]
            seen[sx][sy] = True
            while stack:
                x, y = stack.pop()
                component.append((x, y))
                for nx, ny in ((x + 1, y), (x - 1, y), (x, y + 1), (x, y - 1)):
                    if 0 <= nx < width and 0 <= ny < height and not seen[nx][ny] \
                            and alpha[nx, ny] != 0:
                        seen[nx][ny] = True
                        stack.append((nx, ny))
            if len(component) < min_pixels:
                for x, y in component:
                    r, g, b, _ = px[x, y]
                    px[x, y] = (r, g, b, 0)


def make_check_icon() -> None:
    base = Image.open(BASE).convert("RGBA")

    # 1) Punch: erase a wide band along the check path from the alpha channel.
    punch = Image.new("L", base.size, 0)
    _draw_round_polyline(ImageDraw.Draw(punch), CHECK_POINTS, PUNCH_WIDTH, 255)
    base.putalpha(ImageChops.subtract(base.getchannel("A"), punch))
    _remove_specks(base)

    # 2) Mark: draw the black checkmark on top.
    _draw_round_polyline(ImageDraw.Draw(base), CHECK_POINTS, MARK_WIDTH, (0, 0, 0, 255))

    base.save(CHECK)
    print(f"wrote {CHECK}")


if __name__ == "__main__":
    make_check_icon()
