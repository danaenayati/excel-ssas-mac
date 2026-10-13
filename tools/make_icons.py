"""Draw the add-in icons. Run: uv run --with pillow python tools/make_icons.py

A white wireframe cube on an Excel-green rounded square, drawn large and scaled down
so the small sizes stay crisp.
"""

from __future__ import annotations

import math
from pathlib import Path

from PIL import Image, ImageDraw

OUT = Path(__file__).resolve().parent.parent / "addin" / "public" / "assets"
GREEN = (33, 115, 70, 255)
WHITE = (255, 255, 255, 255)
SIZES = (16, 32, 64, 80, 128)
CANVAS = 1024


def draw() -> Image.Image:
    image = Image.new("RGBA", (CANVAS, CANVAS), (0, 0, 0, 0))
    pen = ImageDraw.Draw(image)
    pen.rounded_rectangle((0, 0, CANVAS - 1, CANVAS - 1), radius=CANVAS * 0.2, fill=GREEN)

    centre = CANVAS / 2
    radius = CANVAS * 0.33
    corners = [
        (centre + radius * math.cos(math.radians(a)), centre + radius * math.sin(math.radians(a)))
        for a in (-90, -30, 30, 90, 150, 210)
    ]
    width = int(CANVAS * 0.07)
    pen.line(corners + [corners[0]], fill=WHITE, width=width, joint="curve")
    for index in (1, 3, 5):  # the three inner edges meeting at the near corner
        pen.line([(centre, centre), corners[index]], fill=WHITE, width=width)
    for x, y in corners + [(centre, centre)]:
        pen.ellipse((x - width / 2, y - width / 2, x + width / 2, y + width / 2), fill=WHITE)
    return image


def main() -> None:
    OUT.mkdir(parents=True, exist_ok=True)
    master = draw()
    for size in SIZES:
        master.resize((size, size), Image.LANCZOS).save(OUT / f"icon-{size}.png", optimize=True)
        print(f"wrote {OUT / f'icon-{size}.png'}")


if __name__ == "__main__":
    main()
