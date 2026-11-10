#!/usr/bin/env python3
"""Draws the stopwatch badge for a share that has been on offer too long.

The share badge is a parcel; this one takes its place, in the same corner at the same
size, when a share is older than /itemshare expire <days>. A red ring around a pale
face reads as "time is up" at 16 px, which is the size that matters: the crown and the
hand are there for the 64 px source and for anyone who opens the file.

Run it from the repository root:

    python3 tools/make_expired_badge.py

It needs Python 3 and Pillow, and writes media/itemshare_expired.dds, which the addon
loads, plus a media/itemshare_expired.png preview beside it. The preview is not
tracked and the game never runs this file; it is here so the texture can be rebuilt
from source rather than redrawn by hand, the same way
tools/make_reserve_only_badge.py does for the reserve-only tag.
"""

from pathlib import Path
import struct

from PIL import Image, ImageDraw

SIZE = 64
# Drawn large and shrunk, so the circle and the hand are smooth at 64 px and survive
# the mip chain down to the 16 px the bags use.
SCALE = 4

FACE = (243, 238, 226, 255)
RED = (196, 74, 58, 255)
INK = (38, 30, 20, 255)


def draw_badge(size=SIZE):
    big = size * SCALE
    image = Image.new("RGBA", (big, big), (0, 0, 0, 0))
    draw = ImageDraw.Draw(image)

    cx = big // 2
    cy = int(big * 0.60)
    radius = int(big * 0.33)
    ink = max(2, int(big * 0.045))

    # Crown and stem first, so the body overlaps their base.
    draw.rounded_rectangle(
        (cx - int(big * 0.10), int(big * 0.07), cx + int(big * 0.10), int(big * 0.16)),
        radius=int(big * 0.03),
        fill=INK,
    )
    draw.rectangle(
        (cx - int(big * 0.035), int(big * 0.13), cx + int(big * 0.035), int(big * 0.29)),
        fill=INK,
    )

    # Face, dark rim, then the red ring that carries the warning.
    draw.ellipse((cx - radius, cy - radius, cx + radius, cy + radius), fill=FACE, outline=INK, width=ink)
    ring = radius - int(big * 0.075)
    draw.ellipse(
        (cx - ring, cy - ring, cx + ring, cy + ring),
        outline=RED,
        width=max(2, int(big * 0.035)),
    )

    # Quarter ticks, kept inside the ring so they do not thicken the outline.
    tick_outer = ring - int(big * 0.03)
    tick_inner = tick_outer - int(big * 0.05)
    for dx, dy in ((0, -1), (1, 0), (0, 1), (-1, 0)):
        draw.line(
            (cx + dx * tick_inner, cy + dy * tick_inner, cx + dx * tick_outer, cy + dy * tick_outer),
            fill=INK,
            width=max(2, int(big * 0.028)),
        )

    # One hand at roughly two o'clock: a watch that is running, not a clock face.
    hand = int(radius * 0.62)
    draw.line(
        (cx, cy, cx + int(hand * 0.5), cy - int(hand * 0.866)),
        fill=INK,
        width=max(3, int(big * 0.055)),
    )
    dot = max(2, int(big * 0.035))
    draw.ellipse((cx - dot, cy - dot, cx + dot, cy + dot), fill=RED)

    return image.resize((size, size), Image.LANCZOS)


def dds_bytes(image):
    """Uncompressed A8R8G8B8 DDS with a full mip chain.

    ESO reads this without any encoder, and the flat artwork loses nothing to DXT5.
    """
    levels = [image]
    while levels[-1].width > 1 or levels[-1].height > 1:
        width = max(1, levels[-1].width // 2)
        height = max(1, levels[-1].height // 2)
        levels.append(levels[-1].resize((width, height), Image.LANCZOS))

    header = bytearray()
    header += b"DDS "
    header += struct.pack("<I", 124)
    # CAPS | HEIGHT | WIDTH | PIXELFORMAT | LINEARSIZE
    header += struct.pack("<I", 0x1 | 0x2 | 0x4 | 0x1000 | 0x80000)
    header += struct.pack("<I", image.height)
    header += struct.pack("<I", image.width)
    header += struct.pack("<I", image.width * 4)
    header += struct.pack("<I", 0)
    header += struct.pack("<I", len(levels))
    header += b"\0" * (11 * 4)
    # Pixel format: RGB with alpha, 32 bpp, A8R8G8B8.
    header += struct.pack("<I", 32)
    header += struct.pack("<I", 0x41)
    header += b"\0" * 4
    header += struct.pack("<I", 32)
    header += struct.pack("<I", 0x00FF0000)
    header += struct.pack("<I", 0x0000FF00)
    header += struct.pack("<I", 0x000000FF)
    header += struct.pack("<I", 0xFF000000)
    # CAPS: TEXTURE | MIPMAP | COMPLEX
    header += struct.pack("<I", 0x1000 | 0x400000 | 0x8)
    header += b"\0" * (4 * 4)
    assert len(header) == 128, len(header)

    body = bytearray()
    for level in levels:
        pixels = level.convert("RGBA").load()
        for y in range(level.height):
            for x in range(level.width):
                red, green, blue, alpha = pixels[x, y]
                body += bytes((blue, green, red, alpha))

    return bytes(header + body)


def main():
    root = Path(__file__).resolve().parent.parent
    media = root / "media"
    badge = draw_badge()

    png_path = media / "itemshare_expired.png"
    dds_path = media / "itemshare_expired.dds"
    badge.save(png_path)
    dds_path.write_bytes(dds_bytes(badge))

    print(f"wrote {png_path.name} ({png_path.stat().st_size} bytes)")
    print(f"wrote {dds_path.name} ({dds_path.stat().st_size} bytes)")


if __name__ == "__main__":
    main()
