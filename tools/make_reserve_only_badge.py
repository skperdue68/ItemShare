#!/usr/bin/env python3
"""Draws the "share reserve, but nothing shared" badge.

The shared reserve badge is a parcel with an amber lock in the corner. This one is
the lock alone, so a bag holding a reserved item that is not shared at all still says
why the item is not on the list - and that there is something to clear.

Run it from the repository root:

    python3 tools/make_reserve_only_badge.py

It needs Python 3 and Pillow, and writes media/itemshare_reserve_only.dds, which the
addon loads, plus a media/itemshare_reserve_only.png preview beside it. The preview is
not tracked; the game never runs this file. It is here so the texture can be rebuilt
from source rather than redrawn by hand.
"""

from pathlib import Path
import struct

from PIL import Image, ImageDraw

SIZE = 64
AMBER = (232, 163, 61, 255)
AMBER_LIGHT = (247, 197, 108, 255)
INK = (38, 30, 20, 255)


def draw_badge(size=SIZE):
    image = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    draw = ImageDraw.Draw(image)

    # Rounded amber tag, the same amber as the corner lock on the shared badge.
    box = (8, 12, size - 8, size - 12)
    draw.rounded_rectangle(box, radius=12, fill=AMBER, outline=INK, width=3)

    # A lighter top edge so it reads as a tag rather than a flat square at 16px.
    draw.rounded_rectangle(
        (box[0] + 4, box[1] + 4, box[2] - 4, box[1] + 12),
        radius=6,
        fill=AMBER_LIGHT,
    )

    # Padlock: shackle over a body, in the dark ink the other badges use.
    centre_x = size // 2
    draw.arc(
        (centre_x - 11, 18, centre_x + 11, 38),
        start=180,
        end=360,
        fill=INK,
        width=4,
    )
    draw.rectangle((centre_x - 11, 28, centre_x - 7, 33), fill=INK)
    draw.rectangle((centre_x + 7, 28, centre_x + 11, 33), fill=INK)
    draw.rounded_rectangle((centre_x - 14, 32, centre_x + 14, 48), radius=3, fill=INK)
    draw.ellipse((centre_x - 3, 37, centre_x + 3, 43), fill=AMBER)
    draw.rectangle((centre_x - 1, 40, centre_x + 1, 45), fill=AMBER)

    return image


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

    png_path = media / "itemshare_reserve_only.png"
    dds_path = media / "itemshare_reserve_only.dds"
    badge.save(png_path)
    dds_path.write_bytes(dds_bytes(badge))

    print(f"wrote {png_path.name} ({png_path.stat().st_size} bytes)")
    print(f"wrote {dds_path.name} ({dds_path.stat().st_size} bytes)")


if __name__ == "__main__":
    main()
