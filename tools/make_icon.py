"""Vytvoří assets/sheetdiff.ico (dva listy přes sebe, zelený a červený pruh). Potřebuje Pillow."""
import os

from PIL import Image, ImageDraw

OUT = os.path.join(os.path.dirname(__file__), "..", "assets", "sheetdiff.ico")


def sheet(d: ImageDraw.ImageDraw, x, y, w, h, s, band):
    d.rounded_rectangle((x, y, x + w, y + h), radius=int(4 * s), fill="#ffffff", outline="#2357c6",
                        width=max(1, int(3 * s)))
    for i in range(4):
        yy = y + int((10 + i * 9) * s)
        d.line((x + int(6 * s), yy, x + w - int(6 * s), yy), fill="#b9c4d8", width=max(1, int(2 * s)))
    d.rectangle((x + int(6 * s), y + int(19 * s), x + w - int(6 * s), y + int(25 * s)), fill=band)


def make(size: int) -> Image.Image:
    s = size / 64
    img = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    sheet(d, int(4 * s), int(4 * s), int(36 * s), int(46 * s), s, "#e05252")
    sheet(d, int(24 * s), int(14 * s), int(36 * s), int(46 * s), s, "#2f8a4f")
    return img


if __name__ == "__main__":
    os.makedirs(os.path.dirname(OUT), exist_ok=True)
    big = make(256)
    big.save(OUT, sizes=[(16, 16), (24, 24), (32, 32), (48, 48), (64, 64), (128, 128), (256, 256)])
    print(os.path.abspath(OUT))
