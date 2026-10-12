#!/usr/bin/env python3
"""إنشاء صورة مُجمّعة (Contact Sheet) لصفحات الكتاب للمراجعة البصرية السريعة.

    python3 tools/preview.py build/book.pdf build/preview.png --cols 4 --dpi 70
"""
from __future__ import annotations

import argparse
import os

import pymupdf
from PIL import Image, ImageDraw


def make_preview(pdf_path: str, out_path: str, cols: int = 4, dpi: int = 70,
                 pages: str | None = None) -> str:
    doc = pymupdf.open(pdf_path)
    n = doc.page_count
    if pages:
        idxs = []
        for part in pages.split(","):
            if "-" in part:
                a, b = part.split("-")
                idxs += list(range(int(a) - 1, min(int(b), n)))
            else:
                idxs.append(int(part) - 1)
    else:
        idxs = list(range(n))

    tiles = []
    for i in idxs:
        pix = doc[i].get_pixmap(dpi=dpi)
        img = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
        tiles.append((i + 1, img))

    tw, th = tiles[0][1].size
    rows = (len(tiles) + cols - 1) // cols
    pad, label = 14, 22
    sheet = Image.new("RGB", (cols * (tw + pad) + pad, rows * (th + pad + label) + pad),
                      (238, 238, 238))
    d = ImageDraw.Draw(sheet)
    for k, (pno, img) in enumerate(tiles):
        r, c = divmod(k, cols)
        x = pad + c * (tw + pad)
        y = pad + r * (th + pad + label)
        sheet.paste(img, (x, y))
        d.rectangle([x, y, x + tw, y + th], outline=(180, 180, 180))
        d.text((x + 2, y + th + 4), f"صفحة {pno}", fill=(60, 60, 60))
    os.makedirs(os.path.dirname(out_path) or ".", exist_ok=True)
    sheet.save(out_path, "PNG")
    print(f"✓ صورة المراجعة: {out_path}  ({len(tiles)} صفحة، {sheet.size[0]}×{sheet.size[1]} بكسل)")
    return out_path


def main() -> int:
    ap = argparse.ArgumentParser(description="صورة مُجمّعة لصفحات PDF")
    ap.add_argument("pdf")
    ap.add_argument("out", nargs="?", default="build/preview.png")
    ap.add_argument("--cols", type=int, default=4)
    ap.add_argument("--dpi", type=int, default=70)
    ap.add_argument("--pages", help="مثال: 1-8 أو 1,3,5")
    a = ap.parse_args()
    make_preview(a.pdf, a.out, a.cols, a.dpi, a.pages)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
