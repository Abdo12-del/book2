#!/usr/bin/env python3
"""أداة البناء الموحّدة: مخطوطة واحدة ← ملف DOCX + ملف PDF.

الاستخدام:
    python3 tools/build.py                      # بناء كل شيء بالإعدادات الافتراضية
    python3 tools/build.py --only pdf           # PDF فقط
    python3 tools/build.py --theme design/theme.yaml --name mybook
    python3 tools/build.py --no-cover --no-toc
"""
from __future__ import annotations

import argparse
import os
import sys
import time

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import yaml  # noqa: E402

import build_docx  # noqa: E402
import build_pdf  # noqa: E402
from manuscript import load_blocks  # noqa: E402

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def load_book(book_yaml: str) -> tuple[dict, list]:
    with open(book_yaml, encoding="utf-8") as f:
        meta = yaml.safe_load(f) or {}
    spine = meta.pop("spine", []) or []
    blocks = []
    base = os.path.dirname(book_yaml)
    for rel in spine:
        path = os.path.join(base, rel)
        if not os.path.exists(path):
            print(f"  ! ملف غير موجود في الـspine: {path}")
            continue
        loaded = load_blocks(path)
        if not [b for b in loaded if b.t not in ("h", "toc")]:
            # ملف بلا متن (مثل ملف الفهرس: عنوان «المحتويات» + <!-- toc -->):
            # لا نُدرجه حتى لا يُنتج صفحة فصل فارغة بعنوان مكرّر للفهرس المولَّد.
            print(f"  · تخطّي {rel}: بلا متن")
            continue
        print(f"  ✓ {rel}: {len(loaded)} كتلة")
        blocks += loaded
    return meta, blocks


def main() -> int:
    ap = argparse.ArgumentParser(description="بناء الكتاب (DOCX + PDF)")
    ap.add_argument("--book", default=os.path.join(ROOT, "book", "book.yaml"))
    ap.add_argument("--theme", default=os.path.join(ROOT, "design", "theme.yaml"))
    ap.add_argument("--outdir", default=os.path.join(ROOT, "output"))
    ap.add_argument("--name", default="book")
    ap.add_argument("--only", choices=["docx", "pdf", "both"], default="both")
    ap.add_argument("--no-cover", action="store_true")
    ap.add_argument("--no-toc", action="store_true")
    args = ap.parse_args()

    os.makedirs(args.outdir, exist_ok=True)
    t0 = time.time()
    meta, blocks = load_book(args.book)
    print(f"إجمالي الكُتل: {len(blocks)}")

    outputs = []
    if args.only in ("docx", "both"):
        out = os.path.join(args.outdir, f"{args.name}.docx")
        build_docx.build_docx(blocks, meta, out, args.theme,
                              cover=not args.no_cover, toc=not args.no_toc)
        outputs.append(out)
        print(f"✓ DOCX → {out}  ({os.path.getsize(out)/1024:.0f} KB)")

    if args.only in ("pdf", "both"):
        out = os.path.join(args.outdir, f"{args.name}.pdf")
        build_pdf.build_pdf(blocks, meta, out, args.theme,
                            cover=not args.no_cover, toc=not args.no_toc)
        outputs.append(out)
        print(f"✓ PDF  → {out}  ({os.path.getsize(out)/1024:.0f} KB)")

    print(f"الزمن: {time.time()-t0:.1f} ثانية")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
