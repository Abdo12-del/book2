#!/usr/bin/env python3
"""فحص صلاحية الخطوط قبل البناء.

لماذا هذا الفحص ضروري؟
  في مسار الـPDF نعتمد على «التشكيل السياقي» عبر arabic_reshaper، وهو يُخرج
  الأشكال السياقية (Arabic Presentation Forms). أي خط لا يحتوي هذه الأشكال في
  خريطة محارفه (cmap) سيُرسم ناقصاً (حروف تظهر فارغة). أشهر الأمثلة: Tajawal
  وAlmarai وNoto Kufi Arabic. أما في مسار الوورد فالمسؤول عن التشكيل هو Word
  نفسه عبر جداول GSUB، فتكفي تغطية الحروف العربية الأساسية.

الاستخدام:
    python3 tools/check_fonts.py                      # تقرير عن كل الخطوط
    python3 tools/check_fonts.py --source book/       # مقابل نص المخطوطة فعلياً
"""
from __future__ import annotations

import argparse
import glob
import os
import sys
import unicodedata

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import arabic as A  # noqa: E402

try:
    from fontTools.ttLib import TTFont as FTFont
except ImportError:
    print("تثبيت fonttools مطلوب: pip install fonttools")
    raise SystemExit(2)

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
FONT_DIR = os.path.join(ROOT, "assets", "fonts")

# نص اختباري يغطي كل الحروف العربية بحالاتها الأربع + التشكيل + الأرقام + علامات شائعة
SAMPLE = ("ابتثجحخدذرزسشصضطظعغفقكلمنهوي ءآأؤإئابةتثجح"
          "خذا رزسشصضطظعغفقكلمنهويًٌٍَُِّْ"
          "، ؛ ؟ ! ( ) [ ] { } « » - – — … . : , ; ? ! / \\ % ٠١٢٣٤٥٦٧٨٩ 0123456789 "
          "اللهم صل وسلم English text mixed 123")


def cmap_of(path: str) -> set[int]:
    t = FTFont(path, fontNumber=0, lazy=True)
    out: set[int] = set()
    for tb in t["cmap"].tables:
        out |= set(tb.cmap.keys())
    return out


def has_gsub(path: str) -> bool:
    return "GSUB" in FTFont(path, fontNumber=0, lazy=True)


def source_text(path: str) -> str:
    import re
    if os.path.isdir(path):
        buf = []
        for f in sorted(glob.glob(os.path.join(path, "**", "*.md"), recursive=True)):
            buf.append(open(f, encoding="utf-8").read())
        return "\n".join(buf)
    return open(path, encoding="utf-8").read()


def main() -> int:
    ap = argparse.ArgumentParser(description="فحص صلاحية الخطوط العربية")
    ap.add_argument("--source", default=os.path.join(ROOT, "book"),
                    help="ملف أو مجلد المخطوطة لفحص التغطية الفعلية")
    args = ap.parse_args()

    fonts = sorted(glob.glob(os.path.join(FONT_DIR, "*.ttf")))
    fonts += sorted(glob.glob("/usr/share/fonts/truetype/dejavu/*.ttf"))
    if not fonts:
        print("لا توجد خطوط في:", FONT_DIR)
        return 2

    shaped = A.shape(SAMPLE)
    extra = ""
    if os.path.exists(args.source):
        extra = source_text(args.source)
    corpus = shaped + A.shape(extra) if extra else shaped
    raw_corpus = SAMPLE + extra
    def is_arabic(cp: int) -> bool:
        return (0x0600 <= cp <= 0x06FF or 0xFB50 <= cp <= 0xFDFF
                or 0xFE70 <= cp <= 0xFEFF or 0x0750 <= cp <= 0x077F)

    needed_shaped = {ord(c) for c in corpus if ord(c) > 31}
    needed_raw = {ord(c) for c in raw_corpus if ord(c) > 31}
    ar_shaped = {c for c in needed_shaped if is_arabic(c)}
    ar_raw = {c for c in needed_raw if is_arabic(c)}
    other_shaped = needed_shaped - ar_shaped

    print(f"{'الخط':32s} {'عربي (PDF)':>11s} {'عربي (Word)':>12s} {'لاتيني/أخرى':>12s} {'GSUB':>5s}")
    print("-" * 78)
    rows = []
    for f in fonts:
        try:
            cm = cmap_of(f)
        except Exception as exc:  # noqa: BLE001
            print(f"{os.path.basename(f):32s}  خطأ: {exc}")
            continue
        ps = 100 * (1 - len(ar_shaped - cm) / max(1, len(ar_shaped)))
        pr = 100 * (1 - len(ar_raw - cm) / max(1, len(ar_raw)))
        po = 100 * (1 - len(other_shaped - cm) / max(1, len(other_shaped)))
        gsub = "✓" if has_gsub(f) else "✗"
        name = os.path.basename(f)
        print(f"{name:32s} {ps:10.1f}% {pr:11.1f}% {po:11.1f}% {gsub:>5s}")
        rows.append((name, ps, pr, po, cm))

    print("\nتقييم الصلاحية:")
    ok = pdf_ok = word_ok = 0
    for name, ps, pr, po, cm in rows:
        if ps >= 99.5:
            note = "صالح للمتن والعناوين في الـPDF والوورد"
            if po < 99:
                note += f" (تغطية لاتينية {po:.0f}% — يُستحسن خط لاتيني مصاحب)"
            print(f"  ✓ {name}: {note}")
            pdf_ok += 1
            word_ok += 1
            ok += 1
        elif pr >= 95:
            print(f"  ~ {name}: صالح للوورد فقط (Word يشكّل عبر GSUB) — "
                  f"غير صالح للـPDF بمسار التشكيل الحالي")
            word_ok += 1
        else:
            print(f"  ✗ {name}: لا يصلح للنص العربي (تغطية عربية {pr:.0f}%) — لاتيني فقط")
    print(f"\nالخلاصة: {pdf_ok} خط صالح للـPDF، و{word_ok} صالح للوورد، من أصل {len(rows)}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
