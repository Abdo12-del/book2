#!/usr/bin/env python3
"""توليد صورة توضيحية تجريبية (تشريح الصفحة) لاختبار مسار الصور في الكتاب."""
import os

from PIL import Image, ImageDraw, ImageFont

import arabic as A

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT = os.path.join(ROOT, "book", "images", "page-anatomy.png")
FONT = os.path.join(ROOT, "assets", "fonts", "Amiri-Regular.ttf")
FONT_B = os.path.join(ROOT, "assets", "fonts", "Tajawal-Bold.ttf")

W, H = 1000, 700
BG = (255, 255, 255)
INK = (26, 26, 26)
ACCENT = (140, 109, 63)
RULE = (201, 205, 211)
FILL = (245, 241, 232)


def main() -> str:
    os.makedirs(os.path.dirname(OUT), exist_ok=True)
    img = Image.new("RGB", (W, H), BG)
    d = ImageDraw.Draw(img)
    f_s = ImageFont.truetype(FONT, 22)
    f_l = ImageFont.truetype(FONT_B, 26)

    # إطار الصفحة الخارجي (ورقة)
    d.rectangle([60, 60, W - 60, H - 60], outline=RULE, width=2)

    # منطقة المتن
    mx0, my0, mx1, my1 = 150, 150, W - 150, H - 150
    d.rectangle([mx0, my0, mx1, my1], fill=FILL, outline=ACCENT, width=2)

    # ترويسة
    d.line([mx0, 110, mx1, 110], fill=ACCENT, width=3)
    d.text((mx1, 92), A.shape("الترويسة"), font=f_s, fill=ACCENT, anchor="ra")

    # تذييل
    d.line([mx0, H - 110, mx1, H - 110], fill=ACCENT, width=3)
    d.text(((mx0 + mx1) // 2, H - 95), A.shape("رقم الصفحة (تذييل)"), font=f_s,
           fill=ACCENT, anchor="ma")

    # هوامش
    d.line([mx0, (my0 + my1) // 2, 100, (my0 + my1) // 2], fill=ACCENT, width=2)
    d.text((95, (my0 + my1) // 2), A.shape("هامش خارجي"), font=f_s, fill=ACCENT, anchor="ra")
    d.line([mx1, (my0 + my1) // 2 - 70, W - 100, (my0 + my1) // 2 - 70], fill=ACCENT, width=2)
    d.text((W - 95, (my0 + my1) // 2 - 70), A.shape("هامش التجليد (داخلي)"), font=f_s,
           fill=ACCENT, anchor="la")

    # أسطر المتن الوهمية
    y = my0 + 40
    d.text((mx1 - 30, y), A.shape("عنوان الفصل"), font=f_l, fill=(22, 50, 79), anchor="ra")
    y += 60
    widths = [mx1 - mx0 - 120, mx1 - mx0 - 180, mx1 - mx0 - 140, mx1 - mx0 - 220,
              mx1 - mx0 - 160, mx1 - mx0 - 200, mx1 - mx0 - 130]
    for w in widths:
        if y > my1 - 60:
            break
        d.line([mx1 - 30, y, mx1 - 30 - w, y], fill=(150, 150, 150), width=7)
        y += 32

    img.save(OUT, "PNG")
    print("تم إنشاء:", OUT)
    return OUT


if __name__ == "__main__":
    main()
