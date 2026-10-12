#!/usr/bin/env python3
"""جلب الخطوط العربية (ترخيص OFL) من GitHub عبر واجهة API.

المضيف raw.githubusercontent.com غير مسموح في هذه البيئة، لذا نستخدم:
  - /repos/{o}/{r}/contents/{path}  للملفات <= 1MB (يُعيد base64 مباشرة)
  - /repos/{o}/{r}/git/blobs/{sha}  للملفات الأكبر
"""
from __future__ import annotations

import base64
import json
import os
import sys
import urllib.request

API = "https://api.github.com"
DEST = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "assets", "fonts")

# (repo, path في المستودع, اسم الملف عند الحفظ)
FONTS = [
    # Amiri - نسخي كلاسيكي (بديل أدبي؛ لم يبقَ الافتراضي بعد تحوّل الكتاب إلى Sans) (OFL)
    ("aliftype/amiri", "fonts/Amiri-Regular.ttf", "Amiri-Regular.ttf"),
    ("aliftype/amiri", "fonts/Amiri-Bold.ttf", "Amiri-Bold.ttf"),
    ("aliftype/amiri", "fonts/Amiri-Italic.ttf", "Amiri-Italic.ttf"),
    ("aliftype/amiri", "fonts/Amiri-BoldItalic.ttf", "Amiri-BoldItalic.ttf"),
    # Noto Naskh Arabic - خط نسخي حديث للمتن (OFL)
    ("notofonts/noto-fonts", "hinted/ttf/NotoNaskhArabic/NotoNaskhArabic-Regular.ttf", "NotoNaskhArabic-Regular.ttf"),
    ("notofonts/noto-fonts", "hinted/ttf/NotoNaskhArabic/NotoNaskhArabic-Bold.ttf", "NotoNaskhArabic-Bold.ttf"),
    # Noto Sans Arabic - خط sans عربي للعناوين (OFL)
    ("notofonts/noto-fonts", "hinted/ttf/NotoSansArabic/NotoSansArabic-Regular.ttf", "NotoSansArabic-Regular.ttf"),
    ("notofonts/noto-fonts", "hinted/ttf/NotoSansArabic/NotoSansArabic-Bold.ttf", "NotoSansArabic-Bold.ttf"),
    # Noto Sans لاتيني للمصاحب الأجنبي داخل الكتاب (OFL)
    ("notofonts/noto-fonts", "hinted/ttf/NotoSans/NotoSans-Regular.ttf", "NotoSans-Regular.ttf"),
    ("notofonts/noto-fonts", "hinted/ttf/NotoSans/NotoSans-Bold.ttf", "NotoSans-Bold.ttf"),
    # وزنان إضافيان:
    #   • مائل لاتيني: للعناوين الأجنبية داخل المتن (*Learning and Instruction*)
    #   • عربي خفيف: بديل «المائل» في العربية (أمثلة النطق مثل *رسمتُ*) لأن
    #     العربية لا تُمال، والتمييز عندنا بوزن أخف لا بإزاحة الحروف.
    ("notofonts/noto-fonts", "hinted/ttf/NotoSans/NotoSans-Italic.ttf", "NotoSans-Italic.ttf"),
    ("notofonts/noto-fonts", "hinted/ttf/NotoSans/NotoSans-BoldItalic.ttf", "NotoSans-BoldItalic.ttf"),
    ("notofonts/noto-fonts", "hinted/ttf/NotoSansArabic/NotoSansArabic-Light.ttf", "NotoSansArabic-Light.ttf"),
    ("notofonts/noto-fonts", "hinted/ttf/NotoSansArabic/NotoSansArabic-Medium.ttf", "NotoSansArabic-Medium.ttf"),
]


def _get(url: str) -> bytes:
    req = urllib.request.Request(url, headers={"Accept": "application/vnd.github+json",
                                               "User-Agent": "book2-font-fetcher"})
    with urllib.request.urlopen(req, timeout=90) as r:
        return r.read()


def download(repo: str, path: str, out_name: str) -> bool:
    out = os.path.join(DEST, out_name)
    if os.path.exists(out) and os.path.getsize(out) > 10000:
        print(f"  ✓ موجود مسبقاً: {out_name}")
        return True
    try:
        meta = json.loads(_get(f"{API}/repos/{repo}/contents/{path}"))
        if isinstance(meta, dict) and meta.get("message"):
            print(f"  ✗ {out_name}: {meta['message']}")
            return False
        data = meta.get("content")
        if data and meta.get("encoding") == "base64":
            blob = base64.b64decode(data)
        else:  # ملف أكبر من 1MB -> واجهة blobs
            blob = base64.b64decode(json.loads(_get(f"{API}/repos/{repo}/git/blobs/{meta['sha']}"))["content"])
        with open(out, "wb") as f:
            f.write(blob)
        print(f"  ✓ {out_name}  ({len(blob)/1024:.0f} KB)")
        return True
    except Exception as exc:  # noqa: BLE001
        print(f"  ✗ {out_name}: {exc}")
        return False


def main() -> int:
    os.makedirs(DEST, exist_ok=True)
    print(f"وجهة الخطوط: {DEST}")
    ok = sum(download(*f) for f in FONTS)
    print(f"\nالنتيجة: {ok}/{len(FONTS)} ملف خط جاهز")
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())
