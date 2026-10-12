#!/usr/bin/env python3
"""أدوات معالجة النص العربي: التشكيل السياقي، والترتيب البصري (Bidi)، والأرقام.

لماذا نحتاج هذا؟
    مكتبة ReportLab لا تُجري "تشكيلاً سياقياً" (Contextual Shaping) للحروف العربية
    ولا ترتيباً ثنائي الاتجاه، لذا يجب تحويل النص المنطقي إلى سلسلة بصرية جاهزة
    للرسم قبل تمريرها إلى المحرّك. أمّا ملفات DOCX فلا تحتاج ذلك؛ لأن Word يقوم
    بهذه المعالجة بنفسه، ونكتفي فيها بضبط اتجاه الفقرة (w:bidi).
"""
from __future__ import annotations

import re

import arabic_reshaper
from bidi.algorithm import get_display

# فاصل داخلي لترميز حدود "الأنماط المضمّنة" (عريض/مائل) قبل تمرير النص إلى خوارزمية Bidi.
#
# ⚠ يجب أن يكون هذا الفاصل **محايداً** في خوارزمية الاتجاه الثنائي (Bidi).
# كان سابقاً U+E000 من منطقة الاستخدام الخاص (PUA) وهو «قوي الاتجاه» LTR، فكان
# يفسد ترتيب ما حوله: أي فقرة تبدأ أو تنتهي بمقطع منسّق تظهر نقاطها وعلاماتها
# في الجهة المعاكسة. البديل U+2000 (EN QUAD) صنفه WS محايد، فلا يغيّر ترتيب
# المقاطع العربية، ومع ذلك يبقى صالحاً كفاصل موثوق (لا يرد في نصّ المخطوطة).
_SEP = "\u2000"
_AR_INDICES = "٠١٢٣٤٥٦٧٨٩"
_EXT_INDICES = "۰۱۲۳۴۵۶۷۸۹"     # أرقام فارسية/أوردية — نطبّعها أيضاً
_EN_DIGITS = "0123456789"
NUMERAL_STYLES = ("western", "arabic-indic")
_numeral_style = "western"        # النمط المعتمد في الكتاب: أرقام لاتينية 123

_reshaper = arabic_reshaper.ArabicReshaper(
    configuration={
        "delete_harakat": False,      # الحفاظ على التشكيل (مهم للكتب العربية)
        "support_ligatures": True,    # دعم الربط مثل: لا، لله
        "shift_harakat_position": False,
        "use_unshaped_instead_of_isolated": False,
    }
)

_ARABIC_RE = re.compile(r"[؀-ۿݐ-ݿ]")


# اتجاه الفقرة الأساسي لخوارزمية الاتجاه الثنائي (Bidi).
#
# خطأ شائع: ترك الخوارزمية تستنتج الاتجاه من أول محرف قوي في النص؛ فالسطر الذي يبدأ
# باسم لاتيني (كتاب أو اختصار: «Made to Stick»، «TED Talks») تُستنتج له فقرة باتجاه
# LTR، فينتقل المقطع اللاتيني إلى أقصى **اليسار** بدل أقصى اليمين، وينزلق ترقيم آخر
# السطر إلى الجهة المعاكسة — فيقرأ القارئ العربي الجملة مقلوبة الترتيب. الاتجاه هنا
# يتبع اتجاه المستند (rtl → R، ltr → L) ويُضبط من الثيمة عبر set_base_dir.
BASE_DIR = "R"


def set_base_dir(direction: str) -> None:
    """يضبط اتجاه الفقرة من اتجاه المستند: ``ltr`` → L، وما عداه → R."""
    global BASE_DIR
    BASE_DIR = "L" if (direction or "").strip().lower() in ("ltr", "left") else "R"


def has_arabic(text: str) -> bool:
    return bool(_ARABIC_RE.search(text))


def shape(text: str) -> str:
    """تحويل نص عربي منطقي إلى سلسلة بصرية جاهزة للرسم في PDF."""
    return get_display(_reshaper.reshape(text), base_dir=BASE_DIR)


def shape_runs(segments: list[str]) -> list[tuple[int, str]]:
    """تشكيل عدّة مقاطع مضمّنة (عريض/مائل/عادي) معاً.

    تُشكَّل المقاطع مجتمعةً حتى لا تنكسر وصلات الحروف عند حدود التنسيق، ثم تُعاد
    بالترتيب البصري الصحيح مع فهرس المقطع الأصلي.

    المُدخل: ["نص ", "عريض", " ونص آخر"]
    الخرج:   [(0, "…"), (1, "…"), (0, "…")] مرتّبة بصرياً
    """
    if len(segments) == 1:
        return [(0, shape(segments[0]))]
    joined = _SEP.join(segments)
    visual = get_display(_reshaper.reshape(joined), base_dir=BASE_DIR)
    return [(i, part) for i, part in enumerate(visual.split(_SEP)) if part != ""]


def to_arabic_indic(text: str) -> str:
    return "".join(_AR_INDICES[_EN_DIGITS.index(c)] if c in _EN_DIGITS else c for c in text)


def to_western(text: str) -> str:
    """توحيد كل صور الأرقام إلى الأرقام اللاتينية (123).

    يحوّل العربية-الهندية (٠١٢) والفارسية (۰۱۲) إلى نظيرها اللاتيني.
    """
    out = []
    for c in text:
        if c in _AR_INDICES:
            out.append(_EN_DIGITS[_AR_INDICES.index(c)])
        elif c in _EXT_INDICES:
            out.append(_EN_DIGITS[_EXT_INDICES.index(c)])
        else:
            out.append(c)
    return "".join(out)


def normalize_numerals(text: str, style: str | None = None) -> str:
    """يطبّع الأرقام في النص وفق نمط الكتاب المعتمد.

    النمط الافتراضي «western» (123) لأن الأرقام العربية-الهندية (١٢٣) تُقرأ
    ببطء لدى شريحة واسعة من القرّاء وتُلتبس في التواريخ والنسب. النمط الآخر
    «arabic-indic» متاح لمن يريده: يُمرَّر عبر ``book.yaml → numerals``.
    """
    style = style or _numeral_style
    if style == "arabic-indic":
        return to_arabic_indic(text)
    return to_western(text)


def set_numeral_style(style: str) -> None:
    """يضبط نمط الأرقام العام (western | arabic-indic)."""
    global _numeral_style
    if style not in NUMERAL_STYLES:
        raise ValueError(f"نمط أرقام غير معروف: {style!r} (المتاح: {NUMERAL_STYLES})")
    _numeral_style = style


def numeral_style() -> str:
    return _numeral_style


def has_arabic_indic(text: str) -> list[str]:
    """يعيد الأرقام العربية-الهندية/الفارسية الموجودة في النص (لفحص الجودة)."""
    return [c for c in text if c in _AR_INDICES or c in _EXT_INDICES]


# علامات تُنهي العبارة؛ لا يجوز أن تبدأ بها سطرٌ عربي (مؤشر خطأ في تقسيم الأسطر).
CLOSING_PUNCT = "،؛:.!؟)%»]}،"


def starts_with_closing_punct(line: str) -> bool:
    s = (line or "").lstrip()
    return bool(s) and s[0] in CLOSING_PUNCT


def format_number(n: int, style: str | None = None) -> str:
    """تنسيق رقم: أرقام لاتينية (123) أو عربية-هندية (١٢٣) أو رومانية."""
    style = style or _numeral_style
    if style == "arabic-indic":
        return to_arabic_indic(str(n))
    if style == "roman":
        return _to_roman(n).lower()
    return str(n)


def _to_roman(n: int) -> str:
    table = [(1000, "M"), (900, "CM"), (500, "D"), (400, "CD"), (100, "C"), (90, "XC"),
             (50, "L"), (40, "XL"), (10, "X"), (9, "IX"), (5, "V"), (4, "IV"), (1, "I")]
    out = []
    for value, sym in table:
        while n >= value:
            out.append(sym)
            n -= value
    return "".join(out)


def strip_diacritics(text: str) -> str:
    """إزالة التشكيل (مفيد لعناوين الفهرس أو للنصوص المنسوخة من مصادر مشكّلة)."""
    return "".join(c for c in text if not (0x064B <= ord(c) <= 0x065F or ord(c) in (0x0670,)))
