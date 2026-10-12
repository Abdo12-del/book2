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
_SEP = ""
_AR_INDICES = "٠١٢٣٤٥٦٧٨٩"
_EN_DIGITS = "0123456789"

_reshaper = arabic_reshaper.ArabicReshaper(
    configuration={
        "delete_harakat": False,      # الحفاظ على التشكيل (مهم للكتب العربية)
        "support_ligatures": True,    # دعم الربط مثل: لا، لله
        "shift_harakat_position": False,
        "use_unshaped_instead_of_isolated": False,
    }
)

_ARABIC_RE = re.compile(r"[؀-ۿݐ-ݿ]")


def has_arabic(text: str) -> bool:
    return bool(_ARABIC_RE.search(text))


def shape(text: str) -> str:
    """تحويل نص عربي منطقي إلى سلسلة بصرية جاهزة للرسم في PDF."""
    return get_display(_reshaper.reshape(text))


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
    visual = get_display(_reshaper.reshape(joined))
    return [(i, part) for i, part in enumerate(visual.split(_SEP)) if part != ""]


def to_arabic_indic(text: str) -> str:
    return "".join(_AR_INDICES[_EN_DIGITS.index(c)] if c in _EN_DIGITS else c for c in text)


def format_number(n: int, style: str = "arabic-indic") -> str:
    """تنسيق رقم الصفحة: أرقام عربية-هندية (١٢٣) أو أرقام غربية (123)."""
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
