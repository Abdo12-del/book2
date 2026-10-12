#!/usr/bin/env python3
"""مُصدّر PDF: يبني كتاباً عربياً كاملاً من شجرة الكُتل باستخدام ReportLab.

المزايا:
  • صفحات متقابلة (هوامش مرآتية مع هامش تجليد) واتجاه RTL كامل.
  • ترويسة/تذييل معاينة، وترقيم صفحات بالأرقام العربية-الهندية أو الغربية.
  • فهرس محتويات تلقائي بعدّة تمريرات (Two-pass TOC).
  • عناوين لا تنفصل عن فقراتها (KeepWithNext) وبداية فصل على صفحة فردية.
  • جداول وأشكال وصناديق ملاحظات وتعليقات مُرقّمة تلقائياً.
  • تضمين الخطوط داخل الملف (Subset Embedding).
"""
from __future__ import annotations

import os
import xml.sax.saxutils as saxutils
from typing import Any

from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
from reportlab.lib.pagesizes import A4, A5, B5
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import (BaseDocTemplate, CondPageBreak, Flowable, Frame,
                                FrameBreak, Image, KeepTogether, NextPageTemplate,
                                PageBreak, PageTemplate, Paragraph, Spacer, Table,
                                TableStyle)

import arabic as A
import textflow as TF
from manuscript import Block
from theme import MM, hex_to_rgb, load as load_theme

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
FONT_DIR = os.path.join(ROOT, "assets", "fonts")

ORDINALS = ["الأول", "الثاني", "الثالت", "الرابع", "الخامس", "السادس", "السابع", "الثامن",
            "التاسع", "العاشر", "الحادي عشر", "الثاني عشر", "الثالث عشر", "الرابع عشر",
            "الخامس عشر", "السادس عشر", "السابع عشر", "الثامن عشر", "التاسع عشر", "العشرون",
            "الحادي والعشرون", "الثاني والعشرون", "الثالث والعشرون", "الرابع والعشرون",
            "الخامس والعشرون", "السادس والعشرون", "السابع والعشرون", "الثامن والعشرون",
            "التاسع والعشرون", "التلثون"]
ORDINALS[2] = "الثالث"
ORDINALS[29] = "الثلاثون"


def chapter_word(n: int) -> str:
    return ORDINALS[n - 1] if 1 <= n <= len(ORDINALS) else str(n)


# --------------------------------------------------------------------------- الخطوط
# اسم العائلة كما يُكتب في Word لكل خطّ مسجَّل (يختلف عن اسم الملف: «Noto Sans Arabic»)
FAMILY: dict[str, str] = {}


def family_name(path: str) -> str:
    """اسم عائلة الخطّ من جدول الأسماء داخل ملف TTF."""
    try:
        from reportlab.pdfbase.ttfonts import TTFontFile
        nm = TTFontFile(path).familyName
        return nm.decode("utf-8", "ignore") if isinstance(nm, bytes) else str(nm)
    except Exception:  # noqa: BLE001
        return os.path.basename(path).rsplit(".", 1)[0]


def register_fonts(theme: dict) -> dict[str, str]:
    """تسجيل الخطوط المتاحة؛ يُعيد خريطة الاسم المنطقي → الاسم المسجّل."""
    avail = {
        "Amiri": "Amiri-Regular.ttf", "Amiri-Bold": "Amiri-Bold.ttf",
        "Amiri-Italic": "Amiri-Italic.ttf", "Amiri-BoldItalic": "Amiri-BoldItalic.ttf",
        # Noto Naskh / Noto Sans Arabic: تغطية كاملة للأشكال السياقية (ممتازة للعناوين)
        "NotoNaskhArabic": "NotoNaskhArabic-Regular.ttf",
        "NotoNaskhArabic-Bold": "NotoNaskhArabic-Bold.ttf",
        "NotoSansArabic": "NotoSansArabic-Regular.ttf",
        "NotoSansArabic-Bold": "NotoSansArabic-Bold.ttf",
        "NotoSans": "NotoSans-Regular.ttf", "NotoSans-Bold": "NotoSans-Bold.ttf",
        "NotoSans-Italic": "NotoSans-Italic.ttf",
        "NotoSans-BoldItalic": "NotoSans-BoldItalic.ttf",
        "NotoSansArabic-Light": "NotoSansArabic-Light.ttf",
        "NotoSansArabic-Medium": "NotoSansArabic-Medium.ttf",
        "Tajawal": "Tajawal-Regular.ttf", "Tajawal-Bold": "Tajawal-Bold.ttf",
    }
    system = {
        "DejaVuSans": "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
        "DejaVuSans-Bold": "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",
        "DejaVuSerif": "/usr/share/fonts/truetype/dejavu/DejaVuSerif.ttf",
        "DejaVuSansMono": "/usr/share/fonts/truetype/dejavu/DejaVuSansMono.ttf",
    }
    registered: dict[str, str] = {}
    for name, fn in avail.items():
        p = os.path.join(FONT_DIR, fn)
        if os.path.exists(p):
            try:
                pdfmetrics.registerFont(TTFont(name, p))
                registered[name] = name
                FAMILY[name] = family_name(p)
            except Exception:  # noqa: BLE001
                pass
    for name, p in system.items():
        if os.path.exists(p):
            try:
                pdfmetrics.registerFont(TTFont(name, p))
                registered[name] = name
                FAMILY[name] = family_name(p)
            except Exception:  # noqa: BLE001
                pass
    for fam in ("NotoSansArabic", "NotoSans", "Amiri", "DejaVuSans"):
        if all(f"{fam}{s}" in registered for s in ("", "-Bold")):
            pdfmetrics.registerFontFamily(
                fam, normal=fam, bold=f"{fam}-Bold",
                italic=registered.get(f"{fam}-Italic", fam),
                boldItalic=registered.get(f"{fam}-BoldItalic", f"{fam}-Bold"))
    # بدائل آمنة إن غاب خط
    def res(want: str, fallback: str = "NotoSansArabic") -> str:
        return want if want in registered else (fallback if fallback in registered else
                                                next(iter(registered), "Helvetica"))

    f = theme["fonts"]
    lat = res(f.get("latin", "NotoSans"), f["body"])
    return {"body": res(f["body"]), "body_bold": res(f["body_bold"], f["body"]),
            "body_italic": res(f["body_italic"], f["body"]),
            "heading": res(f["heading"], f["body"]), "heading_bold": res(f["heading_bold"], f["body"]),
            "mono": res(f["mono"], "DejaVuSansMono"), "latin": lat,
            "latin_bold": res(f.get("latin_bold", "NotoSans-Bold"), lat),
            "latin_italic": res(f.get("latin_italic", "NotoSans-Italic"), lat)}


# --------------------------------------------------------------------------- الأنماط
def build_styles(theme: dict, fonts: dict[str, str]) -> dict[str, ParagraphStyle]:
    s, c, f = theme["sizes"], theme["colors"], fonts
    rtl = theme["direction"] == "rtl"
    align = TA_RIGHT if rtl else TA_LEFT
    ink = colors.Color(*hex_to_rgb(c["text"]))
    head = colors.Color(*hex_to_rgb(c["heading"]))
    s_body = s["body_pt"]
    ratio = s["body_leading_pt"] / s["body_pt"]

    def mk(name, **kw) -> ParagraphStyle:
        kw.setdefault("fontName", fonts["body"])
        kw.setdefault("fontSize", s_body)
        kw.setdefault("leading", s_body * ratio)
        kw.setdefault("alignment", align)
        kw.setdefault("textColor", ink)
        # wordWrap=None مقصود: نُشكّل النصّ بأنفسنا في tools/textflow.py ونمرّره
        # بترتيبه المرئي. لو تركناه على «RTL» لأعاد ReportLab ترتيب المقاطع
        # بنفسه على نصٍّ مرئي أصلاً، فيُنتج تشويشاً (وقد جُرِّب وأثبت الفرق).
        kw.setdefault("wordWrap", None)
        return ParagraphStyle(name, **kw)

    return {
        # allowWidows=0: منع السطر اليتيم في أعلى الصفحة (ReportLab يسمح به افتراضياً)
        "body": mk("body", spaceAfter=s["body_leading_pt"] * 0.35, allowWidows=0, allowOrphans=0),
        "body_first": mk("body_first", spaceAfter=s["body_leading_pt"] * 0.35,
                         allowWidows=0, allowOrphans=0),
        "h1": mk("h1", fontName=fonts["heading"], fontSize=s["h1_pt"], leading=s["h1_pt"] * 1.5,
                 textColor=head, alignment=TA_CENTER, spaceBefore=0, spaceAfter=10),
        "h1_num": mk("h1_num", fontName=fonts["heading"], fontSize=s["h1_pt"] * 0.62,
                     leading=s["h1_pt"], textColor=colors.Color(*hex_to_rgb(c["accent"])),
                     alignment=TA_CENTER, spaceAfter=8),
        "h2": mk("h2", fontName=fonts["heading_bold"], fontSize=s["h2_pt"],
                 leading=s["h2_pt"] * 1.45, textColor=head,
                 spaceBefore=theme["layout"]["space_before_h2_pt"],
                 spaceAfter=theme["layout"]["space_after_h2_pt"],
                 keepWithNext=1 if theme["layout"]["keep_headings_with_next"] else 0),
        "h3": mk("h3", fontName=fonts["heading"], fontSize=s["h3_pt"], leading=s["h3_pt"] * 1.4,
                 textColor=colors.Color(*hex_to_rgb(c["heading"])),
                 spaceBefore=theme["layout"]["space_before_h3_pt"],
                 spaceAfter=theme["layout"]["space_after_h3_pt"],
                 keepWithNext=1 if theme["layout"]["keep_headings_with_next"] else 0),
        "h4": mk("h4", fontName=fonts["body_bold"], fontSize=s["h4_pt"], leading=s["h4_pt"] * 1.4,
                 spaceBefore=10, spaceAfter=4, keepWithNext=1),
        "caption": mk("caption", fontSize=s["caption_pt"], leading=s["caption_pt"] * 1.5,
                      textColor=colors.Color(*hex_to_rgb(c["muted"])), alignment=TA_CENTER,
                      spaceBefore=6, spaceAfter=12),
        "toc0": mk("toc0", fontName=fonts["heading"], fontSize=s["toc_pt"],
                   leading=s["toc_pt"] * 1.85, textColor=ink),
        "toc1": mk("toc1", fontSize=s["toc_pt"] * 0.95, leading=s["toc_pt"] * 1.7),
        "quote": mk("quote", fontName=fonts["body_italic"], fontSize=s["body_pt"] * 0.95,
                    leading=s["body_pt"] * ratio * 1.05, alignment=TA_CENTER,
                    textColor=colors.Color(*hex_to_rgb(c["muted"])), spaceBefore=8, spaceAfter=14),
        "bullet": mk("bullet", spaceAfter=s["body_leading_pt"] * 0.18,
                     allowWidows=0, allowOrphans=0),
        "code": mk("code", fontName=fonts["mono"], fontSize=s["body_pt"] * 0.8,
                   leading=s["body_pt"] * 1.15, alignment=TA_LEFT, textColor=ink),
        "cell": mk("cell", fontSize=s["table_pt"], leading=s["table_pt"] * 1.5,
                   alignment=TA_CENTER, spaceAfter=0),
        "cell_head": mk("cell_head", fontName=fonts["heading"], fontSize=s["table_pt"] * 0.95,
                        leading=s["table_pt"] * 1.5, alignment=TA_CENTER,
                        textColor=colors.Color(*hex_to_rgb(c["heading"])), spaceAfter=0),
        "callout": mk("callout", fontSize=s["body_pt"] * 0.96,
                      leading=s["body_pt"] * ratio * 0.98, spaceAfter=0),
        "callout_title": mk("callout_title", fontName=fonts["heading"],
                            fontSize=s["body_pt"] * 0.95, leading=s["body_pt"] * 1.5,
                            textColor=colors.Color(*hex_to_rgb(c["accent"])), spaceAfter=4),
        "cover_title": mk("cover_title", fontName=fonts["heading"], fontSize=28, leading=42,
                          alignment=TA_CENTER, textColor=colors.Color(*hex_to_rgb(c["heading"]))),
        "cover_subtitle": mk("cover_subtitle", fontSize=15, leading=26, alignment=TA_CENTER,
                             textColor=colors.Color(*hex_to_rgb(c["muted"]))),
        "cover_author": mk("cover_author", fontName=fonts["heading"], fontSize=14, leading=24,
                           alignment=TA_CENTER, textColor=colors.Color(*hex_to_rgb(c["accent"]))),
        "front_h": mk("front_h", fontName=fonts["heading"], fontSize=20, leading=32,
                      alignment=TA_CENTER, textColor=colors.Color(*hex_to_rgb(c["heading"])),
                      spaceAfter=18),
    }


# --------------------------------------------------------------------------- أدوات مساعدة
def esc(t: str) -> str:
    return saxutils.escape(t)


def V(text: str) -> str:
    """نصّ جاهز للرسم: تطبيع الأرقام ثم التشكيل السياقي وترتيب الاتجاه.

    تُستخدم للعناوين والتعليقات وسطور الجداول — أي نصّ **لا يُكسَّر على عدة
    أسطر** بآلية ReportLab. الفقرات الطويلة لها مسار آخر (انظر `para_runs`)
    لأن تقسيم الأسطر يجب أن يجري على النص المنطقي قبل التشكيل.
    """
    return TF.plain_visual(A.normalize_numerals(text))


def para_runs(text: str, style: ParagraphStyle, fonts: dict[str, str],
              width: float) -> Paragraph:
    """فقرة عربية بتنسيق مضمّن (**عريض** / *مائل* / `كود`) وتقسيم أسطر صحيح.

    التقسيم يجري على النص المنطقي (قبل التشكيل)، ثم يُشكَّل كل سطر على حدة،
    فيستقيم ترتيب الأسطر وتستقر مواضع الترقيم داخل السطر الصحيح.
    التفصيل في tools/textflow.py.
    """
    from manuscript import _inline

    def font_of(kind: str, chunk: str = "") -> str:
        """خطّ المجموعة: العربية بخطّها، واللاتينية والأرقام والترقيم بخطّها.

        خطّ العربية (Noto Sans Arabic) لا يحتوي حروفاً لاتينية ولا ترقيمًا
        (— … « »)، وخطّ اللاتينية لا يحتوي الأشكال العربية — فالاختيار بالمحتوى.
        """
        if TF.is_arabic_run(chunk):
            return {"b": fonts["body_bold"], "i": fonts["body_italic"],
                    "c": fonts["mono"]}.get(kind or "", fonts["body"])
        return {"b": fonts["latin_bold"], "i": fonts["latin_italic"],
                "c": fonts["mono"]}.get(kind or "", fonts["latin"])

    segs = _inline(A.normalize_numerals(text))
    html = TF.paragraph_html(segs, width, style.fontSize, font_of,
                              style=style)
    return Paragraph(html or " ", style)


def mark(flowable, **kw):
    for k, v in kw.items():
        setattr(flowable, k, v)
    return flowable


class SectionSwitch(Flowable):
    """يبدّل نمط الصفحة (غلاف / تمهيد / متن) ابتداءً من الصفحة التالية."""

    _ZEROSIZE = True

    def __init__(self, section: str):
        super().__init__()
        self.section = section
        self.height = 0
        self.width = 0

    def wrap(self, *a):
        doc = self._doctemplate
        if doc is not None:
            doc.set_section(self.section)
        return 0, 0

    def draw(self):
        pass


def _start_page(page: int, parity: str) -> int:
    """أول صفحة لا تقلّ عن ``page`` وتحقّق الجهة المطلوبة (right = فردية/وجه)."""
    if parity not in ("right", "left"):
        return page
    want_odd = parity != "left"
    if (page % 2 == 1) != want_odd:
        return page + 1
    return page


class EnsureOddPage(Flowable):
    """يفرض بدء كتلة عنوان الفصل على صفحة الوجه (فردية عادةً) بكامل مساحتها.

    الشرطان معاً ضروريان:
      * **الجهة:** إن كانت الصفحة الحالية على غير الجهة المطلوبة نتخطّاها (فتصير
        صفحة فاصلة إن لم يكن فيها نصّ سابق).
      * **المساحة:** إن لم تتّسع الصفحة لكتلة العنوان (``need`` نقطة) ننتقل، فلا
        ينشقّ عنوان الفصل بين صفحتين.

    لا نُسجّل هنا «صفحات فاصلة»: الصفحة التي نتخطّاها قد تحمل نصّاً من الفصل
    السابق، فليست فاصلة. تحديد الصفحات الفاصلة يجري بقراءة الملف الناتج نفسه
    (``frame_blank_pages``) ثم يُمرَّر إلى مرور البناء التالي لتُخفى ترويسته.
    """

    _ZEROSIZE = True

    def __init__(self, parity: str = "right", need: float = 0.0):
        super().__init__()
        self.parity = parity        # right = فردية | left = زوجية | any = بلا قيد
        self.need = need            # أدنى ارتفاع مطلوب لكتلة عنوان الفصل

    def wrap(self, availWidth, availHeight):
        doc = self._doctemplate
        if doc is None:
            return 0, 0
        start = doc.page + (1 if availHeight < self.need else 0)
        target = _start_page(start, self.parity)
        for _ in range(target - doc.page):
            doc.frame.add_generated_content(FrameBreak)
        return 0, 0

    def draw(self):
        pass


class Rule(Flowable):
    """خط زخرفي أفقي."""

    def __init__(self, width: float, thickness: float = 0.7, color=None, space: float = 6):
        super().__init__()
        self.width, self.thickness, self.color = width, thickness, color or colors.grey
        self.space = space
        self.height = thickness + space * 2

    def wrap(self, *a):
        return self.width, self.height

    def draw(self):
        self.canv.setStrokeColor(self.color)
        self.canv.setLineWidth(self.thickness)
        self.canv.line(0, self.space, self.width, self.space)


class OrnamentalBreak(Flowable):
    """فاصل زخرفي مركزي (◆ ─── ◆)."""

    def __init__(self, width: float, color=None):
        super().__init__()
        self.width, self.color = width, color or colors.Color(0.55, 0.43, 0.25)
        self.height = 22

    def wrap(self, *a):
        return self.width, self.height

    def draw(self):
        """رسم متجهي للمعيّن (لا يعتمد على وجود الرمز في الخط)."""
        c = self.canv
        mid, y, r = self.width / 2, 9.0, 3.4
        c.setStrokeColor(self.color)
        c.setLineWidth(0.5)
        c.line(0, y, mid - 20, y)
        c.line(mid + 20, y, self.width, y)
        c.setFillColor(self.color)
        pth = c.beginPath()
        pth.moveTo(mid, y + r)
        pth.lineTo(mid + r, y)
        pth.lineTo(mid, y - r)
        pth.lineTo(mid - r, y)
        pth.close()
        c.drawPath(pth, stroke=0, fill=1)


class TOCFlowable(Flowable):
    """فهرس المحتويات (يُبنى من الإدخالات المجموعة في التمريرة السابقة).

    يدعم التقسيم على أكثر من صفحة عبر ``split``: الفهرس الطويل (باب + فصول + عناصر)
    يتجاوز ارتفاع الصفحة الواحدة حتمًا، وبدون ``split`` يرفضه ReportLab بخطأ LayoutError.
    """

    def __init__(self, entries: list[tuple[int, str, int]], theme, fonts, width: float):
        super().__init__()
        self.entries, self.theme, self.fonts, self.width = entries, theme, fonts, width
        self.height = 0
        self._wrapped: list[tuple] = []

    def wrap(self, availWidth, availHeight):
        self._wrapped = []
        y = 0.0
        for level, text, page in self.entries:
            st = "toc0" if level == 1 else "toc1"
            p = para_runs(text, self.theme["_styles"][st], self.fonts, self.width - 46)
            p.wrap(self.width - 46, availHeight)
            h = p.height
            y += h + (3 if level == 1 else 0)
            self._wrapped.append((level, text, p, page, h))
        self.height = y
        return self.width, self.height

    def split(self, availWidth, availHeight):
        """يقسم الفهرس عند آخر مدخل يتسع له الإطار الحالي."""
        if not self._wrapped:
            self.wrap(availWidth, availHeight)
        if not self._wrapped or self.height <= availHeight:
            return []
        used, k = 0.0, 0
        for level, text, p, page, h in self._wrapped:
            step = h + (3 if level == 1 else 0)
            # لا نُدخل مدخلاً لا يتسع: إن لم يتسع ولا واحد نعيد [] فيُنهي ReportLab
            # الإطار الحالي وينتقل إلى صفحة جديدة (بدل «Splitting error»).
            if used + step > availHeight:
                break
            used += step
            k += 1
        if not k or k >= len(self._wrapped):
            return []
        return [TOCFlowable(self.entries[:k], self.theme, self.fonts, self.width),
                TOCFlowable(self.entries[k:], self.theme, self.fonts, self.width)]

    def draw(self):
        c = self.canv
        fm = A.format_number
        style = self.theme["running"]["page_number_style"]
        accent = colors.Color(*hex_to_rgb(self.theme["colors"]["accent"]))
        text_col = colors.Color(*hex_to_rgb(self.theme["colors"]["text"]))
        frame_w = self.width - 46          # إطار النص يبدأ من x=46
        y = self.height
        for level, text, p, page, h in self._wrapped:
            y -= h
            # النص على اليمين ورقم الصفحة على اليسار (اتجاه عربي)
            p.drawOn(c, 46, y)
            text_w = pdfmetrics.stringWidth(V(text), p.style.fontName, p.style.fontSize)
            if self.theme["toc"]["dotted_leader"]:
                # النص مُحاذاة لليمين داخل إطاره: يمتد من (46+frame_w-text_w) إلى (46+frame_w)
                x_right = 46 + frame_w - text_w - 6
                x_left = 26
                if x_right > x_left:
                    c.setStrokeColor(accent)
                    c.setLineWidth(0.4)
                    c.setDash(0.5, 2.2)
                    c.line(x_left, y + h * 0.35, x_right, y + h * 0.35)
                    c.setDash()
            c.setFillColor(text_col)
            c.setFont(self.fonts["latin"], self.theme["sizes"]["toc_pt"] * 0.9)
            c.drawString(6, y + h * 0.22, fm(page, style))
            y -= 4 if level == 1 else 0


# --------------------------------------------------------------------------- نموذج المستند
class BookDoc(BaseDocTemplate):
    def __init__(self, out_path, theme, fonts, meta, styles, toc_entries,
                 blank_pages: set[int] | None = None):
        p = theme["page"]
        self.theme, self.fonts, self.meta, self.styles = theme, fonts, meta, styles
        self.toc_entries: list[tuple[int, str, int]] = toc_entries
        self._section = "cover"
        self.body_start_page: int | None = None
        self.front_start_page: int | None = None
        self.blank_pages: set[int] = set(blank_pages or ())   # صفحات فاصلة (من المرور السابق)
        self.chapter_title = meta.get("title", "")
        self.text_width = p["width_pt"] - (p["margin_outer_mm"] * 2 + p["gutter_mm"]) * MM

        BaseDocTemplate.__init__(self, out_path, pagesize=(p["width_pt"], p["height_pt"]),
                                 title=meta.get("title", ""), author=meta.get("author", ""),
                                 subject=meta.get("subtitle", ""),
                                 leftMargin=0, rightMargin=0, topMargin=0, bottomMargin=0,
                                 allowSplitting=1)
        self._templates = self._make_templates()
        self._tmap = dict(self._templates)
        self._tmap["cover"] = self._templates["cover_odd"]
        self.addPageTemplates(list(dict.fromkeys(self._templates.values())))
        self.pageTemplate = self._tmap["cover"]

    # ---- بناء القوالب
    def _frame(self, section: str, odd: bool) -> Frame:
        p = self.theme["page"]
        W, H = p["width_pt"], p["height_pt"]
        top, bottom = p["margin_top_mm"] * MM, p["margin_bottom_mm"] * MM
        outer, inner = p["margin_outer_mm"] * MM, (p["margin_outer_mm"] + p["gutter_mm"]) * MM
        rtl = self.theme["direction"] == "rtl"
        if section == "cover":
            m = 18 * MM
            return Frame(m, m, W - 2 * m, H - 2 * m, id="cover", leftPadding=0, rightPadding=0,
                         topPadding=0, bottomPadding=0, showBoundary=0)
        # في الكتب العربية: الصفحة الفردية (الوجه) يكون التجليد على يمينها
        # إذا كان التجليد على اليمين فالهامش الأيسر هو الخارجي والعكس صحيح
        gutter_right = odd if rtl else (not odd)
        x1 = outer if gutter_right else inner
        w = W - outer - inner
        return Frame(x1, bottom, w, H - top - bottom, id=f"{section}_{'odd' if odd else 'even'}",
                     leftPadding=0, rightPadding=0, topPadding=0, bottomPadding=0, showBoundary=0)

    def _make_templates(self) -> dict[str, PageTemplate]:
        out: dict[str, PageTemplate] = {}
        for section in ("cover", "front", "body"):
            for odd in (True, False):
                name = f"{section}_{'odd' if odd else 'even'}"
                out[name] = PageTemplate(id=name, frames=[self._frame(section, odd)],
                                         onPage=self.decorate, autoNextPageTemplate=None)
        return out

    # ---- تبديل الأقسام
    def set_section(self, section: str) -> None:
        if section != self._section:
            self._section = section

    def handle_pageBegin(self):
        nxt = self.page + 1
        name = f"{self._section}_{'odd' if nxt % 2 else 'even'}"
        t = self._tmap.get(name)
        if t is not None:
            self.pageTemplate = t
            if self._section == "body" and self.body_start_page is None:
                # نبدأ الترقيم من أول صفحة فيها متن: الصفحة الفاصلة التي تسبق
                # افتتاحية الفصل لا تحمل رقماً، ولو بدأنا العدّ منها لانقلبت
                # جهة الأرقام (صفحة الوجه تحمل رقماً زوجياً).
                p0 = nxt
                while p0 in self.blank_pages:
                    p0 += 1
                self.body_start_page = p0
            if self._section == "front" and self.front_start_page is None:
                self.front_start_page = nxt
        BaseDocTemplate.handle_pageBegin(self)

    # ---- الترويسة والتذييل
    def decorate(self, canv, doc):
        th = self.theme
        p = th["page"]
        W, H = p["width_pt"], p["height_pt"]
        page = canv.getPageNumber()
        rtl = th["direction"] == "rtl"
        outer, inner = p["margin_outer_mm"] * MM, (p["margin_outer_mm"] + p["gutter_mm"]) * MM
        odd = page % 2 == 1
        gutter_right = odd if rtl else (not odd)
        x1 = outer if gutter_right else inner            # الحد الأيسر لعمود النص
        x2 = W - (inner if gutter_right else outer)      # الحد الأيمن لعمود النص

        if self._section == "cover" or page in self.blank_pages:
            return

        muted = colors.Color(*hex_to_rgb(th["colors"]["muted"]))
        rule_c = colors.Color(*hex_to_rgb(th["colors"]["rule"]))
        head_c = colors.Color(*hex_to_rgb(th["colors"]["heading"]))

        # الترويسة: عنوان الفصل (فردي) / عنوان الكتاب (زوجي) بجهة التجليد
        cfg = th["running"]
        key = cfg.get("odd_header") if odd else cfg.get("even_header")
        if key and key != "none":
            label = self.chapter_title if key == "chapter" else self.meta.get("title", "")
            if label:
                txt = V(label)
                canv.setFillColor(muted)
                y = H - p["header_from_top_mm"] * MM
                self._draw_runs(canv, txt, x2 if gutter_right else x1, y,
                                th["sizes"]["header_pt"], right=gutter_right,
                                font=self.fonts["body"])
                if cfg.get("header_rule"):
                    canv.setStrokeColor(rule_c)
                    canv.setLineWidth(0.5)
                    canv.line(x1, y - 5, x2, y - 5)

        # التذييل: رقم الصفحة في منتصف المسافة
        if cfg.get("footer") == "page_number":
            if self._section == "front":
                n = page - (self.front_start_page or 1) + 1
                num_style = cfg.get("front_matter_numbering", "arabic-indic")
            else:
                n = page - (self.body_start_page or 1) + 1
                num_style = cfg.get("page_number_style", "arabic-indic")
            txt = A.format_number(n, num_style)
            canv.setFont(self.fonts["latin"], th["sizes"]["footer_pt"])
            canv.setFillColor(head_c)
            canv.drawCentredString((x1 + x2) / 2, p["footer_from_bottom_mm"] * MM, txt)

    def _draw_runs(self, canv, text: str, x: float, y: float, size: float,
                   right: bool = True, color=None, font=None) -> None:
        """يرسم سطراً واحداً بمجاميع خطوطه (عربي/لاتيني) من نقطة محاذاة واحدة."""
        ar_font = font or self.fonts["heading"]
        lat_font = self.fonts["latin"]
        runs = TF.script_runs(text)
        widths = [pdfmetrics.stringWidth(t, ar_font if a else lat_font, size)
                  for a, t in runs]
        if color is not None:
            canv.setFillColor(color)
        if right:
            x -= sum(widths)
        for (a, t), w in zip(runs, widths):
            canv.setFont(ar_font if a else lat_font, size)
            canv.drawString(x, y, t)
            x += w

    # ---- جمع مدخلات الفهرس
    def afterFlowable(self, flowable):
        lvl = getattr(flowable, "_book_level", None)
        if lvl is not None and self._section == "body":
            txt = getattr(flowable, "_book_text", "")
            if lvl <= self.theme["toc"]["depth"]:
                entry = (lvl, txt, self.page - (self.body_start_page or 1) + 1)
                if not self.toc_entries or self.toc_entries[-1][:2] != entry[:2]:
                    self.toc_entries.append(entry)
        if lvl == 1:
            self.chapter_title = getattr(flowable, "_book_text", "")


# --------------------------------------------------------------------------- بناء القصة
class StoryBuilder:
    def __init__(self, theme, fonts, styles, meta, toc_entries):
        self.th, self.f, self.st, self.meta = theme, fonts, styles, meta
        self.toc_entries = toc_entries
        self.chapter = 0
        self.table_n = 0
        self.fig_n = 0
        self.text_width = (theme["page"]["width_pt"]
                           - (theme["page"]["margin_outer_mm"] * 2 + theme["page"]["gutter_mm"]) * MM)
        self.accent = colors.Color(*hex_to_rgb(theme["colors"]["accent"]))
        self.rule_c = colors.Color(*hex_to_rgb(theme["colors"]["rule"]))

    def heading(self, b: Block) -> list[Flowable]:
        out: list[Flowable] = []
        label = self.meta.get("chapter_label", "الفصل")
        if b.level == 1:
            self.chapter += 1
            self.table_n = 0
            self.fig_n = 0
            tail: list[Flowable] = []
            if self.chapter > 1 or True:
                tail.append(Spacer(1, self.th["layout"]["chapter_head_space_before_pt"]))
            if not self.th["layout"].get("drop_chapter_number"):
                num = f"{label} {chapter_word(self.chapter)}"
                tail.append(mark(Paragraph(V(num), self.st["h1_num"]), _book_level=None))
            tail.append(Rule(self.text_width * 0.32, 1.1, self.accent, space=4))
            p = mark(para_runs(b.text, self.st["h1"], self.f, self.text_width),
                     _book_level=1, _book_text=b.text)
            tail.append(p)
            tail.append(Rule(self.text_width * 0.32, 1.1, self.accent, space=4))
            # ارتفاع كتلة العنوان: يمنع انشقاقها، ويُستخدم في فحص الجهة
            need = sum(f.wrap(self.text_width, 10 ** 6)[1] for f in tail)
            parity = self.th["layout"].get("chapter_starts_on", "right")
            out += [SectionSwitch("body"), PageBreak(), EnsureOddPage(parity, need)]
            out += tail
            return out
        style = self.st.get(f"h{min(b.level, 4)}")
        p = mark(para_runs(b.text, style, self.f, self.text_width),
                 _book_level=b.level, _book_text=b.text)
        if self.th["layout"]["keep_headings_with_next"]:
            return [CondPageBreak(90), p]
        return [p]

    def paragraph(self, b: Block) -> list[Flowable]:
        return [para_runs(b.text, self.st["body"], self.f,
                          self.text_width)]

    def bullet_list(self, b: Block, width: float | None = None) -> list[Flowable]:
        """قوائم بمحاذاة عربية دقيقة: عمود العلامة على اليمين، والنص في العمود الآخر.

        استُخدم جدولٌ بعمودين بدل آلية bulletText في ReportLab؛ لأن موضع العلامة
        فيها غير منضبط مع الاتجاه من اليمين إلى اليسار، وقد يقع خارج الهامش.
        """
        ordered = b.data.get("ordered")
        rtl_dir = self.th["direction"] == "rtl"
        counter = 1
        rows = []
        marker_w = 26.0
        avail = width or self.text_width
        for it in b.data["items"]:
            lvl = min(it["level"], 1)
            if ordered:
                # موضع النقطة مهمّ: في السطر العربي تقع النقطة على **يسار** الرقم
                # (بين الرقم والنصّ). وسلسلة مثل «1.» لا تحتوي محرفاً قوي الاتجاه،
                # فلا تعيد خوارزمية Bidi ترتيبها وتُبنى يدوياً هنا.
                digits = A.format_number(counter, A.numeral_style())
                marker = ("." + digits) if rtl_dir else (digits + ".")
                counter += 1
            else:
                marker = "\u2022" if lvl == 0 else "\u2013"
            mk_style = ParagraphStyle(f"mk{lvl}", parent=self.st["bullet"],
                                      alignment=TA_RIGHT, rightIndent=0, spaceAfter=0)
            tx_style = ParagraphStyle(f"it{lvl}", parent=self.st["bullet"],
                                      rightIndent=0 if lvl == 0 else 14, spaceAfter=0)
            # العلامة (• – 1.) وحدها بخطٍّ يحتويها: خطّ العربية لا يحمل «•»/«–»
            mk_html = f'<font name="{self.f["latin"]}">{esc(marker)}</font>'
            rows.append([para_runs(it["text"], tx_style, self.f,
                                       avail - marker_w - (14 if lvl else 0)),
                         Paragraph(mk_html, mk_style)])
        t = Table(rows, colWidths=[avail - marker_w, marker_w],
                  style=TableStyle([("VALIGN", (0, 0), (-1, -1), "TOP"),
                                    ("LEFTPADDING", (0, 0), (-1, -1), 0),
                                    ("RIGHTPADDING", (0, 0), (-1, -1), 0),
                                    ("TOPPADDING", (0, 0), (-1, -1), 1),
                                    ("BOTTOMPADDING", (0, 0), (-1, -1), 4)]))
        return [Spacer(1, 4), t, Spacer(1, 8)]

    def code_block(self, b: Block):
        """مقطع برمجي أحادي المسافة، مع تغليف تلقائي للسطور الطويلة حتى لا تتجاوز الهوامش."""
        import textwrap

        from reportlab.platypus import XPreformatted
        size = self.th["sizes"]["body_pt"] * 0.8
        char_w = pdfmetrics.stringWidth("M", self.f["mono"], size) or 6
        max_chars = max(20, int(self.text_width / char_w) - 2)
        out_lines = []
        for ln in b.text.split("\n"):
            if not ln.strip():
                out_lines.append("")
            else:
                indent = len(ln) - len(ln.lstrip())
                out_lines += textwrap.wrap(ln, width=max_chars, break_long_words=True,
                                           break_on_hyphens=False,
                                           subsequent_indent=" " * (indent + 4)) or [""]
        return XPreformatted("\n".join(out_lines), self.st["code"])

    def table_block(self, b: Block) -> list[Flowable]:
        th, f = self.th, self.f
        header, rows = b.data["header"], b.data["rows"]
        ncol = max([len(header)] + [len(r) for r in rows])
        header = header + [""] * (ncol - len(header))
        rows = [r + [""] * (ncol - len(r)) for r in rows]
        widths = self._col_widths([header] + rows, ncol)

        rtl = th["direction"] == "rtl"
        # في الجداول العربية يكون العمود الأول على اليمين، ومحرّك ReportLab يرصف
        # الأعمدة من اليسار دائماً؛ فنعكس ترتيب الأعمدة (بمعروضاتها معها).
        order = list(range(ncol - 1, -1, -1)) if rtl else list(range(ncol))
        widths = [widths[i] for i in order]

        def cell(text: str, style: ParagraphStyle, wide: float) -> Flowable:
            """خلية جدول: يُكسَّر نصّها على عرض عمودها نفسه قبل التشكيل."""
            from manuscript import _inline

            def font_of(kind: str, chunk: str = "") -> str:
                table = ({"b": f["body_bold"], "i": f["body_italic"], "c": f["mono"]}
                         if TF.is_arabic_run(chunk) else
                         {"b": f["latin_bold"], "i": f["latin_italic"], "c": f["mono"]})
                return table.get(kind or "", f["body"] if TF.is_arabic_run(chunk) else f["latin"])

            plain = A.normalize_numerals(text)
            # الخلايا القصيرة (أرقام، كلمات مفردة) تُوسّط، والطويلة تُحاذى لليمين
            st = style
            if len(plain.strip()) > 14:
                st = ParagraphStyle(f"{style.name}_rtl", parent=style,
                                    alignment=TA_RIGHT if rtl else TA_LEFT)
            html = TF.paragraph_html(_inline(plain), max(wide - 12, 26),
                                     style.fontSize, font_of, style=st)
            return mark(Paragraph(html or " ", st), _book_level=None)

        data = [[cell(header[i], self.st["cell_head"], widths[j])
                 for j, i in enumerate(order)]]
        for r in rows:
            data.append([cell(r[i], self.st["cell"], widths[j]) for j, i in enumerate(order)])
        t = Table(data, colWidths=widths, repeatRows=1, hAlign="CENTER")
        cs = th["colors"]
        style = [
            ("GRID" if th["objects"]["table_style"] != "minimal" else "LINEBELOW",
             (0, 0), (-1, -1), 0.4, colors.Color(*hex_to_rgb(cs["rule"]))),
            ("BACKGROUND", (0, 0), (-1, 0), colors.Color(*hex_to_rgb(cs["table_header_bg"]))),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("TOPPADDING", (0, 0), (-1, -1), 5),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
            ("LEFTPADDING", (0, 0), (-1, -1), 5),
            ("RIGHTPADDING", (0, 0), (-1, -1), 5),
        ]
        if th["objects"]["table_style"] == "zebra":
            for i in range(1, len(data)):
                if i % 2 == 0:
                    style.append(("BACKGROUND", (0, i), (-1, i),
                                  colors.Color(*hex_to_rgb(cs["table_zebra_bg"]))))
        t.setStyle(TableStyle(style))
        out: list[Flowable] = [Spacer(1, 8), t]
        if th["objects"]["caption_numbering"] == "chapter":
            self.table_n += 1
            num = (f'{A.format_number(self.chapter, A.numeral_style())}'
                   f'-{A.format_number(self.table_n, A.numeral_style())}')
        else:
            self.table_n += 1
            num = A.format_number(self.table_n, A.numeral_style())
        cap = b.data.get("caption") or ""
        pre = th["objects"]["caption_prefix_table"]
        out.append(para_runs(f"{pre} ({num}) {cap}".strip(), self.st["caption"],
                             self.f, self.text_width))
        return [KeepTogether([out[0], out[1]]), out[2]]

    def _col_widths(self, rows: list[list[str]], ncol: int) -> list[float]:
        """توزيع عرض الأعمدة دون كسر الكلمات.

        الحد الأدنى لكل عمود هو عرض أطول **كلمة** فيه (زائد الحشو)، حتى لا
        يضطرّ ReportLab إلى قطع كلمة عربية بين سطرين — وهو تشويه واضح في
        الجداول العربية. الفائض يُوزَّع بنسبة طول نصوص الأعمدة.
        """
        size = self.th["sizes"]["table_pt"]
        pad = 12.0
        word_w = [0.0] * ncol
        line_w = [0.0] * ncol
        for row in rows:
            for i, txt in enumerate(row[:ncol]):
                vis = V(str(txt))
                fo = lambda k, t: self.f["body"] if TF.is_arabic_run(t) else self.f["latin"]  # noqa: E731
                line_w[i] = max(line_w[i], TF.run_width(vis, "", size, fo))
                for w in vis.split():
                    word_w[i] = max(word_w[i], TF.run_width(w, "", size, fo))
        avail = self.text_width
        mins = [min(word_w[i] + pad, avail * 0.58) for i in range(ncol)]
        base = sum(mins) or 1.0
        if base >= avail:                      # لا مجال: نتقاسم بالتناسب
            k = avail / base
            return [m * k for m in mins]
        rest = avail - base
        wants = [max(line_w[i] + pad - mins[i], 0.0) for i in range(ncol)]
        tot = sum(wants) or 1.0
        return [mins[i] + rest * wants[i] / tot for i in range(ncol)]

    def figure(self, b: Block) -> list[Flowable]:
        src = b.data.get("src", "")
        if not os.path.exists(src):
            return [para_runs(f"[صورة غير موجودة: {src}]", self.st["caption"],
                              self.f, self.text_width)]
        try:
            from PIL import Image as PILImage
            with PILImage.open(src) as im:
                iw, ih = im.size
        except Exception:  # noqa: BLE001
            iw, ih = 800, 600
        max_w = self.text_width * self.th["objects"]["image_max_width_ratio"]
        w = min(max_w, iw)
        h = w * ih / iw
        img = Image(src, width=w, height=h)
        img.hAlign = "CENTER"
        self.fig_n += 1
        num_style = A.numeral_style()
        num = (f'{A.format_number(self.chapter, num_style)}-{A.format_number(self.fig_n, num_style)}'
               if self.th["objects"]["caption_numbering"] == "chapter"
               else A.format_number(self.fig_n, num_style))
        cap = b.data.get("caption") or b.text or ""
        pre = self.th["objects"]["caption_prefix_figure"]
        caption = para_runs(f"{pre} ({num}) {cap}".strip(), self.st["caption"],
                            self.f, self.text_width)
        return [Spacer(1, 6), KeepTogether([img, caption])]

    def callout(self, b: Block) -> list[Flowable]:
        kind_titles = {"note": "ملاحظة", "tip": "نصيحة", "warning": "تنبيه",
                       "example": "مثال", "quote": "اقتباس", "exercise": "تمرين",
                       "definition": "تعريف", "key": "مفتاح"}
        kind = b.data.get("kind", "note")
        title = b.data.get("title") or kind_titles.get(kind, "ملاحظة")
        inner: list[Flowable] = []
        if title:
            inner.append(para_runs(title, self.st["callout_title"], self.f,
                                   self.text_width - 22))
        for sub in b.data.get("blocks", []) or [Block("p", text=b.text)]:
            if sub.t == "p":
                inner.append(para_runs(sub.text, self.st["callout"], self.f,
                                       self.text_width - 22))
            elif sub.t == "list":
                inner += self.bullet_list(sub, width=self.text_width - 22)
        cs = self.th["colors"]
        bg = colors.Color(*hex_to_rgb(cs["callout_bg"]))
        border = colors.Color(*hex_to_rgb(cs["callout_border"]))
        t = Table([[inner]], colWidths=[self.text_width],
                  style=TableStyle([
                      ("BACKGROUND", (0, 0), (-1, -1), bg),
                      ("LINEBEFORE" if self.th["direction"] == "rtl" else "LINEAFTER",
                       (0, 0), (0, -1), 2.2, border),
                      ("VALIGN", (0, 0), (-1, -1), "TOP"),
                      ("TOPPADDING", (0, 0), (-1, -1), 8),
                      ("BOTTOMPADDING", (0, 0), (-1, -1), 8),
                      ("LEFTPADDING", (0, 0), (-1, -1), 10),
                      ("RIGHTPADDING", (0, 0), (-1, -1), 10),
                  ]))
        return [Spacer(1, 6), KeepTogether([t]), Spacer(1, 10)]

    def build(self, blocks: list[Block], meta: dict) -> list[Flowable]:
        story: list[Flowable] = []
        for b in blocks:
            if b.t == "h":
                story += self.heading(b)
            elif b.t == "p":
                story += self.paragraph(b)
            elif b.t == "list":
                story += self.bullet_list(b)
            elif b.t == "table":
                story += self.table_block(b)
            elif b.t == "figure":
                story += self.figure(b)
            elif b.t == "callout":
                story += self.callout(b)
            elif b.t == "code":
                story += [Spacer(1, 6), KeepTogether([self.code_block(b)]), Spacer(1, 10)]
            elif b.t == "rule":
                story.append(OrnamentalBreak(self.text_width, self.accent))
            elif b.t == "pagebreak":
                story.append(PageBreak())
            elif b.t == "toc":
                story.append(TOCFlowable(list(self.toc_entries), self.th, self.f, self.text_width))
        return story


def front_matter_story(meta, styles, theme, fonts, text_width) -> list[Flowable]:
    out: list[Flowable] = []
    if meta.get("dedication"):
        out += [Spacer(1, 70), para_runs(meta["dedication"], styles["quote"], fonts, text_width)]
    if meta.get("epigraph"):
        out += [Spacer(1, 20), para_runs(meta["epigraph"], styles["quote"], fonts, text_width)]
    return out


def cover_story(meta, styles, theme, fonts, text_width) -> list[Flowable]:
    out: list[Flowable] = [Spacer(1, 26)]
    if meta.get("series"):
        out.append(para_runs(meta["series"], styles["cover_subtitle"], fonts, text_width))
        out.append(Spacer(1, 6))
    out.append(para_runs(meta.get("title", "عنوان الكتاب"), styles["cover_title"], fonts, text_width))
    out.append(Spacer(1, 10))
    rule = Rule(text_width * 0.24, 1.2, colors.Color(*hex_to_rgb(theme["colors"]["accent"])), 6)
    rule.hAlign = "CENTER"
    out.append(rule)
    if meta.get("subtitle"):
        out.append(Spacer(1, 8))
        out.append(para_runs(meta["subtitle"], styles["cover_subtitle"], fonts, text_width))

    # لوحة الغلاف (اختيارية): صورة توليدية بلا نص، تُقاس بنسبة أبعادها الفعلية
    cv = meta.get("cover_image", "")
    if cv:
        cvp = cv if os.path.isabs(cv) else os.path.join(os.path.dirname(os.path.dirname(__file__)), "book", cv)
        if os.path.exists(cvp):
            try:
                from PIL import Image as PILImage
                with PILImage.open(cvp) as im:
                    iw, ih = im.size
            except Exception:  # noqa: BLE001
                iw, ih = 1000, 1500
            w = text_width * 0.42
            h = w * ih / iw
            img = Image(cvp, width=w, height=h)
            img.hAlign = "CENTER"
            out += [Spacer(1, 12), img]
    out.append(Spacer(1, 16))
    out.append(para_runs(meta.get("author", ""), styles["cover_author"], fonts, text_width))
    if meta.get("publisher") or meta.get("year"):
        out.append(Spacer(1, 22))
        line = "  \u2022  ".join(x for x in [meta.get("publisher", ""), str(meta.get("year", ""))] if x)
        out.append(para_runs(line, styles["cover_subtitle"], fonts, text_width))
    return out


def frame_blank_pages(pdf_path: str, theme: dict) -> set[int]:
    """أرقام الصفحات الخالية داخل إطار النص (لا نصّ في المتن، ولو كان فيها ترويسة).

    تُقرأ من الملف الناتج لا من داخل محرّك التخطيط: الصفحة التي نتخطّاها لضبط جهة
    الفصل تصير فاصلة فقط إن لم يكن فيها نصّ سابق — وهذا لا يُعرف إلا بعد الرسم.
    """
    try:
        import pymupdf
    except ImportError:          # بلا مكتبة قراءة، نكتفي بالسلوك الافتراضي
        return set()
    p = theme["page"]
    top = p["margin_top_mm"] * MM - 2
    bottom = p["height_pt"] - p["margin_bottom_mm"] * MM + 2
    out: set[int] = set()
    doc = pymupdf.open(pdf_path)
    try:
        for i, page in enumerate(doc, 1):
            empty = True
            for blk in page.get_text("dict")["blocks"]:
                y0, y1 = blk["bbox"][1], blk["bbox"][3]
                if y1 <= top or y0 >= bottom or not blk.get("lines"):
                    continue
                if any(ch.strip() for l in blk["lines"] for sp in l["spans"]
                       for ch in sp["text"]):
                    empty = False
                    break
            if empty:
                out.add(i)
    finally:
        doc.close()
    return out


def build_pdf(blocks: list[Block], meta: dict, out_path: str, theme_path: str | None = None,
              cover: bool = True, toc: bool = True, max_passes: int = 6) -> str:
    theme = load_theme(theme_path)
    A.set_numeral_style(theme.get("numerals", "western"))
    A.set_base_dir(theme.get("direction", "rtl"))   # اتجاه الفقرة في Bidi (R للعربية)
    if theme["running"].get("front_matter_numbering") != "roman":
        theme["running"]["front_matter_numbering"] = A.numeral_style()
    fonts = register_fonts(theme)
    styles = build_styles(theme, fonts)
    theme["_styles"] = styles
    text_width = (theme["page"]["width_pt"]
                  - (theme["page"]["margin_outer_mm"] * 2 + theme["page"]["gutter_mm"]) * MM)

    toc_entries: list[tuple[int, str, int]] = []
    prev: list[tuple[int, str, int]] | None = None
    last_path = out_path

    blank_pages: set[int] = set()
    for pidx in range(max_passes):
        doc = BookDoc(out_path, theme, fonts, meta, styles, toc_entries, blank_pages)
        story: list[Flowable] = []
        if cover:
            story += cover_story(meta, styles, theme, fonts, text_width)
        front = front_matter_story(meta, styles, theme, fonts, text_width)
        if front:
            story += [SectionSwitch("front"), PageBreak()] + front
        if toc and theme["toc"]["enabled"]:
            story += [SectionSwitch("front"), PageBreak(),
                      para_runs("المحتويات", styles["front_h"], fonts, text_width),
                      Spacer(1, 6),
                      TOCFlowable(list(prev or []), theme, fonts, text_width)]
        builder = StoryBuilder(theme, fonts, styles, meta, toc_entries)
        story += [SectionSwitch("body")]
        if blocks and blocks[0].t != "h":
            story.append(PageBreak())
        story += builder.build(blocks, meta)
        doc.build(story)
        last_path = out_path
        blanks = frame_blank_pages(out_path, theme)
        if prev is not None and toc_entries == prev and blanks == blank_pages:
            break
        prev = list(toc_entries)
        toc_entries = []
        blank_pages = blanks
    # تشخيص: مطابقة الأسطر المرسومة فعلاً بما صمّمناه (انظر textflow.paragraph_html)
    if TF.FALLBACKS or TF.MISFITS:
        print(f"  · ضبط الأسطر: {TF.FALLBACKS} صُحّحت، {TF.MISFITS} لم تُطابق")
    return last_path
