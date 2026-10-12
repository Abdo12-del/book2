#!/usr/bin/env python3
"""تقسيم الفقرات العربية إلى أسطر، ثم تشكيل كل سطر على حدة.

لماذا وحدة مستقلة؟
    محرّك ReportLab يقسّم السطور على السلسلة **المرئية** التي نمرّرها له، من
    اليسار إلى اليمين. فإذا شكّلنا الفقرة كاملةً (Bidi + Reshaping) قبل تمريرها،
    وقع خطآن بنيويان:

    ١. **ترتيب الأسطر ينعكس:** السطر الأول (الأعلى) صار يحمل *آخر* الجملة
       منطقياً، لأن التقسيم يجري على سلسلة معكوسة الكلمات؛ فيقرأ القارئ الفقرة
       من أسفلها إلى أعلاها.
    ٢. **الترقيم يقفز إلى غير موضعه:** الفاصلة أو النقطة التي تلي كلمة في نهاية
       سطر منطقي تظهر في بداية السطر التالي، لأن حسم المحارف المحايدة (Neutral)
       يجري على مستوى الفقرة لا على مستوى السطر.

    الحل: نقسّم النص **منطقياً** (بترتيبه الأصلي) إلى أسطر، ثم نُشكّل كل سطر
    وحده. بذلك يستقيم ترتيب الأسطر، ويستقر الترقيم في نهاية السطر الصحيح، ويبقى
    التشكيل السياقي سليماً داخل السطر.

طريقة العمل:
    1. تُفكَّك المقاطع المنسّقة (**عريض** / *مائل* / `كود`) إلى كلمات، ويُرفع
       التنسيق إلى مستوى الكلمة كاملةً — لأن قطع الكلمة الواحدة بين نمطين يمنع
       اتصال حروفها عند التشكيل، فيظهر «والمبادرة» بحروف مقطوعة.
    2. تُقاس الكلمات، وتُملأ السطور بالجشع (Greedy) بعرض الإطار المتاح.
    3. لكل سطر: يُعطى السطر كاملاً لخوارزمية التشكيل (Bidi + Reshaping) فتستقيم
       مواضع الترقيم، ثم تُقسَّم نتيجته إلى أجزاء تُلبَّس كل واحد نمطه الأصلي.
"""
from __future__ import annotations

import re
import unicodedata
import xml.sax.saxutils as saxutils

import arabic_reshaper
from bidi.algorithm import get_display
from reportlab.pdfbase import pdfmetrics

import arabic as A          # مصدر واحد لاتجاه الفقرة الأساسي (A.BASE_DIR)

_SEP = "\u2000"          # فاصل محايد (WS) — انظر الشرح في tools/arabic.py
_SPACE = " "
SAFETY_PT = 1.2          # هامش أمان حتى لا يعيد ReportLab الكسر من عنده

_ARABIC_RE = re.compile(r"[\u0600-\u06FF\u0750-\u077F\uFB50-\uFDFF\uFE70-\uFEFF]")


def _unshape(text: str) -> str:
    """يُرجع المحارف السياقية (Presentation Forms) إلى أصلها — للمقارنة فقط."""
    return unicodedata.normalize("NFKC", text)


class Word:
    """كلمة واحدة بنمطها الواحد (بعد رفع التنسيق إلى مستوى الكلمة)."""

    __slots__ = ("text", "kind", "width", "visual")

    def __init__(self, text: str, kind: str = "") -> None:
        self.text, self.kind = text, kind
        self.width = 0.0
        self.visual = ""

    def __repr__(self) -> str:  # pragma: no cover - للتشخيص
        return f"Word({self.text!r}, {self.kind!r})"


# --------------------------------------------------------------------- الكلمات
def _kind_of(kinds: list[str]) -> str:
    """النمط السائد للكلمة: عريض > مائل > كود > عادي."""
    for k in ("b", "i", "c"):
        if k in kinds:
            return k
    return ""


def split_words(segments: list[tuple[str, str]]) -> list[Word]:
    """يفكّك مقاطع التنسيق إلى كلمات، مع رفع التنسيق إلى الكلمة كاملةً.

    المُدخل: [("", "قال "), ("b", "المدرب"), ("", " للطفل")]
    المُخرج: [Word("قال"), Word("المدرب", "b"), Word("للطفل")]
    """
    words: list[Word] = []
    buf: list[str] = []
    kinds: list[str] = []

    def flush() -> None:
        if buf:
            words.append(Word("".join(buf), _kind_of(kinds)))
            buf.clear()
            kinds.clear()

    for kind, text in segments:
        for ch in text:
            if ch.isspace():
                flush()
            else:
                buf.append(ch)
                if kind:
                    kinds.append(kind)
    flush()
    return words


def is_arabic_run(text: str) -> bool:
    """هل في النصّ محرف عربي أو شكل سياقي أو حركة؟"""
    return bool(_ARABIC_RE.search(text))


def script_runs(text: str) -> list[tuple[bool, str]]:
    """يقسّم نصّاً مرئياً إلى مجموعات متجانسة: ``(عربي؟, النصّ)``.

    لماذا؟ خطّ العربية المختار (Noto Sans Arabic) لا يحتوي الحروف اللاتينية ولا
    الترقيم اللاتيني (— … « » ·)، وخطّ اللاتينية (Noto Sans) لا يحتوي الأشكال
    العربية. تمرير نصّ مختلط إلى خطّ واحد يُنتج محارف فارغة (NUL)، فنجعل لكل
    مجموعة خطّها. الفراغ يبقى مع المجموعة الجارية حتى لا تتفتّت المجموعة الواحدة.
    """
    out: list[tuple[bool, str]] = []
    buf, cur = "", None
    for ch in text:
        ar = bool(_ARABIC_RE.match(ch))
        if cur is None:
            cur = ar
        elif ch.isspace():
            pass                       # الفراغ يتبع المجموعة الحالية
        elif ar != cur:
            out.append((cur, buf))
            buf, cur = "", ar
        buf += ch
    if buf:
        out.append((bool(cur), buf))
    return out


def run_width(text: str, kind: str, size: float, font_of) -> float:
    """عرض نصّ بخطوط مجموعاته (عربية/لاتينية) — يُستخدم في القياس والتقسيم."""
    return sum(pdfmetrics.stringWidth(part, font_of(kind, part), size)
               for _, part in script_runs(text))


def _shape_word(word: Word) -> str:
    """الشكل المرئي لكلمة منفردة (يُستخدم للقياس فقط)."""
    return get_display(arabic_reshaper.reshape(word.text), base_dir=A.BASE_DIR)


def prepare(words: list[Word], size: float, font_of) -> None:
    """يحسب الشكل المرئي والعرض لكل كلمة."""
    for w in words:
        w.visual = _shape_word(w)
        w.width = run_width(w.visual, w.kind, size, font_of)


def measure(words: list[Word], size: float, font_of) -> float:
    """عرض سطر مكوّن من كلمات بخطوطها، مع فراغ واحد بين كل كلمتين."""
    if not words:
        return 0.0
    total = sum(w.width for w in words)
    total += (len(words) - 1) * pdfmetrics.stringWidth(_SPACE, font_of(""), size)
    return total


# ------------------------------------------------------------------- التقسيم
def break_lines(words: list[Word], width: float, size: float, font_of) -> list[list[Word]]:
    """تقسيم جشع إلى أسطر. الكلمة الأطول من السطر تُقسَّم إن كانت لاتينية."""
    prepare(words, size, font_of)
    space_w = pdfmetrics.stringWidth(_SPACE, font_of(""), size)
    lines: list[list[Word]] = []
    cur: list[Word] = []
    for w in _split_overlong(words, width, size, font_of):
        if not cur:
            cur = [w]
            continue
        if measure(cur, size, font_of) + space_w + w.width <= width:
            cur.append(w)
        else:
            lines.append(cur)
            cur = [w]
    if cur:
        lines.append(cur)
    _avoid_widow(lines)
    return lines


def _split_overlong(words: list[Word], width: float, size: float, font_of) -> list[Word]:
    """يقطع الكلمات اللاتينية الطويلة (روابط، معرّفات) حتى لا تتجاوز الإطار."""
    out: list[Word] = []
    for w in words:
        if w.width <= width or _ARABIC_RE.search(w.text):
            out.append(w)
            continue
        buf = ""
        for ch in w.text:
            trial = buf + ch
            if run_width(get_display(trial, base_dir=A.BASE_DIR), w.kind, size, font_of) > width and buf:
                out.append(Word(buf, w.kind))
                buf = ch
            else:
                buf = trial
        if buf:
            out.append(Word(buf, w.kind))
    prepare(out, size, font_of)
    return out


def _avoid_widow(lines: list[list[Word]]) -> None:
    """يمنع بقاء كلمة وحيدة في السطر الأخير (سطر «يتيم» يشوّه الكتلة)."""
    if len(lines) < 2 or len(lines[-1]) != 1:
        return
    prev = lines[-2]
    if len(prev) <= 2:
        return
    lines[-1].insert(0, prev.pop())


# ------------------------------------------------------------------- التشكيل
def visual_parts(line: list[Word]) -> list[str]:
    """يشكّل السطر كاملاً (لسلامة الترقيم) ويعيد أجزاءه بالترتيب المرئي."""
    joined = _SEP.join(w.text for w in line)
    visual = get_display(arabic_reshaper.reshape(joined), base_dir=A.BASE_DIR)
    return [p for p in visual.split(_SEP) if p != ""]


def _map_styles(line: list[Word], parts: list[str]) -> list[str]:
    """يطابق الأجزاء المرئية بالكلمات المنطقية لإسناد نمط كل جزء.

    الترتيب المرئي لكلمات السطر العربي هو معكوس ترتيبها المنطقي، لكن المحارف
    المحايدة (بين قوسين أو حول أرقام) قد تنتقل. لذلك نُطابق بالمحتوى أولاً،
    ونرجع إلى الترتيب المعكوس عند التباس المطابقة.
    """
    n = len(line)
    if len(parts) != n:                      # حالات نادرة: التبست حدود الكلمات
        return [line[n - 1 - i].kind if 0 <= n - 1 - i < n else "" for i in range(len(parts))]
    plain = [_unshape(p) for p in parts]
    used = [False] * n
    kinds: list[str | None] = [None] * n
    for i, part in enumerate(plain):
        for j, w in enumerate(line):
            if not used[j] and _unshape(w.text) == part:
                used[j] = True
                kinds[i] = w.kind
                break
    for i in range(n):
        if kinds[i] is None:                 # لم تُطابَق: المقابل بالترتيب المعكوس
            kinds[i] = line[n - 1 - i].kind
    return [k or "" for k in kinds]


def line_html(line: list[Word], font_of) -> str:
    """HTML لسطر واحد: أجزاؤه المرئية بترتيبها الصحيح، كل جزء بخطّه."""
    parts = visual_parts(line)
    kinds = _map_styles(line, parts)
    out: list[str] = []
    for part, kind in zip(parts, kinds):
        # كل جزء قد يخلط عربياً ولاتينياً (مثل «وTED») فنُلبّس كل مجموعة خطّها
        out.append("".join(
            f'<font name="{font_of(kind, run)}">{saxutils.escape(run)}</font>'
            for _, run in script_runs(part)))
    return _SPACE.join(out)


_STRIP_TAGS = re.compile(r"<[^>]+>")


def _plain(html: str) -> str:
    """نصّ السطر بلا وسوم HTML وبلا فراغات — للمقارنة فقط."""
    return re.sub(r"\s+", "", _STRIP_TAGS.sub("", html))


def _rendered_lines(html: str, style, width: float) -> list[str] | None:
    """الأسطر التي يرسمها ReportLab فعلاً، كنصوص مرئية (``None`` إن تعذّرت القراءة).

    يُقرأ الهيكل الداخلي ``blPara`` مباشرة: النوع 0 يعطي ``(فراغ، كلمات)`` والنوع 1
    (فقرات متعددة الأنماط) يعطي ``FragLine`` بكلمات لكل مقطع.
    """
    from reportlab.platypus import Paragraph

    try:
        p = Paragraph(html or " ", style)
        p.wrap(width, 10 ** 7)
        bl = p.blPara
        out: list[str] = []
        for line in bl.lines:
            if bl.kind == 1:
                out.append("".join(f.text for f in line.words))
            else:
                out.append("".join(line[1]))
        return out
    except Exception:  # noqa: BLE001
        return None


def _rendered_line_count(html: str, style, width: float) -> int:
    """عدد الأسطر المرسومة فعلاً (لأدوات القياس)."""
    lines = _rendered_lines(html, style, width)
    return -1 if lines is None else len(lines)


# عدد الفقرات التي احتاجت تقليص الأسطر لأن مقاييس ReportLab جاءت أوسع من مقاييسنا.
FALLBACKS = 0
# عدد الفقرات التي لم تُطابق مقاييس ReportLab رغم كل المحاولات (يجب أن يبقى صفراً).
MISFITS = 0


def paragraph_html(segments: list[tuple[str, str]], width: float, size: float,
                   font_of, style=None) -> str:
    """HTML الفقرة: أسطر مكسورة منطقياً، كل سطر مُشكَّل على حدة.

    ``font_of(kind)`` يعيد اسم الخط لنمط: "" عادي، "b" عريض، "i" مائل، "c" كود.

    إذا مُرِّر ``style`` نتحقّق من النتيجة: نقيس ما سيرسمه ReportLab فعلاً
    (``blPara.lines``) ونطابقه بعدد أسطرنا. هذا ضروري لأن ReportLab يطبّق
    تقنيناً (Kerning) وربطاً (Ligatures) بمقاييسه هو، فقد يقسّم سطراً حسبناه
    متّسعاً — فينزلق آخر عنصر إلى السطر التالي وتتغيّر مواضع الترقيم. عند
    الاختلاف نُعيد الكسر بأسطر أضيق تدرّجياً حتى تتطابق النتيجتان.
    """
    global FALLBACKS

    words = split_words(segments)
    if not words:
        return ""
    slacks = (SAFETY_PT, 3.5, 6.0, 9.0, 13.0) if style is not None else (SAFETY_PT,)
    html = ""
    for i, slack in enumerate(slacks):
        lines = break_lines(words, max(width - slack, 10.0), size, font_of)
        parts = [line_html(line, font_of) for line in lines]
        html = "<br/>".join(parts)
        if style is None:
            break
        rendered = _rendered_lines(html, style, width)
        if rendered is not None and len(rendered) == len(parts) and all(
                _plain(r) == _plain(x) for r, x in zip(rendered, parts)):
            if i:
                FALLBACKS += 1
            break
    else:
        # لم تُطابق أي محاولة — نُبلّغ بدل أن نُخفي المشكلة
        MISFITS += 1
    return html


def plain_visual(text: str) -> str:
    """نصّ سطر واحد مرئي (للعناوين والسطور التي لا تُكسَّر)."""
    return get_display(arabic_reshaper.reshape(text), base_dir=A.BASE_DIR)


__all__ = ["Word", "split_words", "prepare", "break_lines", "visual_parts",
           "is_arabic_run", "script_runs", "run_width",
           "line_html", "paragraph_html", "plain_visual", "measure", "SAFETY_PT"]
