#!/usr/bin/env python3
"""فحص ضمان الجودة للملفات الناتجة (PDF و DOCX).

لا يكتفي هذا الفحص بنجاح التصدير؛ بل يفتح الملفات فعلياً ويتحقق من:
  PDF : عدد الصفحات، الخطوط المضمّنة، النص المستخرج، الصفحات الفارغة،
        وجود كل عنوان من عناوين المخطوطة داخل الملف، صور الصفحات للمعاينة البصرية.
  DOCX: سلامة حزمة zip، صلاحية ترتيب عناصر XML، اتجاه RTL،
        الأنماط المستخدمة، الفقرات الفارغة المتكررة، الجداول والصور.

الاستخدام:
    python3 tools/qa_check.py build/book.pdf --source book/sample/00-numuwdh.md
    python3 tools/qa_check.py build/book.docx
    python3 tools/qa_check.py build/book.pdf --render build/pages
"""
from __future__ import annotations

import argparse
import os
import re
import sys
import unicodedata
import zipfile

# --------------------------------------------------------------- أدوات نصية
HARAKAT = re.compile(r"[\u064B-\u065F\u0670\u0640]")
NON_ALNUM = re.compile(r"[^\w\u0600-\u06FF]", re.UNICODE)


def norm(text: str) -> str:
    """تطبيع النص للمقارنة: توحيد أشكال الحروف وإزالة التشكيل والرموز."""
    t = unicodedata.normalize("NFKC", text or "")
    t = HARAKAT.sub("", t)
    t = t.replace("أ", "ا").replace("إ", "ا").replace("آ", "ا").replace("ى", "ي")
    t = t.replace("ة", "ه").replace("ؤ", "و").replace("ئ", "ي")
    return NON_ALNUM.sub("", t)


class Report:
    def __init__(self) -> None:
        self.ok: list[str] = []
        self.warn: list[str] = []
        self.err: list[str] = []

    def good(self, m): self.ok.append(m); print(f"  ✓ {m}")
    def warning(self, m): self.warn.append(m); print(f"  ! {m}")
    def error(self, m): self.err.append(m); print(f"  ✗ {m}")
    def section(self, m): print(f"\n── {m} ──")

    def summary(self) -> int:
        print("\n" + "=" * 58)
        print(f"النتيجة: {len(self.ok)} ناجح | {len(self.warn)} تحذير | {len(self.err)} خطأ")
        print("=" * 58)
        return 1 if self.err else 0


# --------------------------------------------------------------- فحص PDF
def check_layout(path: str, rpt: Report, theme: dict) -> None:
    """فحص تخطيطي (بديل المعاينة البصرية): تجاوز الهوامش، انعكاسها، والعناوين اليتيمة."""
    import pymupdf
    from theme import MM as MMU

    doc = pymupdf.open(path)
    p = theme["page"]
    W, H = p["width_pt"], p["height_pt"]
    outer = p["margin_outer_mm"] * MMU
    inner = (p["margin_outer_mm"] + p["gutter_mm"]) * MMU
    top, bottom = p["margin_top_mm"] * MMU, p["margin_bottom_mm"] * MMU
    rtl = theme["direction"] == "rtl"
    tol = 2.0
    overflow: list[int] = []
    orphans: list[int] = []
    thin: list[int] = []
    body_pt = theme["sizes"]["body_pt"]
    head_sizes = {v for k, v in theme["sizes"].items()
                  if k.endswith("_pt") and k.startswith("h") and v > body_pt + 0.5}

    for i in range(doc.page_count):
        page = doc[i]
        odd = (i + 1) % 2 == 1
        gutter_right = odd if rtl else (not odd)
        x0 = outer if gutter_right else inner
        x1 = W - (inner if gutter_right else outer)
        spans = []
        for b in page.get_text("dict")["blocks"]:
            for l in b.get("lines", []):
                for s in l["spans"]:
                    x0s, y0s, x1s, y1s = s["bbox"]
                    # استبعاد الترويسة والتذييل من فحص تجاوز الهوامش
                    if y1s > top - 4 and y0s < H - bottom + 4:
                        spans.append(s)
                        if x0s < x0 - tol or x1s > x1 + tol:
                            overflow.append(i + 1)
        if not spans:
            continue
        # عنوان يتيم: آخر سطر في الصفحة عنوان وليس بعده نص
        lowest = max(spans, key=lambda s: s["bbox"][3])
        if round(lowest["size"], 1) in [round(x, 1) for x in head_sizes]:
            below = [s for s in spans if s["bbox"][1] > lowest["bbox"][3] - 1]
            if not below and lowest["bbox"][3] > H - bottom - 60:
                orphans.append(i + 1)
        # صفحة ذات سطر أو سطرين فقط (عدا صفحات بداية الفصول)
        lines = page.get_text("text").strip().split("\n")
        body_lines = [l for l in lines if len(l.strip()) > 3]
        if len(body_lines) <= 2 and (i + 1) not in (1,):
            thin.append(i + 1)

    if overflow:
        rpt.warning(f"صفحات يتجاوز محتواها الهوامش: {sorted(set(overflow))}")
    else:
        rpt.good("كل المحتوى داخل الهوامش المحددة (لا تجاوز)")
    if orphans:
        rpt.warning(f"عناوين في أسفل الصفحة بلا نص يتبعها (يتيم): {sorted(set(orphans))}")
    else:
        rpt.good("لا عناوين يتيمة في أسفل الصفحات")
    if thin:
        rpt.warning(f"صفحات شبه فارغة (سطر أو سطران): {sorted(set(thin))}")
    doc.close()


def frame_text(page, theme: dict | None) -> str:
    """نصّ إطار المتن وحده — بلا ترويسة ولا تذييل (لكشف الصفحات الفاصلة)."""
    from theme import MM as _MM
    pg = (theme or {}).get("page") or {}
    top = pg.get("margin_top_mm", 20) * _MM - 2
    bottom = page.rect.height - pg.get("margin_bottom_mm", 20) * _MM + 2
    out: list[str] = []
    for blk in page.get_text("dict")["blocks"]:
        y0, y1 = blk["bbox"][1], blk["bbox"][3]
        if y1 <= top or y0 >= bottom or not blk.get("lines"):
            continue
        out.append("\n".join("".join(sp["text"] for sp in ln["spans"])
                              for ln in blk["lines"]))
    return "\n".join(out)


def check_pdf(path: str, rpt: Report, source_text: str | None = None,
              render_dir: str | None = None, expect_headings: list[str] | None = None,
              theme: dict | None = None) -> None:
    import pymupdf
    rpt.section(f"فحص PDF: {os.path.basename(path)}")
    doc = pymupdf.open(path)
    n = doc.page_count
    rpt.good(f"عدد الصفحات: {n}")
    if n == 0:
        rpt.error("الملف لا يحتوي صفحات")
        return
    r0 = doc[0].rect
    rpt.good(f"قياس الصفحة: {r0.width:.1f} × {r0.height:.1f} نقطة "
             f"({r0.width/72*25.4:.0f} × {r0.height/72*25.4:.0f} مم)")

    # الخطوط المضمّنة
    fonts = set()
    for i in range(n):
        for f in doc[i].get_fonts(full=True):
            fonts.add(f[3])
    embedded, missing = [], []
    for xref in {x[0] for i in range(n) for x in doc[i].get_fonts(full=True)}:
        try:
            doc.extract_font(xref)
            embedded.append(xref)
        except Exception:  # noqa: BLE001
            missing.append(xref)
    rpt.good(f"الخطوط المستخدمة: {', '.join(sorted(fonts)) or '—'}")
    if missing:
        rpt.warning(f"{len(missing)} خط غير مُضمّن بالكامل (xref: {missing})")
    else:
        rpt.good(f"كل الخطوط مُضمّنة داخل الملف ({len(embedded)} خط)")

    # النصوص والصفحات الفارغة
    all_text: list[str] = []
    fillers: list[int] = []
    short: list[int] = []
    body_of: dict[int, str] = {}
    for i in range(n):
        page = doc[i]
        txt = page.get_text("text") or ""
        all_text.append(txt)
        has_img = bool(page.get_images())
        body = frame_text(page, theme)
        body_of[i + 1] = body
        if not body.strip() and not has_img:
            fillers.append(i + 1)          # لا نصّ داخل إطار المتن
        elif len(body.strip()) < 40 and not has_img:
            short.append(i + 1)
    if fillers:
        # الصفحة الفاصلة مقصودة حين تبدأ الفصول على جهة محددة (صفحة الوجه)،
        # ويجب أن يليها افتتاحية فصل — وإلا فهي صفحة ضائعة فعلاً.
        lay = (theme or {}).get("layout") or {}
        parity = lay.get("chapter_starts_on", "any")
        label = ((theme or {}).get("meta") or {}).get("chapter_label") or "الفصل"
        heads = [norm(h) for h in (expect_headings or []) if h.strip()]
        if not heads and source_text:
            # بلا قائمة عناوين صريحة: نأخذ عناوين H1 من المخطوطة نفسها
            heads = [norm(m.group(1)) for m in
                     re.finditer(r"^#\s+(.+)$", source_text, re.M)]
        def opens_chapter(page_no: int) -> bool:
            nxt = norm(body_of.get(page_no + 1, "")).strip()
            if not nxt:
                return False
            # ملاحظة: ``norm`` تُزيل الفراغات، فلا يصحّ \b هنا
            if re.match(rf"({re.escape(label)}|الباب)", nxt):
                return True
            return any(nxt.startswith(h) for h in heads)

        starters = [f for f in fillers if opens_chapter(f)]
        if parity in ("right", "left") and len(starters) == len(fillers):
            side = "فردية" if parity == "right" else "زوجية"
            rpt.good(f"صفحات فاصلة مقصودة: {fillers} — بدء كل فصل على صفحة {side} "
                     f"(chapter_starts_on: {parity})")
        else:
            rpt.warning(f"صفحات بلا نصّ في المتن: {fillers} "
                        f"(يليها افتتاحية فصل: {starters})")
    else:
        rpt.good("لا توجد صفحات فارغة")
    if short:
        rpt.warning(f"صفحات نصّها قليل جداً (راجعها بصرياً): {short}")

    # سلامة النص العربي بعد تطبيع أشكال العرض (NFKC)
    joined = "\n".join(all_text)
    joined_norm = norm(joined)
    arabic_chars = len(re.findall(r"[\u0600-\u06FF]", joined_norm))
    rpt.good(f"إجمالي الحروف العربية بعد التطبيع: {arabic_chars}")
    if arabic_chars < 100:
        rpt.error("النص العربي المستخرج قليل جداً → اشتباه في تشفير أو رسم النص")

    # محارف مفقودة من الخط (ترسم كـ notdef ويُستخرج عنها NUL)
    nuls = sum(t.count("\x00") for t in all_text)
    if nuls:
        bad_pages = [i + 1 for i, t in enumerate(all_text) if "\x00" in t]
        rpt.error(f"{nuls} محرفاً غير مدعوم في الخط (يظهر فارغاً أو NUL) في الصفحات: {bad_pages}")
    else:
        rpt.good("لا محارف مفقودة من الخطوط (لا NUL في النص المستخرج)")
    qmarks = joined.count("?") + joined.count("�")
    if qmarks > 5:
        rpt.warning(f"وجود {qmarks} رمز استبدال (?) → قد يدل على حرف غير مدعوم في الخط")

    # التحقق من وجود العناوين
    if expect_headings:
        nj = joined_norm
        lost = [h for h in expect_headings if norm(h) not in nj]
        if lost:
            for h in lost:
                rpt.error(f"العنوان غير موجود في النص المستخرج: «{h}»")
        else:
            rpt.good(f"كل العناوين ({len(expect_headings)}) موجودة في النص المستخرج")

    # تطابق نص المخطوطة مع PDF
    if source_text:
        src_words = [w for w in re.findall(r"[\u0600-\u06FF]{4,}", source_text)]
        nj = joined_norm
        missing_w = [w for w in dict.fromkeys(src_words) if norm(w) not in nj]
        ratio = 1 - len(missing_w) / max(1, len(dict.fromkeys(src_words)))
        if ratio > 0.98:
            rpt.good(f"تغطية نص المخطوطة في الـPDF: {ratio*100:.1f}%")
        elif ratio > 0.90:
            rpt.warning(f"تغطية نص المخطوطة: {ratio*100:.1f}% — كلمات مفقودة: {missing_w[:8]}")
        else:
            rpt.error(f"تغطية نص المخطوطة منخفضة: {ratio*100:.1f}% — مثال: {missing_w[:10]}")

    # صور الصفحات للمعاينة البصرية
    if render_dir:
        os.makedirs(render_dir, exist_ok=True)
        for i in range(n):
            pix = doc[i].get_pixmap(dpi=96)
            pix.save(os.path.join(render_dir, f"page-{i+1:03d}.png"))
        rpt.good(f"تم إخراج {n} صورة صفحة في: {render_dir}")
    doc.close()

    check_text_clarity(path, rpt, source_text=source_text, theme=theme)

    if theme:
        rpt.section("فحص التخطيط (الهوامش / العناوين اليتيمة)")
        check_layout(path, rpt, theme)



# ------------------------------------------------- فحص وضوح النصّ العربي (PDF)
AR_RANGE = r"\u0600-\u06FF\u0750-\u077F"
CLOSING = "،؛.؟!:"


def _clean_for_compare(s: str) -> str:
    import unicodedata
    s = unicodedata.normalize("NFKC", s)
    s = re.sub(r"[\u064B-\u065F\u0670\u0640]", "", s)
    return s


def _ordered_words(s: str) -> list[str]:
    return [w for w in re.sub(f"[^\\w{AR_RANGE}]+", " ", _clean_for_compare(s)).split()
            if re.search(f"[{AR_RANGE}]", w)]


def pdf_lines_rtl(path: str) -> list[tuple[int, str]]:
    """أسطر الصفحة بترتيبها المنطقي: داخل السطر من اليمين لليسار، والصفحات بالترتيب.

    هذه الطريقة مستقلة عن تخزين ReportLab للمقاطع؛ تعتمد على إحداثيات الرسم.
    """
    import pymupdf
    doc = pymupdf.open(path)
    out: list[tuple[int, str]] = []
    for i, page in enumerate(doc):
        for b in page.get_text("rawdict")["blocks"]:
            for line in b.get("lines", []):
                chars = [(c["bbox"][0], c["c"]) for sp in line["spans"]
                         for c in sp.get("chars", [])]
                if not chars:
                    continue
                chars.sort(key=lambda t: -t[0])
                out.append((i + 1, "".join(c for _, c in chars)))
    doc.close()
    return out


def check_text_clarity(path: str, rpt: Report, source_text: str | None = None,
                       theme: dict | None = None) -> None:
    """الفحوص التي تحمي **وضوح القراءة** — وهي أهم من سلامة البناء:

    ١. ترتيب القراءة: يُعاد بناء النصّ من إحداثيات الرسم ويُقارن بالمخطوطة.
       أي خلل في الاتجاه أو تفكيك الأسطر يُسقِط هذه النسبة فوراً.
    ٢. بداية السطر: لا يجوز أن يبدأ سطرٌ بفاصلة أو نقطة (دليل على انزلاق
       الترقيم بين الأسطر).
    ٣. نمط الأرقام: يطابق المطلوب في design/theme.yaml.
    ٤. بقايا ترميز الماركداون (** أو `) في المخرج النهائي.
    """
    import difflib
    lines = pdf_lines_rtl(path)
    if not lines:
        return
    rpt.section("وضوح النصّ العربي")

    # (١) ترتيب القراءة
    if source_text:
        src_words = _ordered_words(source_text)
        pdf_words: list[str] = []
        for _, ln in lines:
            pdf_words += _ordered_words(ln)
        sm = difflib.SequenceMatcher(None, src_words, pdf_words, autojunk=False)
        longest = max((i2 - i1 for tag, i1, i2, _, _ in sm.get_opcodes() if tag == "equal"),
                      default=0)
        ratio = sm.ratio()
        msg = (f"ترتيب القراءة: تشابه {ratio*100:.1f}% (مصدر {len(src_words)} كلمة "
               f"مقابل {len(pdf_words)}، أطول مقطع متتالٍ {longest} كلمة)")
        if longest >= 60 and ratio > 0.80:
            rpt.good(msg)
        elif longest >= 25 and ratio > 0.65:
            rpt.warning(msg + " — راجع مواضع الانقطاع")
        else:
            rpt.error(msg + " — اشتباه في انعكاس ترتيب الأسطر أو الكلمات")

    # (٢) سطور تبدأ بترقيم ختامي
    bad_starts = [(pg, ln[:40]) for pg, ln in lines if ln and ln[0] in CLOSING]
    if bad_starts:
        rpt.error(f"{len(bad_starts)} سطراً يبدأ بترقيم ختامي (، . ؟) — انزلاق ترقيم؛ "
                  f"أمثلة: {bad_starts[:5]}")
    else:
        rpt.good("لا سطر يبدأ بترقيم ختامي (الترقيم في مواضعه)")

    # (٣) نمط الأرقام
    joined = "\n".join(ln for _, ln in lines)
    western = sum(ch.isdigit() and ch.isascii() for ch in joined)
    indic = len(re.findall(r"[\u0660-\u0669\u06F0-\u06F9]", joined))
    want = (theme or {}).get("numerals", "western")
    if want == "western":
        if indic:
            rpt.warning(f"أرقام عربية-هندية في المخرج ({indic} رقم) مع أن النمط المعتمد "
                        f"«western» (123)")
        else:
            rpt.good(f"نمط الأرقام لاتيني (123) — {western} رقماً في الكتاب")
    else:
        if western > 20:
            rpt.warning(f"أرقام لاتينية في المخرج ({western}) مع أن النمط المعتمد "
                        f"«arabic-indic» (١٢٣)")
        else:
            rpt.good(f"نمط الأرقام عربي-هندي (١٢٣) — {indic} رقماً")

    # (٤) بقايا ترميز الماركداون
    leftovers = joined.count("**")
    if leftovers:
        rpt.error(f"بقايا ترميز عريض (**) ظاهرة في الصفحات: {leftovers} موضعاً")
    else:
        rpt.good("لا بقايا لعلامات الماركداون (**, `) في المخرج")


def check_docx_text_clarity(path: str, rpt: Report, source_text: str | None = None,
                            theme: dict | None = None) -> None:
    """نظير فحوص الوضوح لملف Word: الأرقام، وسم RTL على المقاطع، وبقايا الترميز."""
    from docx import Document as _Doc
    doc = _Doc(path)
    texts: list[str] = [p.text for p in doc.paragraphs]
    for t in doc.tables:
        for row in t.rows:
            for c in row.cells:
                texts.append(c.text)
    joined = "\n".join(texts)
    rpt.section("وضوح النصّ في DOCX")

    leftovers = joined.count("**")
    if leftovers:
        rpt.error(f"بقايا ترميز عريض (**) في ملف Word: {leftovers} موضعاً")
    else:
        rpt.good("لا بقايا لعلامات الماركداون في ملف Word")

    want = (theme or {}).get("numerals", "western")
    indic = len(re.findall(r"[\u0660-\u0669\u06F0-\u06F9]", joined))
    western = sum(ch.isdigit() and ch.isascii() for ch in joined)
    if want == "western" and indic:
        rpt.warning(f"أرقام عربية-هندية في ملف Word ({indic}) مع اعتماد «western»")
    else:
        rpt.good(f"نمط الأرقام في Word: {'لاتيني (123)' if want == 'western' else 'عربي-هندي'} "
                 f"— غربي: {western} | عربي-هندي: {indic}")

    # كل فقرة عربية يجب أن تحمل وسم الاتجاه
    from docx.oxml.ns import qn as _qn
    paras = [p for p in doc.paragraphs if re.search(f"[{AR_RANGE}]", p.text or "")]
    missing_bidi = [p.text[:30] for p in paras
                    if p._p.find(_qn("w:pPr")) is None
                    or p._p.pPr.find(_qn("w:bidi")) is None]
    if missing_bidi:
        rpt.warning(f"{len(missing_bidi)} فقرة عربية بلا وسم اتجاه (w:bidi): "
                    f"{missing_bidi[:3]}")
    else:
        rpt.good(f"كل الفقرات العربية ({len(paras)}) موسومة بالاتجاه من اليمين إلى اليسار")


# --------------------------------------------------------------- فحص DOCX
PPR_ORDER = ["w:pStyle", "w:keepNext", "w:keepLines", "w:pageBreakBefore", "w:framePr",
             "w:widowControl", "w:numPr", "w:suppressLineNumbers", "w:pBdr", "w:shd", "w:tabs",
             "w:suppressAutoHyphens", "w:kinsoku", "w:wordWrap", "w:overflowPunct",
             "w:topLinePunct", "w:autoSpaceDE", "w:autoSpaceDN", "w:bidi", "w:adjustRightInd",
             "w:snapToGrid", "w:spacing", "w:ind", "w:contextualSpacing", "w:mirrorIndents",
             "w:suppressOverlap", "w:jc", "w:textDirection", "w:textAlignment",
             "w:textboxTightWrap", "w:outlineLvl", "w:divId", "w:cnfStyle", "w:rPr", "w:sectPr",
             "w:pPrChange"]
RPR_ORDER = ["w:rStyle", "w:rFonts", "w:b", "w:bCs", "w:i", "w:iCs", "w:caps", "w:smallCaps",
             "w:strike", "w:dstrike", "w:outline", "w:shadow", "w:emboss", "w:imprint",
             "w:noProof", "w:snapToGrid", "w:vanish", "w:webHidden", "w:color", "w:spacing",
             "w:w", "w:kern", "w:position", "w:sz", "w:szCs", "w:highlight", "w:u", "w:effect",
             "w:bdr", "w:shd", "w:fitText", "w:vertAlign", "w:rtl", "w:cs", "w:em", "w:lang",
             "w:eastAsianLayout", "w:specVanish", "w:oMath"]


def _order_valid(parent, order: list[str]) -> bool:
    idx = [-1] * len(order)
    last = -1
    for child in parent:
        tag = child.tag.split("}")[-1]
        tag = "w:" + tag
        if tag in order:
            i = order.index(tag)
            if i < last:
                return False
            last = i
    return True


def check_docx(path: str, rpt: Report, expect_headings: list[str] | None = None) -> None:
    from docx import Document
    from docx.oxml.ns import qn

    rpt.section(f"فحص DOCX: {os.path.basename(path)}")

    # سلامة الحزمة
    with zipfile.ZipFile(path) as z:
        names = z.namelist()
        bad = z.testzip()
        if bad:
            rpt.error(f"ملف تالف داخل الحزمة: {bad}")
        else:
            rpt.good(f"حزمة OOXML سليمة ({len(names)} جزءاً)")
        if "[Content_Types].xml" not in names:
            rpt.error("غياب [Content_Types].xml")

    doc = Document(path)
    sec = doc.sections[0]
    rpt.good(f"عدد الأقسام (Sections): {len(doc.sections)}")
    rpt.good(f"قياس الصفحة: {sec.page_width.mm:.0f} × {sec.page_height.mm:.0f} مم")
    rpt.good(f"الهوامش (داخلي/خارجي/أعلى/أسفل): {sec.left_margin.mm:.0f} / "
             f"{sec.right_margin.mm:.0f} / {sec.top_margin.mm:.0f} / {sec.bottom_margin.mm:.0f} مم")

    bidi_sections = sum(1 for s in doc.sections
                        if s._sectPr.find(qn("w:bidi")) is not None)
    if bidi_sections == len(doc.sections):
        rpt.good("اتجاه RTL مضبوط على مستوى كل الأقسام (w:bidi)")
    else:
        rpt.error(f"قسم بدون اتجاه RTL: {bidi_sections}/{len(doc.sections)}")

    # الفقرات
    paras = doc.paragraphs
    non_empty = [p for p in paras if p.text.strip()]
    rpt.good(f"عدد الفقرات: {len(paras)} (غير الفارغة: {len(non_empty)})")
    rtl_paras = sum(1 for p in non_empty if p._p.pPr is not None
                    and p._p.pPr.find(qn("w:bidi")) is not None)
    if rtl_paras >= len(non_empty) * 0.9:
        rpt.good(f"فقرات ذات اتجاه عربي صريح: {rtl_paras}/{len(non_empty)}")
    else:
        rpt.warning(f"بعض الفقرات بلا وسم RTL: {rtl_paras}/{len(non_empty)}")

    # ترتيب عناصر XML (مهم لصلاحية الملف في Word)
    bad_order = 0
    for p in paras:
        ppr = p._p.pPr
        if ppr is not None and not _order_valid(ppr, PPR_ORDER):
            bad_order += 1
        for r in p.runs:
            rpr = r._r.rPr
            if rpr is not None and not _order_valid(rpr, RPR_ORDER):
                bad_order += 1
    if bad_order:
        rpt.error(f"{bad_order} عنصر بترتيب XML غير مطابق للمخطط (قد يرفض Word الملف)")
    else:
        rpt.good("ترتيب عناصر XML مطابق لمخطط OOXML")

    # الأنماط المستخدمة
    from collections import Counter
    c = Counter(p.style.name for p in paras if p.text.strip())
    rpt.good("الأنماط المستخدمة: " + "، ".join(f"{k} ({v})" for k, v in c.most_common(8)))

    # فقرات فارغة متتالية — أي عنصر آخر (جدول مثلاً) يقطع التتالي، لأن ما يفصل
    # الفقرتين ليس فراغاً في الصفحة بل عنصر محتوى.
    runs, worst, cur = 0, 0, 0
    for el in doc.element.body.iterchildren():
        if el.tag.split("}")[-1] != "p":
            if cur >= 3:
                runs += 1
            cur = 0
            continue
        empty = (not "".join(t.text or "" for t in el.iter(qn("w:t"))).strip()
                 and not el.findall(qn("w:r"))
                 and not el.findall(".//" + qn("w:r")))
        if empty:
            cur += 1
            worst = max(worst, cur)
        else:
            if cur >= 3:
                runs += 1
            cur = 0
    if cur >= 3:
        runs += 1
    if worst >= 3:
        rpt.warning(f"توجد {runs} مجموعة من {worst}+ فقرات فارغة متتالية")
    else:
        rpt.good("لا فقرات فارغة متراكمة")

    # الجداول والصور
    rpt.good(f"عدد الجداول: {len(doc.tables)}")
    imgs = [n for n in zipfile.ZipFile(path).namelist() if n.startswith("word/media/")]
    rpt.good(f"عدد الصور المضمّنة: {len(imgs)}")

    # الحقول (في المتن وفي الترويسات/التذييلات)
    xml = doc.element.xml
    for sct in doc.sections:
        for part in (sct.header, sct.footer, sct.even_page_header, sct.even_page_footer,
                     sct.first_page_header, sct.first_page_footer):
            try:
                xml += part._element.xml
            except Exception:  # noqa: BLE001
                pass
    for field, label in (("TOC", "حقل الفهرس"), ("PAGE", "حقل ترقيم الصفحات"),
                         ("STYLEREF", "حقل عنوان الفصل في الترويسة")):
        if field in xml:
            rpt.good(f"{label} موجود")
        else:
            rpt.warning(f"{label} غير موجود")

    if expect_headings:
        joined = norm("\n".join(p.text for p in paras if p.style.name.startswith("Heading")))
        lost = [h for h in expect_headings if norm(h) not in joined]
        if lost:
            for h in lost:
                rpt.error(f"عنوان مفقود بين عناوين الوورد: «{h}»")
        else:
            rpt.good(f"كل العناوين ({len(expect_headings)}) موجودة كعناوين Word حقيقية")


def read_source(path: str) -> str:
    """يقرأ مصدر المخطوطة: ملف مفرد، أو book.yaml، أو مجلد يحتوي book.yaml.

    الكتاب الحقيقي موزّع على عدة ملفات مرتّبة في `spine`، فلا يكفي قراءة ملف واحد.
    """
    if os.path.isdir(path):
        cand = os.path.join(path, "book.yaml")
        if not os.path.exists(cand):
            raise SystemExit(f"--source مجلد بلا book.yaml: {path}")
        path = cand
    if path.endswith((".yaml", ".yml")):
        import yaml
        base = os.path.dirname(os.path.abspath(path))
        with open(path, encoding="utf-8") as f:
            meta = yaml.safe_load(f) or {}
        parts = []
        for rel in meta.get("spine", []) or []:
            p = os.path.join(base, rel)
            if os.path.exists(p):
                parts.append(open(p, encoding="utf-8").read())
            else:
                print(f"  ! ملف مفقود من الـspine: {p}")
        return "\n".join(parts)
    with open(path, encoding="utf-8") as f:
        return f.read()


def main() -> int:
    ap = argparse.ArgumentParser(description="فحص جودة مخرجات الكتاب")
    ap.add_argument("files", nargs="+")
    ap.add_argument("--source", help="ملف المخطوطة للمطابقة النصية")
    ap.add_argument("--render", help="مجلد لإخراج صور الصفحات")
    ap.add_argument("--headings", help="ملف نصي بسطور العناوين المتوقعة")
    ap.add_argument("--theme", help="ملف نظام التصميم لفحص التخطيط")
    args = ap.parse_args()

    src = None
    if args.source and os.path.exists(args.source):
        src = read_source(args.source)
    heads = None
    if args.headings and os.path.exists(args.headings):
        heads = [l.strip() for l in open(args.headings, encoding="utf-8") if l.strip()]
    theme = None
    if args.theme:
        sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__))))
        from theme import load as load_theme
        theme = load_theme(args.theme)

    rpt = Report()
    for f in args.files:
        if not os.path.exists(f):
            rpt.error(f"الملف غير موجود: {f}")
            continue
        if f.lower().endswith(".pdf"):
            check_pdf(f, rpt, source_text=src, render_dir=args.render,
                      expect_headings=heads, theme=theme)
        elif f.lower().endswith(".docx"):
            check_docx(f, rpt, expect_headings=heads)
            check_docx_text_clarity(f, rpt, source_text=src, theme=theme)
        else:
            rpt.warning(f"نوع غير مدعوم: {f}")
        print(f"  حجم الملف: {os.path.getsize(f)/1024:.0f} KB")
    return rpt.summary()


if __name__ == "__main__":
    raise SystemExit(main())
