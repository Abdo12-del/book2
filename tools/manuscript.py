#!/usr/bin/env python3
"""مُحلّل المخطوطة: يحول نصوص الماركداون العربية إلى شجرة كُتل محايدة.

الهدف: مصدر واحد للمحتوى ← مُصدّران (DOCX و PDF) بنفس التصميم.
صيغة المخطوطة المدعومة (ماركداون موسّع بشكل خفيف):

    # الفصل الأول: العنوان            → عنوان فصل (h1)
    ## عنوان فرعي                     → h2     ### → h3     #### → h4
    نص الفقرة...                      → فقرة (تُوصل الأسطر تلقائياً)
    - عنصر                            → قائمة نقطية
    1. عنصر                           → قائمة مرقّمة (بمسافتين للإزاحة = مستوى ثانٍ)
    | العمود | العمود |               → جدول (سطر الفواصل |---|---| إلزامي)
    : عنوان الجدول                    → تعليق للجدول/الشكل السابق
    ![وصف](images/x.png "تعليق")      → صورة مع تعليق
    ::: note  عنوان                   → إطار ملاحظة (note|tip|warning|example|quote|exercise)
    ...
    :::
    ```                               → مقطع برمجي (لاتيني، من اليسار لليمين)
    ---                               → خط فاصل زخرفي
    <!-- pagebreak -->                → فاصل صفحة إجباري
    <!-- toc -->                      → موضع جدول المحتويات

    التنسيق المضمّن: **عريض**  *مائل*  `كود`
"""
from __future__ import annotations

import os
import re
from dataclasses import dataclass, field
from typing import Any

RE_H = re.compile(r"^(#{1,6})\s+(.*)$")
RE_BULLET = re.compile(r"^(\s*)[-*+]\s+(.*)$")
RE_ORDERED = re.compile(r"^(\s*)(?:\d+|٫|[٠-٩]+)[.)]\s+(.*)$")
RE_IMG = re.compile(r"^!\[([^\]]*)\]\(([^)\s]+)(?:\s+\"([^\"]*)\")?\)\s*$")
RE_TABLE_SEP = re.compile(r"^\s*\|?[\s:\-\|]+\|[\s:\-\|]*$")
RE_CALLOUT_OPEN = re.compile(r"^:::+\s*(\w+)?\s*(.*)$")
RE_CALLOUT_CLOSE = re.compile(r"^:::+\s*$")
RE_DIRECTIVE = re.compile(r"^\s*<!--\s*(pagebreak|newpage|toc|rule)\s*-->\s*$")
RE_CAPTION = re.compile(r"^:\s+(.*)$")

CALLOUT_KINDS = {"note", "tip", "warning", "example", "quote", "exercise", "definition", "key"}


@dataclass
class Block:
    t: str
    text: str = ""
    level: int = 0
    data: dict[str, Any] = field(default_factory=dict)


def _inline(text: str) -> list[tuple[str, str]]:
    """تفكيك التنسيق المضمّن إلى مقاطع (نمط، نص).

    الأنماط: b (عريض) و i (مائل) و c (كود) و '' (عادي)

    تُطبَّع الأرقام هنا (٠١٢ → 012) لأن هذه الدالة هي نقطة المرور الوحيدة لكل نصّ
    مضمّن في المخرجين (PDF وWord)؛ فتُطبَّق قاعدة «الأرقام لاتينية» على كل مسار
    تلقائياً، ولا يسقط نصّ من التطبيع إن نسي مُستدعٍ ما نداءه الصريح.
    """
    try:                       # استيراد محلي لتفادي دورة استيراد مع build_pdf
        import arabic as _A
        text = _A.normalize_numerals(text)
    except Exception:          # noqa: BLE001 - التطبيع تحسين لا شرط
        pass
    tokens: list[tuple[str, str]] = []
    pattern = re.compile(r"(\*\*|__)(.+?)\1|(\*|_)(.+?)\3|`([^`]+)`")
    pos = 0
    for m in pattern.finditer(text):
        if m.start() > pos:
            tokens.append(("", text[pos:m.start()]))
        if m.group(2) is not None:
            tokens.append(("b", m.group(2)))
        elif m.group(4) is not None:
            tokens.append(("i", m.group(4)))
        else:
            tokens.append(("c", m.group(5)))
        pos = m.end()
    if pos < len(text):
        tokens.append(("", text[pos:]))
    return tokens or [("", text)]


def parse_markdown(text: str, base_dir: str = "") -> list[Block]:
    lines = text.replace("\r\n", "\n").replace("\r", "\n").split("\n")
    out: list[Block] = []
    i = 0
    n = len(lines)

    while i < n:
        raw = lines[i]
        line = raw.rstrip()
        stripped = line.strip()

        # أسطر فارغة
        if not stripped:
            i += 1
            continue

        # توجيهات HTML المعلّقة
        m = RE_DIRECTIVE.match(stripped)
        if m:
            kind = m.group(1)
            out.append(Block({"pagebreak": "pagebreak", "newpage": "pagebreak",
                              "toc": "toc", "rule": "rule"}[kind]))
            i += 1
            continue

        # مقطع برمجي
        if stripped.startswith("```"):
            i += 1
            buf = []
            while i < n and not lines[i].strip().startswith("```"):
                buf.append(lines[i])
                i += 1
            i += 1
            out.append(Block("code", text="\n".join(buf).strip("\n")))
            continue

        # إطار (callout)
        m = RE_CALLOUT_OPEN.match(line)
        if m and m.group(1) in CALLOUT_KINDS:
            kind = m.group(1)
            title = (m.group(2) or "").strip()
            i += 1
            buf = []
            while i < n and not RE_CALLOUT_CLOSE.match(lines[i]):
                buf.append(lines[i])
                i += 1
            i += 1
            body = "\n".join(buf).strip("\n")
            out.append(Block("callout", text=body, data={"kind": kind, "title": title,
                                                         "blocks": parse_markdown(body, base_dir)}))
            continue

        # خط فاصل
        if re.match(r"^\s*([-*_])(?:\s*\1){2,}\s*$", line):
            out.append(Block("rule"))
            i += 1
            continue

        # عنوان
        m = RE_H.match(line)
        if m:
            level = min(len(m.group(1)), 4)
            out.append(Block("h", text=m.group(2).strip(), level=level))
            i += 1
            continue

        # صورة
        m = RE_IMG.match(stripped)
        if m:
            src = m.group(2)
            if base_dir and not os.path.isabs(src):
                src = os.path.normpath(os.path.join(base_dir, src))
            out.append(Block("figure", text=m.group(1) or "",
                             data={"src": src, "caption": m.group(3) or ""}))
            i += 1
            continue

        # تعليق (: نص) يلحق بالجدول أو الشكل السابق مباشرة
        m = RE_CAPTION.match(line)
        if m and out and out[-1].t in ("table", "figure"):
            out[-1].data["caption"] = m.group(1).strip()
            i += 1
            continue

        # جدول
        if stripped.startswith("|") and i + 1 < n and RE_TABLE_SEP.match(lines[i + 1]):
            header = [c.strip() for c in stripped.strip("|").split("|")]
            i += 2
            rows = []
            while i < n and lines[i].strip().startswith("|"):
                rows.append([c.strip() for c in lines[i].strip().strip("|").split("|")])
                i += 1
            out.append(Block("table", data={"header": header, "rows": rows}))
            continue

        # قوائم
        if RE_BULLET.match(line) or RE_ORDERED.match(line):
            ordered = bool(RE_ORDERED.match(line)) and not RE_BULLET.match(line)
            items: list[dict[str, Any]] = []
            while i < n:
                cur = lines[i]
                is_b = RE_BULLET.match(cur)
                is_o = RE_ORDERED.match(cur)
                if not (is_b or is_o):
                    if not cur.strip():
                        # سطر فارغ: نستمر فقط إن كان العنصر التالي أعمق
                        nxt = lines[i + 1] if i + 1 < n else ""
                        if RE_BULLET.match(nxt) or RE_ORDERED.match(nxt):
                            i += 1
                            continue
                    break
                mo = is_b or is_o
                assert mo is not None
                items.append({"text": mo.group(2).strip(), "level": len(mo.group(1)) // 2})
                i += 1
            out.append(Block("list", data={"ordered": ordered, "items": items}))
            continue

        # فقرة (تُجمع الأسطر المتتالية)
        buf = []
        while i < n:
            cur = lines[i].strip()
            if (not cur or RE_H.match(cur) or RE_IMG.match(cur) or RE_BULLET.match(cur)
                    or RE_ORDERED.match(cur) or RE_DIRECTIVE.match(cur)
                    or cur.startswith("|") or cur.startswith("```")
                    or (RE_CALLOUT_OPEN.match(cur) and RE_CALLOUT_OPEN.match(cur).group(1) in CALLOUT_KINDS)
                    or re.match(r"^\s*([-*_])(?:\s*\1){2,}\s*$", cur)):
                break
            buf.append(cur)
            i += 1
        if buf:
            out.append(Block("p", text=" ".join(buf)))
        else:
            i += 1

    return out


ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def load_blocks(path: str) -> list[Block]:
    """تحميل ملف مخطوطة مع إسناد احتياطي لمسارات الصور (مجلد الملف ← جذر الكتاب ← جذر المشروع)."""
    with open(path, encoding="utf-8") as f:
        blocks = parse_markdown(f.read(), base_dir=os.path.dirname(path))
    book_root = os.path.dirname(os.path.dirname(os.path.abspath(path)))
    for b in blocks:
        if b.t == "figure" and not os.path.exists(b.data.get("src", "")):
            name = os.path.basename(b.data["src"])
            for base in (book_root, ROOT):
                for sub in ("images", os.path.join("book", "images"), ""):
                    cand = os.path.join(base, sub, name)
                    if os.path.exists(cand):
                        b.data["src"] = cand
                        break
                if os.path.exists(b.data["src"]):
                    break
    return blocks


__all__ = ["Block", "parse_markdown", "load_blocks", "_inline"]
