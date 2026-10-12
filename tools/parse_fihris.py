#!/usr/bin/env python3
"""تحليل الفهرس المعتمد: تحويله إلى خريطة بنيوية قابلة للقراءة الآلية + تقرير تحقّق.

هذه الأداة **لا تعدّل الفهرس**؛ تحوّله كما هو إلى خريطة، ثم تفحص اتساقه:
  • تسلسل أرقام الفصول (١…٢٠) ومناسبتها لأسمائها العربية
  • تسلسل أرقام العناصر الفرعية داخل كل فصل (١،٢،٣… بلا قفز أو تكرار)
  • العناصر المكرّرة أو المتقاربة بين الفصول (تُعرض للمراجعة، لا تُدمج تلقائياً)
  • توافق «٨ محاور في كل مستوى = ١٦ محوراً» مع الوارد فعلياً

الاستخدام:
    python3 tools/parse_fihris.py book/fihris.txt docs/structure-map.yaml
"""
from __future__ import annotations

import argparse
import json
import os
import re
import sys
from collections import Counter

import yaml

ORDINALS = {
    "الأول": 1, "الثاني": 2, "الثالث": 3, "الرابع": 4, "الخامس": 5, "السادس": 6,
    "السابع": 7, "الثامن": 8, "التاسع": 9, "العاشر": 10,
    "الحادي عشر": 11, "الحادى عشر": 11, "الحادي والعشرون": 21,
    "الثاني عشر": 12, "الثانى عشر": 12, "الثالث عشر": 13, "الرابع عشر": 14,
    "الخامس عشر": 15, "السادس عشر": 16, "السابع عشر": 17, "الثامن عشر": 18,
    "التاسع عشر": 19, "العشرون": 20, "الحادي والعشرین": 21,
}

RE_PART = re.compile(r"^الباب\s+(.+?)\s*[:：]\s*(.+)$")
RE_CHAPTER = re.compile(r"^الفصل\s+(.+?)\s*[:：]\s*(.+)$")
RE_ITEM = re.compile(r"^(\d+)\.\s*(.+)$")
RE_APPENDIX = re.compile(r"^[-–]\s*(.+)$")


def norm(s: str) -> str:
    """تطبيع خفيف للمقارنة (إزالة التشكيل والتنوين وعلامات الترقيم)."""
    s = re.sub(r"[ً-ْٰـ]", "", s)
    s = s.replace("أ", "ا").replace("إ", "ا").replace("آ", "ا").replace("ى", "ي")
    s = s.replace("ة", "ه").replace("ؤ", "و").replace("ئ", "ي")
    s = re.sub(r"[^\w\s]", "", s)
    return re.sub(r"\s+", " ", s).strip()


def parse(text: str) -> dict:
    lines = [l.strip() for l in text.replace("\r\n", "\n").split("\n")]
    doc: dict = {"parts": [], "appendices": [], "notes": []}
    cur_part = cur_ch = None
    header: list[str] = []
    in_appendix = False
    pending_intro: list[str] = []

    for ln in lines:
        if not ln:
            continue
        m = RE_PART.match(ln)
        if m:
            in_appendix = False
            cur_part = {"ordinal_text": m.group(1).strip(), "title": m.group(2).strip(),
                        "chapters": [], "intro": None, "raw": ln}
            doc["parts"].append(cur_part)
            cur_ch = None
            continue
        m = RE_CHAPTER.match(ln)
        if m:
            in_appendix = False
            word = m.group(1).strip()
            num = ORDINALS.get(word)
            cur_ch = {"number": num, "ordinal_text": word, "title": m.group(2).strip(),
                      "items": [], "raw": ln}
            if cur_part:
                cur_part["chapters"].append(cur_ch)
            else:
                doc.setdefault("orphan_chapters", []).append(cur_ch)
            continue
        m = RE_ITEM.match(ln)
        if m and cur_ch is not None and not in_appendix:
            cur_ch["items"].append({"n": int(m.group(1)),
                                    "text": m.group(2).strip().rstrip(".")})
            continue
        if ln.startswith("الملاحق"):
            in_appendix = True
            doc["appendices_title"] = ln
            cur_ch = None
            continue
        m = RE_APPENDIX.match(ln)
        if m and in_appendix:
            doc["appendices"].append(m.group(1).strip().rstrip("."))
            continue
        # سطر نصّي: عنوان/وصف/تمهيد
        if cur_part and cur_ch is None and not in_appendix:
            cur_part["intro"] = (cur_part["intro"] + " " if cur_part.get("intro") else "") + ln
        elif cur_ch is not None and not in_appendix:
            cur_ch.setdefault("intro", (cur_ch.get("intro", "") + " " + ln).strip())
        elif not cur_part:
            header.append(ln)

    # ترويسة المستند: العنوان / العنوان الفرعي / الوصف / تمهيد
    doc["header"] = header
    if len(header) >= 1:
        doc["title"] = header[0]
    if len(header) >= 2:
        doc["subtitle"] = header[1]
    if len(header) >= 3:
        doc["description"] = header[2]
    if len(header) >= 4:
        doc["preamble"] = " ".join(header[3:])
    return doc


def validate(doc: dict) -> list[str]:
    issues: list[str] = []
    chapters = [c for p in doc["parts"] for c in p["chapters"]]

    # 1) تسلسل الفصول
    nums = [c["number"] for c in chapters]
    if None in nums:
        bad = [c["ordinal_text"] for c in chapters if c["number"] is None]
        issues.append(f"تعذّر تحويل رقم الفصل إلى رقم صحيح: {bad}")
    else:
        expected = list(range(1, len(chapters) + 1))
        if nums != expected:
            issues.append(f"تسلسل أرقام الفصول غير متصل: {nums}")
    print(f"  عدد الأبواب: {len(doc['parts'])} | عدد الفصول: {len(chapters)} | "
          f"عدد العناصر الفرعية: {sum(len(c['items']) for c in chapters)} | "
          f"عدد الملاحق: {len(doc['appendices'])}")

    # 2) تسلسل العناصر داخل كل فصل
    for c in chapters:
        ns = [i["n"] for i in c["items"]]
        if ns != list(range(1, len(ns) + 1)):
            issues.append(f"ترقيم غير متسلسل في «{c['title']}»: {ns}")

    # 3) التكرار والتقارب بين العناوين
    seen: dict[str, list[str]] = {}
    for c in chapters:
        for i in c["items"]:
            seen.setdefault(norm(i["text"]), []).append(f"{c['number']}-{i['n']}")
    dups = {k: v for k, v in seen.items() if len(v) > 1}
    if dups:
        issues.append("عناوين مكرّرة حرفياً: " + "; ".join(f"{k} ← {v}" for k, v in dups.items()))

    # تقارب في الكلمات المفتاحية (اشتباه تداخل لا تكرار تام)
    keys = ["أدوات الربط", "التقويم المرحلي", "تقويم", "التطبيقات", "التعبير عن", "المفردات"]
    for k in keys:
        hits = [(c["number"], i["n"], i["text"]) for c in chapters for i in c["items"]
                if k in i["text"]]
        if len(hits) > 2:
            issues.append(f"تكرار المفتاح «{k}» في {len(hits)} موضعًا: "
                          + " | ".join(f"{a}-{b}" for a, b, _ in hits))

    # 4) توافق «8 محاور لكل مستوى»
    level_parts = [p for p in doc["parts"] if "المستوى" in p["title"]]
    for p in level_parts:
        if len(p["chapters"]) != 8:
            issues.append(f"«{p['title']}» يحتوي {len(p['chapters'])} فصول (المُعلَن: 8)")

    # 5) فصول بلا عنصر «تقويم» (ملاحظة تنظيمية، لا تُعدّ خطأ)
    no_eval = [c["number"] for c in chapters
               if not any("تقويم" in i["text"] or "تقييم" in i["text"] for i in c["items"])]
    print(f"  فصول لا ينتهي عناصرها بـ«تقويم»: {no_eval}")
    return issues


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("src", nargs="?", default="book/fihris.txt")
    ap.add_argument("out", nargs="?", default="docs/structure-map.yaml")
    a = ap.parse_args()

    with open(a.src, encoding="utf-8") as f:
        doc = parse(f.read())

    print("=== ملخّص البنية ===")
    issues = validate(doc)

    chapters = [c for p in doc["parts"] for c in p["chapters"]]
    doc["stats"] = {
        "parts": len(doc["parts"]),
        "chapters": len(chapters),
        "items": sum(len(c["items"]) for c in chapters),
        "appendices": len(doc["appendices"]),
        "items_per_chapter": {c["number"]: len(c["items"]) for c in chapters},
    }
    doc["validation_issues"] = issues

    os.makedirs(os.path.dirname(a.out) or ".", exist_ok=True)
    with open(a.out, "w", encoding="utf-8") as f:
        yaml.safe_dump(doc, f, allow_unicode=True, sort_keys=False, width=100)

    print(f"\n=== الخريطة البنيوية → {a.out} ===")
    for p in doc["parts"]:
        print(f"\n{p['raw']}  ({len(p['chapters'])} فصول، "
              f"{sum(len(c['items']) for c in p['chapters'])} عنصراً)")
        for c in p["chapters"]:
            print(f"   الفصل {c['number']:>2}: {c['title'][:58]:<58} ({len(c['items'])} عنصراً)")
    print(f"\nالملاحق ({len(doc['appendices'])}):")
    for i, x in enumerate(doc["appendices"], 1):
        print(f"   {i}. {x}")

    print("\n=== ملاحظات التحقق ===")
    if issues:
        for i in issues:
            print("  ! " + i)
    else:
        print("  ✓ لا ملاحظات")
    return 0


if __name__ == "__main__":
    sys.exit(main())
