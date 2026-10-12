#!/usr/bin/env python3
"""إدارة سجل المراجع التراكمي: إضافة، تحقّق، مصفوفة تغطية، كشف فجوات.

    python3 tools/refs.py stats
    python3 tools/refs.py coverage
    python3 tools/refs.py add new-sources.yaml
    python3 tools/refs.py show S001
    python3 tools/refs.py gaps --part 2
"""
from __future__ import annotations

import argparse
import os
import sys

import yaml

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(ROOT, "tools"))

REGISTER = os.path.join(ROOT, "references", "register.yaml")
COVERAGE = os.path.join(ROOT, "references", "coverage.md")
PROPOSALS = os.path.join(ROOT, "references", "proposals.md")
FIHRIS = os.path.join(ROOT, "book", "fihris.txt")

from parse_fihris import parse  # noqa: E402

REQUIRED = ["id", "origin", "type", "title"]
VALID_ORIGIN = {"user", "agent"}
VALID_ACCESS = {"full", "partial", "metadata-only", "none"}
VALID_RELIABILITY = {"علمي محكّم", "مهني موثوق", "خبرة عملية", "رأي المؤلف"}


def load_register() -> list[dict]:
    if not os.path.exists(REGISTER):
        return []
    with open(REGISTER, encoding="utf-8") as f:
        data = yaml.safe_load(f) or {}
    return data.get("sources", []) if isinstance(data, dict) else (data or [])


def save_register(sources: list[dict]) -> None:
    with open(REGISTER, "w", encoding="utf-8") as f:
        yaml.safe_dump({"sources": sources}, f, allow_unicode=True,
                       sort_keys=False, width=100)


def structure():
    with open(FIHRIS, encoding="utf-8") as f:
        return parse(f.read())


def validate(sources: list[dict], doc: dict) -> list[str]:
    errs: list[str] = []
    chapters = {c["number"]: c for p in doc["parts"] for c in p["chapters"]}
    ids = set()
    for s in sources:
        sid = s.get("id", "?")
        for k in REQUIRED:
            if not s.get(k):
                errs.append(f"{sid}: الحقل الإلزامي «{k}» فارغ")
        if s.get("origin") not in VALID_ORIGIN:
            errs.append(f"{sid}: origin يجب أن يكون user أو agent، ورد: {s.get('origin')}")
        if s.get("access") and s["access"] not in VALID_ACCESS:
            errs.append(f"{sid}: access غير معروف: {s['access']}")
        if s.get("reliability") and s["reliability"] not in VALID_RELIABILITY:
            errs.append(f"{sid}: reliability غير معروف: {s['reliability']}")
        if sid in ids:
            errs.append(f"{sid}: معرّف مكرّر")
        ids.add(sid)
        for m in s.get("mapping", []) or []:
            ch = m.get("chapter")
            if ch not in chapters:
                errs.append(f"{sid}: رقم فصل غير موجود في الفهرس: {ch}")
                continue
            maxn = len(chapters[ch]["items"])
            for it in m.get("items", []) or []:
                if not (1 <= it <= maxn):
                    errs.append(f"{sid}: عنصر {it} خارج نطاق الفصل {ch} (١–{maxn})")
    return errs


def stats(sources: list[dict], doc: dict) -> None:
    chapters = [c for p in doc["parts"] for c in p["chapters"]]
    print("=== ملخّص سجل المراجع ===")
    print(f"  إجمالي المصادر: {len(sources)}")
    by_origin = {}
    for s in sources:
        by_origin[s.get("origin", "?")] = by_origin.get(s.get("origin", "?"), 0) + 1
    print(f"  ├─ منك (user): {by_origin.get('user', 0)}")
    print(f"  └─ ببحثي (agent): {by_origin.get('agent', 0)}")
    acc = {}
    for s in sources:
        acc[s.get("access", "?")] = acc.get(s.get("access", "?"), 0) + 1
    print("  مستوى الوصول: " + "، ".join(f"{k}={v}" for k, v in acc.items()))
    rel = {}
    for s in sources:
        rel[s.get("reliability", "?")] = rel.get(s.get("reliability", "?"), 0) + 1
    print("  الموثوقية: " + "، ".join(f"{k}={v}" for k, v in rel.items()))
    covered = {m["chapter"] for s in sources for m in (s.get("mapping") or [])}
    print(f"  فصول مغطّاة: {len(covered)}/{len(chapters)}")
    print(f"  فصول بلا أي مرجع: {[c['number'] for c in chapters if c['number'] not in covered]}")


COVER_ICON = {"none": "⬜ بلا مرجع", "weak": "� ضعيفة", "ok": "🟢 كافية", "good": "🔵 جيدة"}


def coverage(sources: list[dict], doc: dict, write: bool = True) -> str:
    lines = ["# مصفوفة التغطية: الفهرس ↔ المراجع",
             "",
             "> يُولَّد هذا الملف آليًا بـ `python3 tools/refs.py coverage`.",
             "> **⬜ بلا مرجع** = فجوة معرفية تحتاج بحثًا · "
             "**🟢 كافية** = مرجع واحد محكّم أو مرجعان مهنيان · "
             "**🔵 جيدة** = ثلاثة مصادر فأكثر أو مصدر محكّم + دليل تدريبي.",
             ""]
    by_ch: dict[int, list[dict]] = {}
    for s in sources:
        for m in s.get("mapping") or []:
            by_ch.setdefault(m["chapter"], []).append(s)

    for p in doc["parts"]:
        lines.append(f"## {p['raw']}")
        lines.append("")
        lines.append("| الفصل | العنوان | عناصر | مصادر | حالة التغطية | نوع المصادر |")
        lines.append("|---|---|---|---|---|---|")
        for c in p["chapters"]:
            ss = by_ch.get(c["number"], [])
            rel_hi = sum(1 for s in ss if s.get("reliability") == "علمي محكّم")
            n = len(ss)
            if n == 0:
                state = "none"
            elif n == 1 and rel_hi == 0:
                state = "weak"
            elif n >= 3 or (rel_hi >= 1 and n >= 2):
                state = "good"
            else:
                state = "ok"
            ids = "، ".join(s["id"] for s in ss) or "—"
            kinds = "، ".join(sorted({s.get("type", "؟") for s in ss})) or "—"
            lines.append(f"| {c['number']} | {c['title']} | {len(c['items'])} | {ids} | "
                         f"{COVER_ICON[state]} | {kinds} |")
        lines.append("")

    total = sum(len(c["items"]) for p in doc["parts"] for c in p["chapters"])
    gaps = [c["number"] for p in doc["parts"] for c in p["chapters"]
            if c["number"] not in by_ch]
    lines += ["---", "",
              f"**الإجمالي:** {len([c for p in doc['parts'] for c in p['chapters']])} فصلًا · "
              f"{total} عنصرًا · "
              f"{len(sources)} مصدرًا · **{len(gaps)} فجوة**",
              ""]
    if gaps:
        lines.append("**فجوات تحتاج بحثًا مستقلًا:** " + "، ".join(
            f"الفصل {g}" for g in gaps))
        lines.append("")
    out = "\n".join(lines)
    if write:
        with open(COVERAGE, "w", encoding="utf-8") as f:
            f.write(out)
        print(f"✓ {COVERAGE}")
    return out


def gaps(sources: list[dict], doc: dict, part: int | None = None) -> None:
    by_ch: dict[int, list[dict]] = {}
    for s in sources:
        for m in s.get("mapping") or []:
            by_ch.setdefault(m["chapter"], []).append(s)
    print("=== الفجوات المعرفية ===")
    for p in doc["parts"]:
        if part and p["chapters"] and part != doc["parts"].index(p) + 1:
            continue
        miss = [c for c in p["chapters"] if c["number"] not in by_ch]
        weak = [c for c in p["chapters"]
                if len(by_ch.get(c["number"], [])) == 1
                and not any(s.get("reliability") == "علمي محكّم"
                            for s in by_ch.get(c["number"], []))]
        print(f"\n{p['raw']}")
        print(f"  بلا مرجع : {[c['number'] for c in miss] or '—'}")
        print(f"  تغطية ضعيفة: {[c['number'] for c in weak] or '—'}")


def add(path: str, doc: dict) -> int:
    with open(path, encoding="utf-8") as f:
        new = yaml.safe_load(f) or {}
    items = new.get("sources", []) if isinstance(new, dict) else (new or [])
    sources = load_register()
    start = 1 + max([int(str(s.get("id", "S000")).lstrip("S") or 0) for s in sources] or [0])
    for i, s in enumerate(items):
        if not s.get("id"):
            s["id"] = f"S{start + i:03d}"
    errs = validate(items, doc)
    if errs:
        print("✗ أخطاء التحقق:")
        for e in errs:
            print("   " + e)
        return 1
    sources += items
    save_register(sources)
    print(f"✓ أُضيف {len(items)} مصدرًا (الإجمالي الآن: {len(sources)})")
    coverage(sources, doc)
    return 0


def main() -> int:
    ap = argparse.ArgumentParser(description="إدارة سجل المراجع")
    ap.add_argument("cmd", choices=["stats", "coverage", "add", "show", "gaps", "check"])
    ap.add_argument("arg", nargs="?")
    ap.add_argument("--part", type=int)
    a = ap.parse_args()

    doc = structure()
    sources = load_register()

    if a.cmd == "stats":
        stats(sources, doc)
    elif a.cmd == "coverage":
        coverage(sources, doc)
    elif a.cmd == "add":
        if not a.arg:
            print("حدد ملف YAML")
            return 1
        return add(a.arg, doc)
    elif a.cmd == "show":
        for s in sources:
            if s["id"] == a.arg:
                print(yaml.safe_dump(s, allow_unicode=True, sort_keys=False, width=100))
                break
        else:
            print("غير موجود:", a.arg)
            return 1
    elif a.cmd == "gaps":
        gaps(sources, doc, a.part)
    elif a.cmd == "check":
        errs = validate(sources, doc)
        print("✓ السجل سليم" if not errs else "✗ أخطاء:\n  " + "\n  ".join(errs))
        return 1 if errs else 0
    return 0


if __name__ == "__main__":
    sys.exit(main())
