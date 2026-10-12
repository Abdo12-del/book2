#!/usr/bin/env bash
# تهيئة بيئة العمل: تثبيت حزم بايثون اللازمة لبناء الكتاب (DOCX + PDF)
# يُستدعى مرة واحدة، أو تلقائياً إذا اكتشفنا نقص الحزم (قد لا تستمر الحزم بين الجلسات).
set -euo pipefail

# uharfbuzz اختياري لكنه مفيد: يمكّن محرّك ReportLab من التشكيل السياقي والربط
# (Ligatures) داخل الكلمة الواحدة، وإن غاب فالنصّ يبقى سليماً عبر الأشكال
# السياقية التي يطبّقها tools/textflow.py بنفسه.
PKGS="python-docx reportlab pymupdf pillow arabic-reshaper python-bidi pyyaml uharfbuzz"

python3 - <<'PY'
import importlib.util, sys
missing = [m for m in ("docx", "reportlab", "fitz", "PIL", "arabic_reshaper", "bidi", "yaml")
           if importlib.util.find_spec(m) is None]
if missing:
    print("ناقص:", ", ".join(missing))
    sys.exit(1)
print("✓ كل حزم بايثون مثبتة مسبقاً")
PY

if [ $? -ne 0 ]; then
  echo "→ تثبيت الحزم: $PKGS"
  python3 -m pip install --break-system-packages --quiet $PKGS
fi

python3 "$(dirname "$0")/fetch_fonts.py"
