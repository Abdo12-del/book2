#!/usr/bin/env bash
# تهيئة بيئة العمل: تثبيت حزم بايثون اللازمة لبناء الكتاب (DOCX + PDF)
# يُستدعى مرة واحدة، أو تلقائياً إذا اكتشفنا نقص الحزم (قد لا تستمر الحزم بين الجلسات).
set -euo pipefail

PKGS="python-docx reportlab pymupdf pillow arabic-reshaper python-bidi pyyaml"

python3 - <<PY || true
import importlib, sys
missing = [m for m in ["docx","reportlab","fitz","PIL","arabic_reshaper","bidi","yaml"]
           if not importlib.util.find_spec(m)]
sys.exit(1 if missing else 0)
PY

if [ $? -eq 0 ]; then
  echo "✓ كل حزم بايثون مثبتة مسبقاً"
else
  echo "→ تثبيت الحزم: $PKGS"
  python3 -m pip install --break-system-packages --quiet $PKGS
fi

python3 "$(dirname "$0")/fetch_fonts.py"
