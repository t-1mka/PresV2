#!/bin/bash
# Rebuild the deck end to end: charts -> pptx -> pdf -> render -> audit.
set -o pipefail
cd "$(dirname "$0")"
python3 gen_charts.py >/dev/null || exit 1
python3 build_deck.py || exit 1
rm -f "Будущее_труда_в_автоматизированной_экономике.pdf"
LD_LIBRARY_PATH=/usr/lib/libreoffice/program timeout 300 soffice --headless --norestore \
  -env:UserInstallation=file:///tmp/lo_profile --convert-to pdf --outdir . \
  "Будущее_труда_в_автоматизированной_экономике.pptx" >/dev/null 2>&1 || exit 1
python3 verify_layout.py | tail -3
echo "== precise audit =="
python3 audit_pptx.py
