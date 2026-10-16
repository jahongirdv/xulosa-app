#!/bin/bash
cd "$(dirname "$0")"
source venv/bin/activate

pyinstaller \
  --onefile \
  --windowed \
  --name "XulosaGenerator" \
  --add-data "shablon.docx:." \
  main.py

echo ""
echo "================================================"
echo "  Linux-версия готова: dist/XulosaGenerator"
echo "================================================"