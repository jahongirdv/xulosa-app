#!/bin/bash
cd "$(dirname "$0")"
source venv/bin/activate

pyinstaller \
  --onefile \
  --windowed \
  --name "XulosaGenerator" \
  --icon "icon.ico" \               # ⬅️ ДОБАВИТЬ ЭТУ СТРОКУ
  --add-data "shablon.docx:." \
  --add-data "images.ico:." \         # ⬅️ И ЭТУ (чтобы иконка попала в .exe)
  main.py

echo ""
echo "================================================"
echo "  Версия готова: dist/XulosaGenerator"
echo "================================================"