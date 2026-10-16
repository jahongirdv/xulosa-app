# -*- coding: utf-8 -*-
"""
Финальная правка shablon.docx:
- исправляет порядок строк
- заменяет inspection_date на inspection_date_bold в нужной фразе
- ставит жирный на плейсхолдер даты
"""
import re
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH


SRC = "shablon.docx"
DST = "shablon_final.docx"


def replace_in_paragraph(p, old, new):
    """Простая замена текста в параграфе — заменяет только первый run."""
    if old in p.text:
        # склеиваем весь текст
        full = "".join(r.text for r in p.runs)
        if old in full:
            new_full = full.replace(old, new)
            # сохраняем только первый run, остальным ставим пустоту
            if p.runs:
                p.runs[0].text = new_full
                for r in p.runs[1:]:
                    r.text = ""
            else:
                p.add_run(new_full)
            return True
    return False


def main():
    doc = Document(SRC)

    # ---------- 1. Замены в параграфах ----------
    replacements = [
        # Жирная дата в ключевой фразе
        ("Tekshiruv natijalari asosida {{ inspection_date }} sanasida",
         "Tekshiruv natijalari asosida {{ inspection_date_bold }} sanasida"),
        # Подпись эксперта — из шаблона
        ("Kichik muhandis: B.K.Utemisov",
         "Kichik muhandis: {{ fish }}"),
    ]

    for p in doc.paragraphs:
        for old, new in replacements:
            replace_in_paragraph(p, old, new)

    # ---------- 2. Жирный на inspection_date_bold ----------
    for p in doc.paragraphs:
        if "{{ inspection_date_bold }}" in p.text:
            # пересоберём параграф: обычный текст + жирный плейсхолдер
            full = p.text
            before, after = full.split("{{ inspection_date_bold }}", 1)

            # очищаем runs
            for r in list(p.runs):
                r._element.getparent().remove(r._element)

            p.add_run(before)
            bold_run = p.add_run("{{ inspection_date_bold }}")
            bold_run.bold = True
            p.add_run(after)

    # ---------- 3. Замены в таблицах ----------
    for t in doc.tables:
        for row in t.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    for old, new in replacements:
                        replace_in_paragraph(p, old, new)

    doc.save(DST)
    print(f"✅ Готово: {DST}")


if __name__ == "__main__":
    main()
