# -*- coding: utf-8 -*-
"""
Xulosa Generator — генератор отчётов по мед. оборудованию
"""
import os
import sys
import re
import traceback

import pandas as pd
from docxtpl import DocxTemplate
from docx import Document
from docxcompose.composer import Composer

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from data import CITIES, ORGANIZATIONS, LAVOZIMLAR
from ttkwidgets.autocomplete import AutocompleteCombobox


# ---------- Утилиты ----------

UZ_MONTHS = {
    1: "yanvar", 2: "fevral", 3: "mart", 4: "aprel",
    5: "may", 6: "iyun", 7: "iyul", 8: "avgust",
    9: "sentabr", 10: "oktabr", 11: "noyabr", 12: "dekabr"
}


def parse_date(text):
    text = text.strip()
    for sep in (".", "-", "/"):
        if sep in text:
            parts = text.split(sep)
            if len(parts) == 3:
                d, m, y = parts
                return int(d), int(m), int(y)
    raise ValueError("Sana formati noto'g'ri. Misol: 11.06.2026")


def date_to_uz(dt_tuple):
    d, m, y = dt_tuple
    return f"{d} {UZ_MONTHS[m]} {y} yil"


def date_to_uz_bold(dt_tuple):
    d, m, y = dt_tuple
    return f"{y}-yil «{d}» {UZ_MONTHS[m]}"


def clean(val):
    if val is None:
        return "-"
    if isinstance(val, float) and pd.isna(val):
        return "-"
    s = str(val).strip()
    if s == "" or s.lower() == "nan":
        return "-"
    return s


def resource_path(relative):
    if hasattr(sys, "_MEIPASS"):
        return os.path.join(sys._MEIPASS, relative)
    return os.path.join(os.path.abspath("."), relative)


# ---------- Чтение Excel ----------

def read_equipment(excel_path):
    df = pd.read_excel(excel_path, header=None, dtype=object)

    header_row = None
    for i in range(min(20, len(df))):
        row_vals = [str(x).strip().lower() for x in df.iloc[i].tolist()]
        has_name = any(
            ("название" in v) or ("наимен" in v) or ("nomi" in v) or ("номи" in v)
            for v in row_vals
        )
        has_inv = any(
            ("инвентар" in v) or ("inventar" in v) or ("инв" in v)
            for v in row_vals
        )
        if has_name and has_inv:
            header_row = i
            break

    if header_row is None:
        for i in range(min(20, len(df))):
            row_vals = [str(x).strip().lower() for x in df.iloc[i].tolist()]
            if any(("название" in v) or ("наимен" in v) or ("nomi" in v) for v in row_vals):
                header_row = i
                break

    if header_row is None:
        header_row = 0

    header = [str(x).strip().lower() for x in df.iloc[header_row].tolist()]

    def find_col(*keys):
        for idx, h in enumerate(header):
            h_low = h.lower()
            for k in keys:
                if k.lower() in h_low:
                    return idx
        return None

    col_nomi = find_col(
        "название оборуд", "наименование оборуд", "наименование",
        "nomi", "номи", "apparat nomi", "аппарат номи",
        "название", "оборудование", "оборуд",
        "name", "equipment"
    )
    col_model = find_col("модель", "model", "мод")
    col_inv = find_col(
        "инвентарный номер", "инвентарный", "inventar nomer", "inventar",
        "инвентар", "инв", "inv"
    )
    col_year = find_col(
        "chiqarilgan yili", "год выпуска", "год производства",
        "дата производства", "дата выпуска",
        "chiqarilgan", "yili", "йили", "год",
        "year", "manufactured"
    )
    col_zavod = find_col(
        "заводской", "заводск", "zavod raqami", "zavod",
        "зав", "factory", "serial", "серийный"
    )
    col_masul = find_col(
        "материально ответственное", "ответственное лицо",
        "mas'ul shaxs", "mas'ul", "masul", "shaxs",
        "ответствен", "ответ", "фио", "fio", "масъул"
    )

    items = []
    for i in range(header_row + 1, len(df)):
        row = df.iloc[i].tolist()

        def get(idx):
            if idx is None or idx >= len(row):
                return None
            return row[idx]

        nomi = clean(get(col_nomi))
        inv = clean(get(col_inv))
        model = clean(get(col_model))
        zavod = clean(get(col_zavod))
        year = clean(get(col_year))
        masul = clean(get(col_masul))

        if nomi == "-" and inv == "-" and model == "-" and zavod == "-":
            continue
        if nomi.strip().lower() in (
            "итого", "jami", "total",
            "общее количество", "общее кол-во", "jami soni"
        ):
            continue

        items.append({
            "equipment_name": nomi,
            "equipment_model": model,
            "factory_number": zavod,
            "inventory_number": inv,
            "year_manufactured": year,
            "masul_from_excel": masul,
        })

    return items


# ---------- GUI ----------

class App:
    def __init__(self, root):
        self.root = root
        self.root.title("Xulosa Generator v1.0")
        self.root.geometry("1100x850")
        self.root.minsize(1000, 750)
        self.root.resizable(True, True)

        self.excel_path = None
        self.last_out_file = None

        ttk.Label(root, text="XULOSA GENERATOR",
                  font=("Arial", 16, "bold")).pack(pady=15)
        
        f0 = ttk.Frame(root); f0.pack(fill="x", padx=10, pady=6)
        ttk.Button(f0, text="1. Excel faylni tanlash",
                   command=self.select_excel, width=25).pack(side="left")
        self.lbl_excel = ttk.Label(f0, text="tanlanmagan", foreground="gray")
        self.lbl_excel.pack(side="left", padx=8)

        ttk.Separator(root).pack(fill="x", pady=8)

        self.fields = {}

        labels = [
            ("city",            "Shahar:",                     "combo", CITIES),
            ("inspection_date", "Sana (11.06.2026):",          "entry", None),
            ("org_name",        "Tashkilot nomi:",             "combo", ORGANIZATIONS),
            ("org_stir",        "STIR:",                       "entry", None),
            ("org_address",     "Tashkilot manzili:",          "entry", None),
            ("org_head",        "Tashkiloting rahbari:",       "entry", None),
            ("bosh_hisobchi",   "Bosh hisobchi:",              "entry", None),
            ("masul_shaxs",     "Mas'ul shaxs (bo'sh = Excel'dan):", "entry", None),
            ("lavozim",         "Lavozimi:",                   "combo", LAVOZIMLAR),
            ("fish",            "F.I.Sh.:",                    "entry", None),
        ]

        for key, lbl, kind, values in labels:
            fr = ttk.Frame(root); fr.pack(fill="x", padx=15, pady=4)
            ttk.Label(fr, text=lbl, width=32, font=("Arial", 11)).pack(side="left")

            if kind == "combo":
                cb = AutocompleteCombobox(fr, width=60, completevalues=values,
                                          font=("Arial", 11))
                cb.pack(side="left", fill="x", expand=True)
                self.fields[key] = cb
            else:
                e = ttk.Entry(fr, width=60, font=("Arial", 11))
                e.pack(side="left", fill="x", expand=True)
                self.fields[key] = e

        self.fields["city"].set("Nukus shahri")
        self.fields["lavozim"].set("Kichik muhandis")
        self.fields["fish"].insert(0, "B.K.Utemisov")

        ttk.Separator(root).pack(fill="x", pady=10)

        btn_frame = ttk.Frame(root)
        btn_frame.pack(pady=10)

        ttk.Button(btn_frame, text="▶  XULOSA YARATISH",
                   command=self.generate, width=30).pack(side="left", padx=8)

        ttk.Button(btn_frame, text="📂  Oxirgi faylni ochish",
                   command=self.open_last, width=30).pack(side="left", padx=8)

        self.log = tk.Text(root, height=12, width=100, font=("Consolas", 10))
        self.log.pack(padx=15, pady=10, fill="both", expand=True)

    def _bind_search(self, combobox, all_values):
        """Фильтрация списка при вводе."""
        def on_keyrelease(event):
            if event.keysym in ("Up", "Down", "Return", "Escape", "Tab",
                                "Left", "Right", "Home", "End"):
                return

            typed = combobox.get().strip().lower()

            if not typed:
                combobox["values"] = all_values
            else:
                filtered = [v for v in all_values if typed in v.lower()]
                combobox["values"] = filtered if filtered else all_values

        combobox.bind("<KeyRelease>", on_keyrelease)

        def on_focus(event):
            if not combobox.get():
                combobox["values"] = all_values

        combobox.bind("<FocusIn>", on_focus)

    def select_excel(self):
        """Открывает свой большой диалог выбора файла."""
        win = tk.Toplevel(self.root)
        win.title("Excel faylni tanlang")
        win.geometry("900x600")
        win.transient(self.root)
        win.grab_set()

        # Текущая папка
        current_dir = [os.path.expanduser("~")]

        # Верхняя панель с путём
        top = ttk.Frame(win)
        top.pack(fill="x", padx=10, pady=8)

        ttk.Label(top, text="Papka:", font=("Arial", 11)).pack(side="left")
        path_var = tk.StringVar(value=current_dir[0])
        path_entry = ttk.Entry(top, textvariable=path_var, font=("Arial", 11))
        path_entry.pack(side="left", fill="x", expand=True, padx=8)

        def go_up():
            current_dir[0] = os.path.dirname(current_dir[0])
            refresh()

        ttk.Button(top, text="⬆ Up", command=go_up).pack(side="left")

        # Список файлов и папок
        mid = ttk.Frame(win)
        mid.pack(fill="both", expand=True, padx=10, pady=5)

        scrollbar = ttk.Scrollbar(mid)
        scrollbar.pack(side="right", fill="y")

        listbox = tk.Listbox(mid, font=("Arial", 12),
                             yscrollcommand=scrollbar.set)
        listbox.pack(side="left", fill="both", expand=True)
        scrollbar.config(command=listbox.yview)

        # Нижняя панель — кнопки
        bottom = ttk.Frame(win)
        bottom.pack(fill="x", padx=10, pady=10)

        def refresh():
            listbox.delete(0, tk.END)
            try:
                items = sorted(os.listdir(current_dir[0]))
            except Exception:
                return

            # Папки сначала
            for it in items:
                full = os.path.join(current_dir[0], it)
                if os.path.isdir(full):
                    listbox.insert(tk.END, "📁 " + it)

            # Потом Excel-файлы
            for it in items:
                full = os.path.join(current_dir[0], it)
                if os.path.isfile(full) and it.lower().endswith((".xlsx", ".xls")):
                    listbox.insert(tk.END, "📄 " + it)

            path_var.set(current_dir[0])

        def on_double_click(event):
            sel = listbox.curselection()
            if not sel:
                return
            name = listbox.get(sel[0])
            name = name[2:]  # убираем "📁 " или "📄 "
            full = os.path.join(current_dir[0], name)
            if os.path.isdir(full):
                current_dir[0] = full
                refresh()
            else:
                self._choose_file(full, win)

        listbox.bind("<Double-Button-1>", on_double_click)

        def on_choose():
            sel = listbox.curselection()
            if not sel:
                return
            name = listbox.get(sel[0])[2:]
            full = os.path.join(current_dir[0], name)
            if os.path.isfile(full):
                self._choose_file(full, win)
            else:
                current_dir[0] = full
                refresh()

        ttk.Button(bottom, text="Tanlash",
                   command=on_choose, width=15).pack(side="right", padx=5)
        ttk.Button(bottom, text="Bekor qilish",
                   command=win.destroy, width=15).pack(side="right", padx=5)

        refresh()

    def _choose_file(self, full_path, win):
        """Сохраняет выбранный файл."""
        self.excel_path = full_path
        self.lbl_excel.config(
            text=os.path.basename(full_path),
            foreground="black"
        )
        win.destroy()

    def log_msg(self, msg):
        self.log.insert("end", msg + "\n")
        self.log.see("end")
        self.root.update()

    def open_last(self):
        if not self.last_out_file or not os.path.exists(self.last_out_file):
            messagebox.showinfo("Ma'lumot",
                "Hali xulosa yaratilmagan yoki fayl o'chirilgan.")
            return
        try:
            if sys.platform.startswith("win"):
                os.startfile(self.last_out_file)
            elif sys.platform == "darwin":
                os.system(f'open "{self.last_out_file}"')
            else:
                os.system(f'xdg-open "{self.last_out_file}"')
        except Exception as e:
            messagebox.showerror("Xato", str(e))

    def generate(self):
        if not self.excel_path:
            messagebox.showwarning("Diqqat", "Excel faylni tanlang!")
            return

        data = {k: v.get().strip() for k, v in self.fields.items()}

        if not data["inspection_date"]:
            messagebox.showwarning("Diqqat", "Sanani kiriting!")
            return

        try:
            dt = parse_date(data["inspection_date"])
        except Exception as e:
            messagebox.showerror("Xato", str(e))
            return

        data["inspection_date"] = date_to_uz(dt)
        data["inspection_date_bold"] = date_to_uz_bold(dt)

        try:
            self.log_msg(f"📖 Excel o'qilmoqda: {os.path.basename(self.excel_path)}")
            items = read_equipment(self.excel_path)
            if not items:
                messagebox.showerror("Xato", "Excel'da uskunalar topilmadi!")
                return
            self.log_msg(f"✅ Topildi: {len(items)} ta uskuna")

            template_path = resource_path("shablon.docx")
            if not os.path.exists(template_path):
                messagebox.showerror("Xato", f"Shablon topilmadi:\n{template_path}")
                return

            date_folder = f"Xulosalar_{dt[2]:04d}-{dt[1]:02d}-{dt[0]:02d}"
            base_dir = os.path.dirname(self.excel_path)
            out_dir = os.path.join(base_dir, date_folder)
            os.makedirs(out_dir, exist_ok=True)

            org_name = data["org_name"] or "Xulosa"
            safe_name = re.sub(r'[\\/*?:"<>|]', "_", org_name)
            out_file = os.path.join(out_dir, f"{safe_name}.docx")

            self.log_msg(f"📝 {len(items)} ta xulosa yaratilmoqda...")

            first = items[0]
            ctx = dict(data)
            ctx.update(first)
            ctx["masul_shaxs"] = data["masul_shaxs"] or first["masul_from_excel"]
            ctx["num"] = 1

            tpl = DocxTemplate(template_path)
            tpl.render(ctx)
            master_path = os.path.join(out_dir, "_master.docx")
            tpl.save(master_path)
            self.log_msg(f"   [1/{len(items)}] {first['equipment_name']}")

            master = Document(master_path)
            composer = Composer(master)

            for i, item in enumerate(items[1:], 2):
                ctx = dict(data)
                ctx.update(item)
                ctx["masul_shaxs"] = data["masul_shaxs"] or item["masul_from_excel"]
                ctx["num"] = i

                tpl = DocxTemplate(template_path)
                tpl.render(ctx)
                tmp_path = os.path.join(out_dir, f"_tmp_{i}.docx")
                tpl.save(tmp_path)

                doc_item = Document(tmp_path)
                composer.append(doc_item)
                os.remove(tmp_path)
                self.log_msg(f"   [{i}/{len(items)}] {item['equipment_name']}")

            composer.save(out_file)
            os.remove(master_path)

            self.last_out_file = out_file
            self.log_msg(f"✅ TAYYOR: {out_file}")
            messagebox.showinfo("Tayyor",
                f"{len(items)} ta xulosa yaratildi.\n\n{out_file}")

        except Exception as e:
            self.log_msg("❌ XATO: " + str(e))
            self.log_msg(traceback.format_exc())
            messagebox.showerror("Xato", str(e))


if __name__ == "__main__":
    root = tk.Tk()
    App(root)
    root.mainloop()