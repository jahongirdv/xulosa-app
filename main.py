# -*- coding: utf-8 -*-
"""
Xulosa & Dalolatnoma Generator
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
UZ_MONTHS = {1: "yanvar", 2: "fevral", 3: "mart", 4: "aprel", 5: "may", 6: "iyun", 7: "iyul", 8: "avgust", 9: "sentabr", 10: "oktabr", 11: "noyabr", 12: "dekabr"}

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
    if val is None: return "-"
    if isinstance(val, float) and pd.isna(val): return "-"
    s = str(val).strip()
    if s == "" or s.lower() == "nan": return "-"
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
        has_name = any(("название" in v) or ("наимен" in v) or ("nomi" in v) for v in row_vals)
        has_inv = any(("инвентар" in v) or ("inventar" in v) for v in row_vals)
        if has_name and has_inv:
            header_row = i
            break
    if header_row is None: header_row = 0
    header = [str(x).strip().lower() for x in df.iloc[header_row].tolist()]

    def find_col(*keys):
        for idx, h in enumerate(header):
            h_low = h.lower()
            for k in keys:
                if k.lower() in h_low: return idx
        return None

    col_nomi = find_col(
        "название", "наименование", "nomi", "номи", "оборудование",
        "аппаратура атамасы", "атамасы", "аппарат", "жихоз", "жихоз номи"
    )
    col_model = find_col(
        "модель", "model", "мод", "марка", "маркасы", "маркасы/модель"
    )
    col_inv = find_col(
        "инвентарный", "inventar", "инв", "инвентар номери", "инв №", "инвентарный номер", "инвентар", "инвентар номери таза"
    )
    col_year = find_col(
        "chiqarilgan", "год", "year", "нешинши жыл келинген", "келинген жылы", "жылы", "жыл", "год выпуска"
    )
    col_zavod = find_col(
        "заводской", "zavod", "serial", "серия", "серия номери", "заводской номер", "серийный номер", "seriya nomeri", "заводской номер (серийный номер)", "заводской номер (серийный)", "заводской №", "заводской № (серийный)", "заводской № (серийный номер)"
    )
    col_masul = find_col(
        "материально", "mas'ul", "masul", "ответствен", "ответ", "ст м/с", "ст м/с (ответственное лицо)", "ответственное лицо", "фио", "фио (ответственное лицо)", "фио материально ответственного лица", "фио материально-ответственного лица", "фио материально-ответственного", "фио материально ответственного",
    )

    items = []
    for i in range(header_row + 1, len(df)):
        row = df.iloc[i].tolist()
        def get(idx): return row[idx] if idx is not None and idx < len(row) else None

        nomi = clean(get(col_nomi))
        if nomi == "-" or nomi.lower() in ("итого", "jami", "total"): continue

        items.append({
            "equipment_name": nomi, "equipment_model": clean(get(col_model)),
            "factory_number": clean(get(col_zavod)), "inventory_number": clean(get(col_inv)),
            "year_manufactured": clean(get(col_year)), "masul_from_excel": clean(get(col_masul)),
        })
    return items

# ---------- Генерация ----------
def generate_combined_docs(items, template_path, out_file, data, mode, log_func):
    master = None
    composer = None

    show_xulosa = mode in ("xulosa", "both")
    show_dalolatnoma = mode in ("dalolatnoma", "both")

    for i, item in enumerate(items, 1):
        ctx = dict(data)
        ctx.update(item)
        ctx["num"] = i
        ctx["masul_shaxs"] = data["masul_shaxs"] or item.get("masul_from_excel", "")
        ctx["region_manufactured"] = item.get("region_manufactured", "")
        ctx["org_head_lavozim"] = data.get("org_head_lavozim", "")
        ctx["masul_shaxs_lavozim"] = data.get("masul_shaxs_lavozim", "")
        
        ctx["org_stir"] = data.get("org_stir", "")
        ctx["org_str"] = data.get("org_stir", "") 
        
        ctx["show_xulosa"] = show_xulosa
        ctx["show_dalolatnoma"] = show_dalolatnoma

        tpl = DocxTemplate(template_path)
        tpl.render(ctx)
        
        tmp_file = f"_tmp_{i}.docx"
        tpl.save(tmp_file)
        
        doc = Document(tmp_file)
        os.remove(tmp_file)

        from docx.oxml.ns import qn
        for p in doc.paragraphs:
            if not p.text.strip():
                p_element = p._element
                p_element.getparent().remove(p_element)

        doc.add_page_break()

        if master is None:
            master = doc
            composer = Composer(master)
        else:
            composer.append(doc)
            
        log_func(f"   [{i}/{len(items)}] {item['equipment_name']} - tayyor")

    composer.save(out_file)

# ---------- Словари переводов ----------
TEXTS = {
    "uz": {
        "title": "XULOSA VA DALOLATNOMA GENERATOR",
        "btn_excel": "📁 Excel faylni tanlash",
        "no_file": "fayl tanlanmagan",
        "lbl_mode": "Nima yaratilsin?",
        "btn_xulosa": "📄 Faqat Xulosa",
        "btn_dalolatnoma": "📑 Faqat Dalolatnoma",
        "btn_both": "📚 Ikkalasi ham",
        "frame_main": " Asosiy ma'lumotlar ",
        "frame_dalolatnoma": " Nuqson dalolatnomasi uchun ",
        "btn_generate": "▶  YARATISH",
        "btn_open": "📂  Oxirgi faylni ochish"
    },
    "ru": {
        "title": "ГЕНЕРАТОР XULOSA И DALOLATNOMA",
        "btn_excel": "📁 Выбрать Excel файл",
        "no_file": "файл не выбран",
        "lbl_mode": "Что сгенерировать?",
        "btn_xulosa": "📄 Только Xulosa",
        "btn_dalolatnoma": "📑 Только Dalolatnoma",
        "btn_both": "📚 Оба документа",
        "frame_main": " Основные данные ",
        "frame_dalolatnoma": " Для Dalolatnoma ",
        "btn_generate": "▶  СГЕНЕРИРОВАТЬ",
        "btn_open": "📂  Открыть посл. файл"
    }
}

# ---------- GUI ----------
class App:
    def __init__(self, root):
        self.root = root
        self.root.title("Xulosa Generator")
        
        # 🔥 УМЕНЬШАЕМ РАЗМЕР И РАЗВОРАЧИВАЕМ НА ВЕСЬ ЭКРАН
        self.root.geometry("1100x850")
        self.root.minsize(900, 700)


        style = ttk.Style()
        style.configure(".", font=("Arial", 10))
        style.configure("TButton", padding=5)
        style.configure("TLabelframe", font=("Arial", 10, "bold"))
        style.configure("Action.TButton", font=("Arial", 11, "bold"), padding=8)

        self.current_lang = "uz"
        self.ui_elements = {}

        self.excel_path = None
        self.last_out_file = None
        self.gen_mode = tk.StringVar(value="both")

        # --- ПЕРЕКЛЮЧАТЕЛЬ ЯЗЫКА ---
        lang_frame = ttk.Frame(root)
        lang_frame.pack(fill="x", padx=15, pady=5)
        
        self.lang_var = tk.StringVar(value="uz")
        ttk.Radiobutton(lang_frame, text="O'zbekcha", variable=self.lang_var, value="uz", command=lambda: self.change_lang("uz")).pack(side="right", padx=5)
        ttk.Radiobutton(lang_frame, text="Русский", variable=self.lang_var, value="ru", command=lambda: self.change_lang("ru")).pack(side="right", padx=5)
        ttk.Label(lang_frame, text="Til / Язык:", font=("Arial", 10, "bold")).pack(side="right", padx=5)

        # --- ЗАГОЛОВОК ---
        self.lbl_title = ttk.Label(root, text="", font=("Arial", 16, "bold"))
        self.lbl_title.pack(pady=5)

        ttk.Separator(root).pack(fill="x", pady=5)

        # --- ВЫБОР EXCEL ---
        f0 = ttk.Frame(root, padding=10)
        f0.pack(fill="x")
        self.btn_excel = ttk.Button(f0, text="", command=self.select_excel, width=25)
        self.btn_excel.pack(side="left")
        self.lbl_excel = ttk.Label(f0, text="", foreground="#666")
        self.lbl_excel.pack(side="left", padx=15)

        # --- РЕЖИМЫ ---
        self.frame_mode = ttk.LabelFrame(root, text="", padding=10)
        self.frame_mode.pack(fill="x", padx=15, pady=10)
        
        style.configure("Xulosa.TButton", background="#d4edda", foreground="#155724")
        style.configure("Dalolatnoma.TButton", background="#cce5ff", foreground="#004085")
        style.configure("Both.TButton", background="#fff3cd", foreground="#856404")

        def set_mode(mode):
            self.gen_mode.set(mode)
            self.btn_xulosa.configure(style="TButton")
            self.btn_dalolatnoma.configure(style="TButton")
            self.btn_both.configure(style="TButton")
            if mode == "xulosa": self.btn_xulosa.configure(style="Xulosa.TButton")
            elif mode == "dalolatnoma": self.btn_dalolatnoma.configure(style="Dalolatnoma.TButton")
            elif mode == "both": self.btn_both.configure(style="Both.TButton")

        self.btn_xulosa = ttk.Button(self.frame_mode, text="", command=lambda: set_mode("xulosa"))
        self.btn_xulosa.pack(side="left", padx=10)
        self.btn_dalolatnoma = ttk.Button(self.frame_mode, text="", command=lambda: set_mode("dalolatnoma"))
        self.btn_dalolatnoma.pack(side="left", padx=10)
        self.btn_both = ttk.Button(self.frame_mode, text="", command=lambda: set_mode("both"))
        self.btn_both.pack(side="left", padx=10)
        set_mode("both")

        self.fields = {}

        # --- БЛОК 1 ---
        self.frame_main = ttk.LabelFrame(root, text="", padding=10)
        self.frame_main.pack(fill="x", padx=15, pady=5)
        labels_xulosa = [
            ("city", "Shahar:" if self.current_lang=="uz" else "Город:", "combo", CITIES),
            ("inspection_date", "Sana:" if self.current_lang=="uz" else "Дата:", "entry", None),
            ("org_name", "Tashkilot nomi:" if self.current_lang=="uz" else "Название организации:", "combo", ORGANIZATIONS),
            ("org_stir", "STIR:", "entry", None),
            ("org_address", "Tashkilot manzili:" if self.current_lang=="uz" else "Адрес организации:", "entry", None),
            ("org_head", "Tashkiloting rahbari:" if self.current_lang=="uz" else "Руководитель:", "entry", None),
            ("bosh_hisobchi", "Bosh hisobchi:" if self.current_lang=="uz" else "Главный бухгалтер:", "entry", None),
            ("masul_shaxs", "Mas'ul shaxs:" if self.current_lang=="uz" else "Ответственное лицо:", "entry", None),
            ("fish", "Mutaxassis F.I.Sh.:" if self.current_lang=="uz" else "Ф.И.О. эксперта:", "entry", None),
            ("lavozim", "Mutaxassis lavozimi:" if self.current_lang=="uz" else "Должность эксперта:", "combo", LAVOZIMLAR),
        ]
        self._build_fields(self.frame_main, labels_xulosa)

        # --- БЛОК 2 ---
        self.frame_dalolatnoma = ttk.LabelFrame(root, text="", padding=10)
        self.frame_dalolatnoma.pack(fill="x", padx=15, pady=5)
        labels_dalolatnoma = [
            ("org_head_lavozim", "Rahbar lavozimi:" if self.current_lang=="uz" else "Должность руководителя:", "combo", ["Bosh shifokori", "Direktor"]),
            ("masul_shaxs_lavozim", "Mas'ul shaxs lavozimi:" if self.current_lang=="uz" else "Должность отв. лица:", "combo", ["texnik", "buxgalter", "muhandis"]),
            ("kafolat", "Kafolat muddati:" if self.current_lang=="uz" else "Гарантийный срок:", "combo", ["amal qilmoqda", "tugagan"]),
            ("tamirlash", "Uskuna ta'mirlashga:" if self.current_lang=="uz" else "Оборудование к ремонту:", "combo", ["yaroqli", "yaroqsiz"]),
            ("metrologik", "Metrologik tekshiruv:" if self.current_lang=="uz" else "Метрологическая проверка:", "combo", ["qayta o'tkazish zarur", "qayta o'tkazishga zarurat yo'q"]),
            ("foydalanish", "Foydalanishni davom ettirish:" if self.current_lang=="uz" else "Дальнейшая эксплуатация:", "combo", ["ha (mumkin)", "yo'q (mumkin emas)"]),
        ]
        self._build_fields(self.frame_dalolatnoma, labels_dalolatnoma)

        self.fields["city"].set("Nukus shahri")
        self.fields["org_head_lavozim"].set("Bosh shifokori")
        self.fields["masul_shaxs_lavozim"].set("texnik")
        self.fields["lavozim"].set("Kichik muhandis")
        self.fields["fish"].insert(0, "B.K.Utemisov")

        # 🔥 НИЖНЯЯ ПАНЕЛЬ (прикреплена к низу)
        btn_frame = ttk.Frame(root, padding=10)
        btn_frame.pack(side="bottom", fill="x") # Кнопки всегда внизу
        
        self.btn_generate = ttk.Button(btn_frame, text="", style="Action.TButton", command=self.generate, width=30)
        self.btn_generate.pack(side="left", padx=10)
        
        self.btn_open = ttk.Button(btn_frame, text="", command=self.open_last, width=25)
        self.btn_open.pack(side="left", padx=10)

        # Лог-окно (занимает оставшееся место)
        self.log = tk.Text(root, height=5, width=100, font=("Consolas", 9), bg="#f9f9f9", relief="solid", bd=1)
        self.log.pack(side="bottom", padx=15, pady=(0, 10), fill="x")

        # Применяем перевод при запуске
        self.update_ui_texts()

    def _build_fields(self, parent_frame, fields_config):
        for key, lbl, kind, values in fields_config:
            fr = ttk.Frame(parent_frame); fr.pack(fill="x", pady=3)
            lbl_widget = ttk.Label(fr, text=lbl, width=30)
            lbl_widget.pack(side="left")
            
            if kind == "combo":
                cb = AutocompleteCombobox(fr, width=60, completevalues=values, font=("Arial", 10))
                cb.pack(side="left", fill="x", expand=True)
                self.fields[key] = cb
            else:
                e = ttk.Entry(fr, width=60, font=("Arial", 10))
                e.pack(side="left", fill="x", expand=True)
                self.fields[key] = e

    def change_lang(self, lang):
        self.current_lang = lang
        self.update_ui_texts()

    def update_ui_texts(self):
        t = TEXTS[self.current_lang]
        self.lbl_title.config(text=t["title"])
        self.btn_excel.config(text=t["btn_excel"])
        if not self.excel_path:
            self.lbl_excel.config(text=t["no_file"])
        
        self.frame_mode.config(text=t["lbl_mode"])
        self.btn_xulosa.config(text=t["btn_xulosa"])
        self.btn_dalolatnoma.config(text=t["btn_dalolatnoma"])
        self.btn_both.config(text=t["btn_both"])
        
        self.frame_main.config(text=t["frame_main"])
        self.frame_dalolatnoma.config(text=t["frame_dalolatnoma"])
        
        self.btn_generate.config(text=t["btn_generate"])
        self.btn_open.config(text=t["btn_open"])

    def select_excel(self):
        # 🔥 ОТКРЫВАЕМ СТАНДАРТНЫЙ ПРОВОДНИК WINDOWS
        file_path = filedialog.askopenfilename(
            title="Excel faylni tanlang",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        if file_path:
            self.excel_path = file_path
            self.lbl_excel.config(text=os.path.basename(file_path), foreground="black")

    def _choose_file(self, full_path, win):
        self.excel_path = full_path
        self.lbl_excel.config(text=os.path.basename(full_path), foreground="black")
        win.destroy()

    def log_msg(self, msg):
        self.log.insert("end", msg + "\n"); self.log.see("end"); self.root.update()

    def open_last(self):
        if not self.last_out_file or not os.path.exists(self.last_out_file):
            messagebox.showinfo("Info", "Fayl topilmadi.")
            return
        try:
            if sys.platform.startswith("win"): os.startfile(self.last_out_file)
            elif sys.platform == "darwin": os.system(f'open "{self.last_out_file}"')
            else: os.system(f'xdg-open "{self.last_out_file}"')
        except Exception as e: messagebox.showerror("Error", str(e))

    def generate(self):
        if not self.excel_path: return messagebox.showwarning("Warning", "Excel tanlang!")
        data = {k: v.get().strip() for k, v in self.fields.items()}
        if not data["inspection_date"]: return messagebox.showwarning("Warning", "Sanani kiriting!")
        try: dt = parse_date(data["inspection_date"])
        except Exception as e: return messagebox.showerror("Error", str(e))

        data["inspection_date"] = date_to_uz(dt)
        data["inspection_date_bold"] = date_to_uz_bold(dt)
        mode = self.gen_mode.get()

        try:
            items = read_equipment(self.excel_path)
            if not items: return messagebox.showerror("Error", "Uskunalar topilmadi!")
            
            template_path = resource_path("shablon.docx")
            if not os.path.exists(template_path): return messagebox.showerror("Error", f"Shablon yo'q: {template_path}")

            date_folder = f"Xulosalar_{dt[2]:04d}-{dt[1]:02d}-{dt[0]:02d}"
            out_dir = os.path.join(os.path.dirname(self.excel_path), date_folder)
            os.makedirs(out_dir, exist_ok=True)

            org_name = re.sub(r'[\\/*?:"<>|]', "_", data["org_name"] or "Hujjat")
            
            if mode == "xulosa": out_file = os.path.join(out_dir, f"Xulosa_{org_name}.docx")
            elif mode == "dalolatnoma": out_file = os.path.join(out_dir, f"Dalolatnoma_{org_name}.docx")
            else: out_file = os.path.join(out_dir, f"Barchasi_{org_name}.docx")

            generate_combined_docs(items, template_path, out_file, data, mode, self.log_msg)
            
            self.last_out_file = out_file
            messagebox.showinfo("Tayyor", f"Yaratildi: {out_file}")
        except Exception as e:
            self.log_msg("❌ XATO: " + str(e)); messagebox.showerror("Error", str(e))

if __name__ == "__main__":
    root = tk.Tk()
    App(root)
    root.mainloop()