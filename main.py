# -*- coding: utf-8 -*-
"""
Xulosa & Dalolatnoma Generator — UI v2.3 (2 Side-by-Side Columns + Smooth Scroll)
"""
import os
import sys
import re
import pandas as pd
import uuid
import time
from docxtpl import DocxTemplate
from docx import Document
from docxcompose.composer import Composer

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from data import CITIES, ORGANIZATIONS, LAVOZIMLAR
from ttkwidgets.autocomplete import AutocompleteCombobox
import warnings
warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

# ---------- Палитра (Dark Blue) ----------
C = {
    "bg":        "#0B132B",   # Глубокий тёмно-синий фон
    "card":      "#1C2541",   # Синий фон карточек
    "border":    "#3A506B",   # Сине-серые границы
    "text":      "#E2E8F0",   # Светло-серый текст (для читаемости)
    "muted":     "#8B9EB7",   # Приглушенный синеватый текст
    "primary":   "#3B82F6",   # Ярко-синий акцент
    "primary_d": "#2563EB",   
    "green":     "#10B981",
    "green_d":   "#059669",
    "amber":     "#F59E0B",
    "amber_d":   "#D97706",
    "blue":      "#6366F1",
    "blue_d":    "#4F46E5",
    "log_bg":    "#060B19",   # Почти черный (для контраста с терминалом)
    "log_fg":    "#CBD5E1",
    "log_ok":    "#34D399",
    "log_err":   "#F87171",
    "log_info":  "#93C5FD",
    "log_dim":   "#475569",
    "log_prompt":"#22D3EE",
}

# ---------- Утилиты ----------
UZ_MONTHS = {1: "yanvar", 2: "fevral", 3: "mart", 4: "aprel", 5: "may", 6: "iyun",
             7: "iyul", 8: "avgust", 9: "sentabr", 10: "oktabr", 11: "noyabr", 12: "dekabr"}

def parse_date(text):
    text = str(text).strip()
    if not text:
        raise ValueError("Sana kiritilmadi!")

    # 1. Если есть разделители (. - /), просто делим по ним
    for sep in (".", "-", "/", " "):
        if sep in text:
            parts = text.split(sep)
            if len(parts) == 3:
                try:
                    d, m, y = int(parts[0]), int(parts[1]), int(parts[2])
                    if 1 <= m <= 12 and 1 <= d <= 31:
                        return d, m, y
                except ValueError:
                    pass

    # 2. Если ввели слитно (например, 11062026 или 110626)
    clean_text = re.sub(r'\D', '', text)  # Удаляем всё, кроме цифр
    if len(clean_text) == 8:   # ДДММГГГГ
        return int(clean_text[0:2]), int(clean_text[2:4]), int(clean_text[4:8])
    elif len(clean_text) == 6: # ДДММГГ
        return int(clean_text[0:2]), int(clean_text[2:4]), int(clean_text[4:6])

    # 3. Если ничего не подошло
    raise ValueError("Sana noto'g'ri. Misol: 11.06.2026 yoki 11062026")

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
        has_name = any(("название" in v) or ("наимен" in v) or ("nomi" in v) for v in row_vals)
        has_inv = any(("инвентар" in v) or ("inventar" in v) for v in row_vals)
        if has_name and has_inv:
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

    col_nomi = find_col("название", "наименование", "nomi", "номи", "оборудование",
                        "аппаратура атамасы", "атамасы", "аппарат", "жихоз", "жихоз номи")
    col_model = find_col("модель", "model", "мод", "марка", "маркасы", "маркасы/модель")
    col_inv = find_col("инвентарный", "inventar", "инв", "инвентар номери", "инв №",
                       "инвентарный номер", "инвентар", "инвентар номери таза")
    col_year = find_col("chiqarilgan", "год", "year", "нешинши жыл келинген",
                        "келинген жылы", "жылы", "жыл", "год выпуска")
    col_zavod = find_col("заводской", "zavod", "serial", "серия", "серия номери",
                         "заводской номер", "серийный номер", "seriya nomeri",
                         "заводской номер (серийный номер)", "заводской номер (серийный)",
                         "заводской №", "заводской № (серийный)", "заводской № (серийный номер)")
    col_masul = find_col("материально", "mas'ul", "masul", "ответствен", "ответ",
                         "ст м/с", "ст м/с (ответственное лицо)", "ответственное лицо", "фио",
                         "фио (ответственное лицо)", "фио материально ответственного лица",
                         "фио материально-ответственного лица", "фио материально-ответственного",
                         "фио материально ответственного")

    items = []
    for i in range(header_row + 1, len(df)):
        row = df.iloc[i].tolist()

        def get(idx):
            return row[idx] if idx is not None and idx < len(row) else None

        nomi = clean(get(col_nomi))
        if nomi == "-" or nomi.lower() in ("итого", "jami", "total"):
            continue

        items.append({
            "equipment_name": nomi,
            "equipment_model": clean(get(col_model)),
            "factory_number": clean(get(col_zavod)),
            "inventory_number": clean(get(col_inv)),
            "year_manufactured": clean(get(col_year)),
            "masul_from_excel": clean(get(col_masul)),
        })
    return items

def generate_combined_docs(items, template_path, out_file, data, mode, log_func, progress_callback=None):
    master = None
    composer = None
    show_xulosa = mode in ("xulosa", "both")
    show_dalolatnoma = mode in ("dalolatnoma", "both")

    run_id = str(uuid.uuid4())[:8]
    tmp_files = []
    total = len(items)

    try:
        for i, item in enumerate(items, 1):
            ctx = dict(data)
            ctx.update(item)
            ctx["num"] = i
            ctx["masul_shaxs"] = data.get("masul_shaxs") or item.get("masul_from_excel", "")
            ctx["region_manufactured"] = item.get("region_manufactured", "")
            ctx["org_head_lavozim"] = data.get("org_head_lavozim", "")
            ctx["masul_shaxs_lavozim"] = data.get("masul_shaxs_lavozim", "")
            ctx["org_stir"] = data.get("org_stir", "")
            ctx["org_str"] = data.get("org_stir", "")
            ctx["show_xulosa"] = show_xulosa
            ctx["show_dalolatnoma"] = show_dalolatnoma

            tpl = DocxTemplate(template_path)
            tpl.render(ctx)
            
            tmp_file = os.path.abspath(f"_tmp_{run_id}_{i}.docx")
            tmp_files.append(tmp_file)
            tpl.save(tmp_file)
            
            doc = None
            for attempt in range(5):
                try:
                    doc = Document(tmp_file)
                    break
                except Exception:
                    if attempt == 4:
                        raise
                    time.sleep(0.1)

            try:
                if os.path.exists(tmp_file):
                    os.remove(tmp_file)
                tmp_files.remove(tmp_file)
            except Exception:
                pass

            for p in doc.paragraphs:
                if not p.text.strip():
                    p._element.getparent().remove(p._element)

            if i < total:
                doc.add_page_break()

            if master is None:
                master = doc
                composer = Composer(master)
            else:
                composer.append(doc)

            log_func(f"[{i}/{total}] {item['equipment_name']} — tayyor", "ok")
            
            # ⬇️ НОВОЕ: Сообщаем интерфейсу, что один файл готов
            if progress_callback:
                progress_callback(i, total)

        composer.save(out_file)

    except Exception as e:
        for f in tmp_files:
            try:
                if os.path.exists(f):
                    os.remove(f)
            except Exception:
                pass
        raise e
# ---------- Словари переводов ----------
TEXTS = {
    "uz": {
        "title": "XULOSA VA DALOLATNOMA",
        "subtitle": "Tibbiy uskunalar uchun xulosa va nuqson dalolatnomasi",
        "btn_excel": "📁  Excel faylni tanlash",
        "no_file": "fayl tanlanmagan",
        "lbl_mode": "Nima yaratilsin?",
        "btn_xulosa": "📄  Faqat Xulosa",
        "btn_dalolatnoma": "📑  Faqat Dalolatnoma",
        "btn_both": "📚  Ikkalasi ham",
        "frame_main": " Asosiy ma'lumotlar (Xulosa) ",
        "frame_dalolatnoma": " Nuqson dalolatnomasi uchun ",
        "btn_generate": "▶   YARATISH",
        "btn_open": "📂  Oxirgi faylni ochish",
        "log_title": "Jurnal",
        "log_clear": "Tozalash",
        "log_ready": "Tizim tayyor. Excel faylini tanlang va YARATISH tugmasini bosing.",
        "ready": "Tayyor",
    },
    "ru": {
        "title": "ЗАКЛЮЧЕНИЕ И ДЕФЕКТНЫЙ АКТ",
        "subtitle": "Заключения и дефектные акты для медицинского оборудования",
        "btn_excel": "📁  Выбрать Excel файл",
        "no_file": "файл не выбран",
        "lbl_mode": "Что написать?",
        "btn_xulosa": "📄  Только Xulosa",
        "btn_dalolatnoma": "📑  Только Dalolatnoma",
        "btn_both": "📚  Оба документа",
        "frame_main": " Основные данные (Xulosa) ",
        "frame_dalolatnoma": " Для Dalolatnoma ",
        "btn_generate": "▶   СОЗДАТЬ",
        "btn_open": "📂  Открыть посл. файл",
        "log_title": "Журнал",
        "log_clear": "Очистить",
        "log_ready": "Система готова. Выберите Excel и нажмите СОЗДАТЬ.",
        "ready": "Готово",
    }
}

# ---------- Вспомогательные виджеты ----------
def _hex_lerp(c1, c2, t):
    c1 = c1.lstrip("#")
    c2 = c2.lstrip("#")
    r1, g1, b1 = int(c1[0:2], 16), int(c1[2:4], 16), int(c1[4:6], 16)
    r2, g2, b2 = int(c2[0:2], 16), int(c2[2:4], 16), int(c2[4:6], 16)
    r = int(r1 + (r2 - r1) * t)
    g = int(g1 + (g2 - g1) * t)
    b = int(b1 + (b2 - b1) * t)
    return f"#{r:02X}{g:02X}{b:02X}"

class GradientButton(tk.Canvas):
    def __init__(self, parent, text, command, width=260, height=48,
                 color_a=C["primary"], color_b=C["primary_d"], **kw):
        super().__init__(parent, width=width, height=height, highlightthickness=0,
                         bg=kw.pop("bg", C["bg"]), bd=0)
        self.command = command
        self.text = text
        self.color_a = color_a
        self.color_b = color_b
        self.w = width
        self.h = height
        self._t = 0.0
        self._hover = False
        self._pressed = False
        self._disabled = False          # ← флаг блокировки
        self._draw()
        self.bind("<Enter>", lambda e: self._set_hover(True))
        self.bind("<Leave>", lambda e: self._set_hover(False))
        self.bind("<ButtonPress-1>", self._on_press)
        self.bind("<ButtonRelease-1>", self._on_release)

    def _rounded_rect(self, x1, y1, x2, y2, r, **kw):
        pts = [x1 + r, y1, x2 - r, y1, x2, y1, x2, y1 + r, x2, y2 - r, x2, y2,
               x2 - r, y2, x1 + r, y2, x1, y2, x1, y2 - r, x1, y1 + r, x1, y1]
        return self.create_polygon(pts, smooth=True, **kw)

    def _draw(self):
        self.delete("all")
        if self._disabled:
            # серый вариант для заблокированного состояния
            self._rounded_rect(2, 2, self.w - 2, self.h - 2, 12,
                               fill="#9CA3AF", outline="")
            self.create_text(self.w / 2, self.h / 2, text=self.text,
                             fill="#E5E7EB", font=("Arial", 12, "bold"))
            return

        t = self._t
        self._rounded_rect(2, 2, self.w - 2, self.h - 2, 12, fill=self.color_a, outline="")
        steps = 40
        for i in range(steps):
            ratio = i / (steps - 1)
            col = _hex_lerp(self.color_a, self.color_b, ratio * (0.35 + 0.65 * t))
            x1 = 2 + (self.w - 4) * i / steps
            x2 = 2 + (self.w - 4) * (i + 1) / steps + 1
            self.create_rectangle(x1, 2, x2, self.h - 2, fill=col, outline="")
        self._rounded_rect(2, 2, self.w - 2, self.h - 2, 12, fill="", outline=C["card"], width=0)
        self.create_text(self.w / 2, self.h / 2, text=self.text, fill="white",
                         font=("Arial", 12, "bold"))

    def _set_hover(self, val):
        if self._disabled:
            return
        self._hover = val
        self._animate_to(1.0 if val else 0.0)

    def _animate_to(self, target):
        def step():
            diff = target - self._t
            if abs(diff) < 0.05:
                self._t = target
                self._draw()
                return
            self._t += diff * 0.25
            self._draw()
            self.after(16, step)
        step()

    def _on_press(self, _):
        if self._disabled:
            return
        self._pressed = True
        self.move("all", 0, 2)

    def _on_release(self, _):
        if self._disabled or not self._pressed:
            return
        self.move("all", 0, -2)
        self._pressed = False
        if callable(self.command):
            self.command()

    # ── Публичные методы блокировки ──
    def set_disabled(self, disabled: bool):
        self._disabled = disabled
        self._hover = False
        self._t = 0.0
        self._pressed = False
        self._draw()

class HoverButton(ttk.Button):
    def __init__(self, parent, text, command, base_style="TButton",
                 hover_style="Hover.TButton", **kw):
        super().__init__(parent, text=text, command=command, style=base_style, **kw)
        self._base = base_style
        self._hover = hover_style
        self.bind("<Enter>", lambda e: self.configure(style=self._hover))
        self.bind("<Leave>", lambda e: self.configure(style=self._base))





# ---------- GUI ----------
class App:
    def __init__(self, root):
        self.root = root
        self.root.title("Xulosa Generator")
        self.root.geometry("1180x900")
        self.root.state("zoomed")
        self.root.minsize(1050, 750)
        self.root.configure(bg=C["bg"])

        self.current_lang = "uz"
        self.excel_path = None
        self.last_out_file = None
        self.gen_mode = tk.StringVar(value="both")
        self._mode_buttons = {}
        self.fields = {}

        self._setup_styles()
        self._build_ui()
        self._apply_defaults()
        self.update_ui_texts()
        self._log_ready()

    # ---------- Стили ----------
    def _setup_styles(self):
        s = ttk.Style()
        try:
            s.theme_use("clam")
        except tk.TclError:
            pass
        s.configure(".", font=("Segoe UI", 10), background=C["bg"], foreground=C["text"])
        s.configure("TFrame", background=C["bg"])
        s.configure("Card.TFrame", background=C["card"])
        s.configure("TLabel", background=C["bg"], foreground=C["text"])
        s.configure("Muted.TLabel", background=C["bg"], foreground=C["muted"], font=("Segoe UI", 9))

        s.configure("TButton", padding=(12, 8), font=("Segoe UI", 10))
        s.configure("Hover.TButton", padding=(12, 8), font=("Segoe UI", 10),
                    background="#EEF2FF", foreground=C["primary_d"])
        s.configure("Log.TButton", padding=(8, 4), font=("Consolas", 9))

        s.configure("Mode.Xulosa.TButton", padding=(14, 10), font=("Segoe UI", 10, "bold"),
                    background="#E7F8F0", foreground=C["green_d"])
        s.configure("Mode.Dalolatnoma.TButton", padding=(14, 10), font=("Segoe UI", 10, "bold"),
                    background="#EEF2FF", foreground=C["blue_d"])
        s.configure("Mode.Both.TButton", padding=(14, 10), font=("Segoe UI", 10, "bold"),
                    background="#FFF7E6", foreground=C["amber_d"])

        s.configure("TEntry", padding=6, fieldbackground=C["card"],
                    bordercolor=C["border"], lightcolor=C["border"], darkcolor=C["border"])
        s.configure("TCombobox", padding=6, fieldbackground=C["card"],
                    bordercolor=C["border"], lightcolor=C["border"], darkcolor=C["border"])

    def _card(self, parent):
        return tk.Frame(parent, bg=C["card"], highlightbackground=C["border"],
                        highlightthickness=1, bd=0)

    def _section_title(self, parent, text):
        return tk.Label(parent, text=text, bg=C["card"], fg=C["text"],
                        font=("Segoe UI", 11, "bold"), anchor="w")

    # ---------- Построение ----------
    def _build_ui(self):
        # Язык
        top = tk.Frame(self.root, bg=C["bg"])
        top.pack(fill="x", padx=20, pady=(16, 0))
        self.lang_var = tk.StringVar(value="uz")
        for code, label in (("ru", "RU"), ("uz", "UZ")):
            ttk.Radiobutton(top, text=label, variable=self.lang_var, value=code,
                            command=lambda c=code: self.change_lang(c)).pack(side="right", padx=4)

        # Заголовок
        head = tk.Frame(self.root, bg=C["bg"])
        head.pack(fill="x", padx=20, pady=(10, 4))
        self.lbl_title = tk.Label(head, text="", bg=C["bg"], fg=C["text"],
                                  font=("Segoe UI", 20, "bold"), anchor="w")
        self.lbl_title.pack(anchor="w")
        self.lbl_subtitle = tk.Label(head, text="", bg=C["bg"], fg=C["muted"],
                                     font=("Segoe UI", 10), anchor="w")
        self.lbl_subtitle.pack(anchor="w", pady=(2, 0))

        tk.Frame(self.root, bg=C["border"], height=1).pack(fill="x", padx=20, pady=10)



        # ------- САМАЯ НИЖНЯЯ ПАНЕЛЬ (Водяной знак) -------
        author_frame = tk.Frame(self.root, bg=C["bg"])
        author_frame.pack(side="bottom", fill="x", padx=20, pady=(0, 6))
        self.lbl_author = tk.Label(
            author_frame, 
            text="jahondv", 
            bg=C["bg"], 
            fg="#B0B8C4",  # Еле заметный цвет
            font=("Segoe UI", 8)
        )
        self.lbl_author.pack(side="right")

        # ------- НИЖНЯЯ ЧАСТЬ: терминал -------
        self._build_terminal(self.root)

        # ------- НИЖНЯЯ ПАНЕЛЬ КНОПОК -------
        bottom = tk.Frame(self.root, bg=C["bg"])
        bottom.pack(side="bottom", fill="x", padx=20, pady=(4, 8))
        self.btn_generate = GradientButton(bottom, text="", command=self.generate,
                                           width=300, height=50)
        self.btn_generate.pack(side="left")
        self.btn_open = HoverButton(bottom, text="", command=self.open_last, width=26)
        self.btn_open.pack(side="left", padx=10)

        # ------- СКРОЛЛ-ОБЛАСТЬ -------
        self.scroll_frame = tk.Frame(self.root, bg=C["bg"])
        self.scroll_frame.pack(side="top", fill="both", expand=True, padx=20)

        canvas = tk.Canvas(self.scroll_frame, bg=C["bg"], highlightthickness=0)
        scrollbar = ttk.Scrollbar(self.scroll_frame, orient="vertical", command=canvas.yview)
        
        self.inner = tk.Frame(canvas, bg=C["bg"])
        
        # Привязываем ширину inner к ширине canvas, чтобы не было прилипания влево
        canvas.bind("<Configure>", lambda e: canvas.itemconfigure(canvas_window, width=e.width))
        canvas_window = canvas.create_window((0, 0), window=self.inner, anchor="nw")
        
        self.inner.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.configure(yscrollcommand=scrollbar.set)
        
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        self._canvas = canvas

        
        # Прокрутка колесиком мыши (умная)
        def _on_mousewheel(event):
            # Если курсор над выпадающим списком — не скроллим окно, отдаем событие списку
            if isinstance(event.widget, str):
                return
            if event.widget.winfo_class() == "Listbox":
                return
            # Скроллим окно
            canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
            
        canvas.bind_all("<MouseWheel>", _on_mousewheel)

        # Excel (в карточке)
        row_card = self._card(self.inner)
        row_card.pack(fill="x", pady=(0, 8))
        row = tk.Frame(row_card, bg=C["card"])
        row.pack(fill="x", padx=16, pady=12)
        self.btn_excel = HoverButton(row, text="", command=self.select_excel, width=26)
        self.btn_excel.pack(side="left")
        self.lbl_excel = tk.Label(row, text="", bg=C["card"], fg=C["muted"], font=("Segoe UI", 10))
        self.lbl_excel.pack(side="left", padx=12)

        # Режим
        self.mode_card = self._card(self.inner)
        self.mode_card.pack(fill="x", pady=8)
        mode_inner = tk.Frame(self.mode_card, bg=C["card"])
        mode_inner.pack(fill="x", padx=16, pady=14)
        self.lbl_mode = tk.Label(mode_inner, text="", bg=C["card"], fg=C["text"],
                                 font=("Segoe UI", 11, "bold"), anchor="w")
        self.lbl_mode.pack(anchor="w", pady=(0, 10))
        modes_row = tk.Frame(mode_inner, bg=C["card"])
        modes_row.pack(fill="x")

        def make_mode_btn(key, style_name):
            b = ttk.Button(modes_row, text="", style=style_name,
                           command=lambda: self._set_mode(key))
            b.pack(side="left", padx=(0, 8))
            self._mode_buttons[key] = b
            return b

        self.btn_xulosa = make_mode_btn("xulosa", "Mode.Xulosa.TButton")
        self.btn_dalolatnoma = make_mode_btn("dalolatnoma", "Mode.Dalolatnoma.TButton")
        self.btn_both = make_mode_btn("both", "Mode.Both.TButton")
        self._set_mode("both")

        # ------- ДВЕ КОЛОНКИ -------
        self.columns_frame = tk.Frame(self.inner, bg=C["bg"])
        self.columns_frame.pack(fill="both", expand=True)
        
        # Делаем 2 колонки строго 50/50
        self.columns_frame.columnconfigure(0, weight=1, uniform="col")
        self.columns_frame.columnconfigure(1, weight=1, uniform="col")

        # Xulosa (Слева)
        self.frame_main = self._build_form_card(self.columns_frame, dalolatnoma=False)
        self.frame_main.grid(row=0, column=0, sticky="nsew", padx=(0, 8))

        # Dalolatnoma (Справа)
        self.frame_dalolatnoma = self._build_form_card(self.columns_frame, dalolatnoma=True)
        self.frame_dalolatnoma.grid(row=0, column=1, sticky="nsew", padx=(8, 0))

    def _build_terminal(self, parent):
        """Терминал внизу: тёмный, с заголовком, прогресс-баром и кнопкой очистки."""
        wrap = tk.Frame(parent, bg=C["log_bg"], highlightbackground=C["border"],
                        highlightthickness=1)
        wrap.pack(side="bottom", fill="x", padx=20, pady=(6, 12))

        # --- Шапка: заголовок + прогресс + кнопка ---
        header = tk.Frame(wrap, bg=C["log_bg"])
        header.pack(fill="x", padx=10, pady=(6, 0))
        
        self.lbl_log_title = tk.Label(header, text="", bg=C["log_bg"], fg=C["log_prompt"],
                                      font=("Consolas", 10, "bold"), anchor="w")
        self.lbl_log_title.pack(side="left")
        
        self.btn_log_clear = ttk.Button(header, text="", style="Log.TButton",
                                        command=self._clear_log)
        self.btn_log_clear.pack(side="right")

        # --- ПРОГРЕСС-БАР ---
        prog_frame = tk.Frame(wrap, bg=C["log_bg"])
        prog_frame.pack(fill="x", padx=10, pady=(4, 2))
        
        self.lbl_progress = tk.Label(prog_frame, text="Kutilmoqda...", bg=C["log_bg"], 
                                     fg=C["log_dim"], font=("Consolas", 9))
        self.lbl_progress.pack(side="left")
        
        self.progress_var = tk.DoubleVar(value=0)
        style = ttk.Style()
        style.configure("Custom.Horizontal.TProgressbar", 
                        troughcolor="#1E293B", background=C["primary"], 
                        lightcolor=C["primary"], darkcolor=C["primary"])
        self.progress_bar = ttk.Progressbar(prog_frame, variable=self.progress_var, 
                                            maximum=100, style="Custom.Horizontal.TProgressbar")
        self.progress_bar.pack(side="right", fill="x", expand=True, padx=(10, 0))

        # --- Тело лога ---
        body = tk.Frame(wrap, bg=C["log_bg"])
        body.pack(fill="x", padx=10, pady=(4, 10))

        self.log = tk.Text(body, height=7, font=("Consolas", 9),
                           bg=C["log_bg"], fg=C["log_fg"], relief="flat", bd=0,
                           insertbackground=C["log_fg"], wrap="word")
        self.log.pack(side="left", fill="both", expand=True)

        sb = ttk.Scrollbar(body, orient="vertical", command=self.log.yview)
        sb.pack(side="right", fill="y")
        self.log.configure(yscrollcommand=sb.set)

        self.log.tag_configure("ok", foreground=C["log_ok"])
        self.log.tag_configure("err", foreground=C["log_err"])
        self.log.tag_configure("info", foreground=C["log_info"])
        self.log.tag_configure("dim", foreground=C["log_dim"])
        self.log.tag_configure("prompt", foreground=C["log_prompt"])
        self.log.configure(state="disabled")

    def _build_form_card(self, parent, dalolatnoma=False):
        card = self._card(parent)
        inner = tk.Frame(card, bg=C["card"])
        inner.pack(fill="both", expand=True, padx=16, pady=14)
        
        title = self._section_title(inner, "")
        title.pack(anchor="w", pady=(0, 10))
        card._title = title

        # Внутри карточки поля идут в одну колонку (чтобы влезло в 50% ширины)
        fields_container = tk.Frame(inner, bg=C["card"])
        fields_container.pack(fill="both", expand=True)
        fields_container.columnconfigure(0, weight=1)

        if not dalolatnoma:
            fields = [
                ("city", "Shahar:" if self.current_lang == "uz" else "Город:", "combo", CITIES),
                ("inspection_date", "Sana:" if self.current_lang == "uz" else "Дата:", "entry", None),
                ("org_name", "Tashkilot nomi:" if self.current_lang == "uz" else "Название организации:", "combo", ORGANIZATIONS),
                ("org_stir", "STIR:", "entry", None),
                ("org_address", "Tashkilot manzili:" if self.current_lang == "uz" else "Адрес организации:", "entry", None),
                ("org_head", "Tashkiloting rahbari:" if self.current_lang == "uz" else "Руководитель:", "entry", None),
                ("bosh_hisobchi", "Bosh hisobchi:" if self.current_lang == "uz" else "Главный бухгалтер:", "entry", None),
                ("masul_shaxs", "Mas'ul shaxs:" if self.current_lang == "uz" else "Ответственное лицо:", "entry", None),
                ("fish", "Mutaxassis F.I.Sh.:" if self.current_lang == "uz" else "Ф.И.О. эксперта:", "entry", None),
                ("lavozim", "Mutaxassis lavozimi:" if self.current_lang == "uz" else "Должность эксперта:", "combo", LAVOZIMLAR),
            ]
            title.configure(text=TEXTS[self.current_lang]["frame_main"])
            self._build_fields_grid(fields_container, fields, columns=1)
            return card
        else:
            fields = [
                ("org_head_lavozim", "Rahbar lavozimi:" if self.current_lang == "uz" else "Должность руководителя:", "combo", ["Bosh shifokori", "Direktor"]),
                ("masul_shaxs_lavozim", "Mas'ul shaxs lavozimi:" if self.current_lang == "uz" else "Должность отв. лица:", "combo", ["texnik", "buxgalter", "muhandis"]),
                ("kafolat", "Kafolat muddati:" if self.current_lang == "uz" else "Гарантийный срок:", "combo", ["amal qilmoqda", "tugagan"]),
                ("tamirlash", "Uskuna ta'mirlashga:" if self.current_lang == "uz" else "Оборудование к ремонту:", "combo", ["yaroqli", "yaroqsiz"]),
                ("metrologik", "Metrologik tekshiruv:" if self.current_lang == "uz" else "Метрологическая проверка:", "combo", ["qayta o'tkazish zarur", "qayta o'tkazishga zarurat yo'q"]),
                ("foydalanish", "Foydalanishni davom ettirish:" if self.current_lang == "uz" else "Дальнейшая эксплуатация:", "combo", ["ha (mumkin)", "yo'q (mumkin emas)"]),
            ]
            title.configure(text=TEXTS[self.current_lang]["frame_dalolatnoma"])
            self._build_fields_grid(fields_container, fields, columns=1)
            return card

    def _build_fields_grid(self, parent, fields_config, columns=1):
        """Раскладывает поля. Поля растягиваются по ширине."""
        for idx, (key, lbl, kind, values) in enumerate(fields_config):
            if key in self.fields:
                continue
            row = idx // columns
            col = idx % columns
            
            cell = tk.Frame(parent, bg=C["card"])
            cell.grid(row=row, column=col, sticky="nsew", pady=6, padx=0)

            tk.Label(cell, text=lbl, bg=C["card"], fg=C["muted"],
                     font=("Segoe UI", 9), anchor="w").pack(anchor="w", fill="x")
            
            if kind == "combo":
                w = AutocompleteCombobox(cell, completevalues=values, font=("Segoe UI", 10))
                w.pack(fill="x", expand=True, pady=(2, 0))
            else:
                w = ttk.Entry(cell, font=("Segoe UI", 10))
                w.pack(fill="x", expand=True, pady=(2, 0))
                
                # ⬇️ АВТО-ФОРМАТИРОВАНИЕ ДАТЫ ⬇️
                if key == "inspection_date":
                    w.bind("<KeyRelease>", self._auto_format_date)
                    
            self.fields[key] = w

    def _apply_defaults(self):
        if "city" in self.fields:
            self.fields["city"].set("Nukus shahri")
        if "org_head_lavozim" in self.fields:
            self.fields["org_head_lavozim"].set("Bosh shifokori")
        if "masul_shaxs_lavozim" in self.fields:
            self.fields["masul_shaxs_lavozim"].set("texnik")
        if "lavozim" in self.fields:
            self.fields["lavozim"].set("Kichik muhandis")
        if "fish" in self.fields:
            self.fields["fish"].insert(0, "B.K.Utemisov")


            
    def _auto_format_date(self, event):
        """Автоматически ставит точки в дату при вводе (11062006 -> 11.06.2006)"""
        # Игнорируем клавиши удаления, чтобы не мешать исправлять
        if event.keysym in ("BackSpace", "Delete", "Left", "Right", "Up", "Down", "Tab"):
            return
            
        entry = event.widget
        text = entry.get()
        
        # Оставляем только цифры
        digits = re.sub(r'\D', '', text)
        
        # Форматируем в зависимости от количества введенных цифр
        formatted = digits
        if len(digits) > 4:
            formatted = f"{digits[:2]}.{digits[2:4]}.{digits[4:8]}"
        elif len(digits) > 2:
            formatted = f"{digits[:2]}.{digits[2:4]}"
            
        # Если текст изменился, обновляем поле и ставим курсор в конец
        if text != formatted:
            entry.delete(0, tk.END)
            entry.insert(0, formatted)

    # ---------- Лог/терминал ----------
    def _log_ready(self):
        self.log_msg(TEXTS[self.current_lang]["log_ready"], "dim")

    def _clear_log(self):
        self.log.configure(state="normal")
        self.log.delete("1.0", "end")
        self.log.configure(state="disabled")
        self._log_ready()

    def log_msg(self, msg, tag="info"):
        self.log.configure(state="normal")
        prefix = {
            "ok": "✓ ",
            "err": "✗ ",
            "info": "› ",
            "dim": "$ ",
        }.get(tag, "› ")
        self.log.insert("end", f"{prefix}{msg}\n", tag)
        self.log.see("end")
        self.log.configure(state="disabled")
    
    def _update_progress(self, current, total):
        percent = int((current / total) * 100)
        self.progress_var.set(percent)
        if current == total:
            self.lbl_progress.config(text=f"Tayyor: {current} / {total} (100%)", fg=C["log_ok"])
        else:
            self.lbl_progress.config(text=f"Jarayon: {current} / {total} ({percent}%)", fg=C["log_info"])
        self.root.update_idletasks()


    # ---------- Режим ----------
    def _set_mode(self, mode):
        self.gen_mode.set(mode)
        for key, btn in self._mode_buttons.items():
            btn.configure(style="TButton", state="normal")
        active = {
            "xulosa": "Mode.Xulosa.TButton",
            "dalolatnoma": "Mode.Dalolatnoma.TButton",
            "both": "Mode.Both.TButton",
        }[mode]
        self._mode_buttons[mode].configure(style=active)

    # ---------- Тексты ----------
    def change_lang(self, lang):
        self.current_lang = lang
        self.update_ui_texts()

    def update_ui_texts(self):
        t = TEXTS[self.current_lang]
        self.lbl_title.config(text=t["title"])
        self.lbl_subtitle.config(text=t["subtitle"])
        self.btn_excel.config(text=t["btn_excel"])
        if not self.excel_path:
            self.lbl_excel.config(text=t["no_file"])
        self.lbl_mode.config(text=t["lbl_mode"])
        self.btn_xulosa.config(text=t["btn_xulosa"])
        self.btn_dalolatnoma.config(text=t["btn_dalolatnoma"])
        self.btn_both.config(text=t["btn_both"])
        self.frame_main._title.config(text=t["frame_main"])
        self.frame_dalolatnoma._title.config(text=t["frame_dalolatnoma"])
        self.btn_generate.text = t["btn_generate"]
        self.btn_generate._draw()
        self.btn_open.config(text=t["btn_open"])
        self.lbl_log_title.config(text=f"> {t['log_title']}")
        self.btn_log_clear.config(text=t["log_clear"])

    # ---------- Действия ----------
    def select_excel(self):
        path = filedialog.askopenfilename(
            title="Excel faylni tanlang",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        if path:
            self.excel_path = path
            self.lbl_excel.config(text=os.path.basename(path), fg=C["text"])
            self.log_msg(f"Excel tanlandi: {os.path.basename(path)}", "info")

    def open_last(self):
        if not self.last_out_file or not os.path.exists(self.last_out_file):
            messagebox.showinfo("Info", "Fayl topilmadi.")
            return
        try:
            if sys.platform.startswith("win"):
                os.startfile(self.last_out_file)
            elif sys.platform == "darwin":
                os.system(f'open "{self.last_out_file}"')
            else:
                os.system(f'xdg-open "{self.last_out_file}"')
        except Exception as e:
            messagebox.showerror("Error", str(e))

    def generate(self):
        if getattr(self, "_is_generating", False):
            return
        self._is_generating = True
        self.btn_generate.set_disabled(True)     # ← блокируем Canvas-кнопку
        self.root.update_idletasks()

        try:
            if not self.excel_path:
                messagebox.showwarning("Warning", "Excel tanlang!")
                return
            required = ("inspection_date", "org_name", "city")
            missing = [k for k in required if k not in self.fields]
            if missing:
                messagebox.showerror("Error", f"Maydonlar topilmadi: {missing}")
                return
            data = {k: v.get().strip() for k, v in self.fields.items()}
            if not data["inspection_date"]:
                messagebox.showwarning("Warning", "Sanani kiriting!")
                return
            try:
                dt = parse_date(data["inspection_date"])
            except Exception as e:
                messagebox.showerror("Error", str(e))
                return

            data["inspection_date"] = date_to_uz(dt)
            data["inspection_date_bold"] = date_to_uz_bold(dt)
            mode = self.gen_mode.get()

            self.progress_var.set(0)
            self.lbl_progress.config(text="0%", fg=C["log_dim"])
            self.log_msg("Boshlanmoqda...", "info")
            self.log_msg("Boshlanmoqda...", "info")
            items = read_equipment(self.excel_path)
            if not items:
                messagebox.showerror("Error", "Uskunalar topilmadi!")
                return
            self.log_msg(f"{len(items)} ta uskuna topildi", "info")

            template_path = resource_path("shablon.docx")
            if not os.path.exists(template_path):
                messagebox.showerror("Error", f"Shablon yo'q: {template_path}")
                return

            date_folder = f"Xulosalar_{dt[2]:04d}-{dt[1]:02d}-{dt[0]:02d}"
            out_dir = os.path.join(os.path.dirname(self.excel_path), date_folder)
            os.makedirs(out_dir, exist_ok=True)
            org_name = re.sub(r'[\\/*?:"<>|]', "_", data["org_name"] or "Hujjat")

            if mode == "xulosa":
                out_file = os.path.join(out_dir, f"Xulosa_{org_name}.docx")
            elif mode == "dalolatnoma":
                out_file = os.path.join(out_dir, f"Dalolatnoma_{org_name}.docx")
            else:
                out_file = os.path.join(out_dir, f"Barchasi_{org_name}.docx")

            if os.path.exists(out_file):
                base, ext = os.path.splitext(out_file)
                counter = 2
                while os.path.exists(f"{base}_v{counter}{ext}"):
                    counter += 1
                out_file = f"{base}_v{counter}{ext}"

            generate_combined_docs(items, template_path, out_file, data, mode, self.log_msg, 
                                   progress_callback=self._update_progress)
            self.last_out_file = out_file
            self.log_msg(f"Saqlandi: {out_file}", "ok")

            # ── messagebox показываем ВНЕ try/except, но всё равно в нём ──
            messagebox.showinfo(TEXTS[self.current_lang]["ready"], f"{out_file}")

        except Exception as e:
            self.log_msg("XATO: " + str(e), "err")
            messagebox.showerror("Error", str(e))
        finally:
            # ── ГАРАНТИРОВАННЫЙ сброс кнопки, даже если было исключение ──
            self._is_generating = False
            self.btn_generate.set_disabled(False)
            self.root.update_idletasks()

if __name__ == "__main__":
    root = tk.Tk()
    App(root)
    root.mainloop()