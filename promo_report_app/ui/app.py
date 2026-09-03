# -*- coding: utf-8 -*-
"""Tkinter-приложение: отчёт по товарам для участия в акции Ozon."""

import os
import queue
import threading
from collections import Counter

import tkinter as tk
from tkinter import ttk, filedialog, messagebox

from config import (
    MAX_CO_FILES,
    DEFAULT_RED_MAX,
    DEFAULT_GREEN_MIN,
    DEFAULT_STOCK_DAYS_HIGH,
    DEFAULT_LOW_SALES_PER_WEEK,
)
from readers.co_reader import load_co_files, list_available_schemes
from readers.promo_reader import load_promo_file
from calc.evaluate import build_report, Thresholds, STATUS_LABELS
from export.excel_export import export_report

TREE_COLUMNS = [
    ("article", "Артикул", 140),
    ("name", "Название", 320),
    ("category", "Категория", 160),
    ("scheme", "Схема", 60),
    ("price_co", "Наша цена", 90),
    ("price_promo_max", "Цена акции", 90),
    ("price_used", "Цена участия", 100),
    ("pct_of_cost", "% от себест.", 90),
    ("net_profit", "Прибыль", 90),
    ("status_label", "Статус", 200),
    ("price_change_note", "Изменить цену в ЦО", 160),
    ("ordered_qty_7d", "Заказано, шт/нед", 110),
    ("stock_ozon", "Остаток Ozon", 90),
    ("stock_own", "Остаток мой склад", 110),
    ("stock_note", "Рекомендация по остаткам", 260),
]

FILTER_OPTIONS = [
    ("all", "Все"),
    ("green", "Одобрен"),
    ("yellow", "Пограничный"),
    ("red", "Отклонён"),
    ("no_co", "Нет данных в ЦО / акции"),
]

TAG_COLORS = {
    "green": "#C6EFCE",
    "yellow": "#FFEB9C",
    "red": "#FFC7CE",
    "no_co": "#D9D9D9",
    "no_price_data": "#D9D9D9",
    "error": "#D9D9D9",
}


class App:
    def __init__(self, root):
        self.root = root
        root.title("Отчёт по товарам для участия в акции Ozon")
        root.geometry("1280x760")

        self.co_file_rows = []
        self.promo_path = None
        self.promo_name = None
        self.report_rows = []
        self.last_thresholds = Thresholds()
        self.result_queue = queue.Queue()
        self._sort_state = {}

        self._build_ui()

    # ---------------------------------------------------------------- UI ---

    def _build_ui(self):
        top = ttk.Frame(self.root, padding=8)
        top.pack(fill="x")

        co_frame = ttk.LabelFrame(top, text=f"Файлы Ценообразования (до {MAX_CO_FILES})", padding=6)
        co_frame.pack(fill="x", pady=(0, 6))
        self.co_files_container = ttk.Frame(co_frame)
        self.co_files_container.pack(fill="x")
        ttk.Button(co_frame, text="Добавить файл ЦО...", command=self.add_co_file).pack(anchor="w", pady=(4, 0))

        promo_frame = ttk.LabelFrame(top, text="Файл акции Ozon", padding=6)
        promo_frame.pack(fill="x", pady=(0, 6))
        ttk.Button(promo_frame, text="Выбрать файл акции...", command=self.choose_promo_file).pack(side="left")
        self.promo_label = ttk.Label(promo_frame, text="(не выбран)")
        self.promo_label.pack(side="left", padx=8)

        thr_frame = ttk.LabelFrame(top, text="Пороги одобрения по % от себестоимости", padding=6)
        thr_frame.pack(fill="x", pady=(0, 6))
        ttk.Label(thr_frame, text="Красная зона ниже, %:").pack(side="left")
        self.red_max_var = tk.StringVar(value=str(int(DEFAULT_RED_MAX * 100)))
        ttk.Entry(thr_frame, textvariable=self.red_max_var, width=5).pack(side="left", padx=(4, 16))
        ttk.Label(thr_frame, text="Зелёная зона от, %:").pack(side="left")
        self.green_min_var = tk.StringVar(value=str(int(DEFAULT_GREEN_MIN * 100)))
        ttk.Entry(thr_frame, textvariable=self.green_min_var, width=5).pack(side="left", padx=(4, 16))

        stock_thr_frame = ttk.LabelFrame(
            top, text="Пороги рекомендации по остаткам (справочно, на статус не влияет)", padding=6
        )
        stock_thr_frame.pack(fill="x", pady=(0, 6))
        ttk.Label(stock_thr_frame, text="«Много остатков» от, дней запаса:").pack(side="left")
        self.stock_days_high_var = tk.StringVar(value=str(DEFAULT_STOCK_DAYS_HIGH))
        ttk.Entry(stock_thr_frame, textvariable=self.stock_days_high_var, width=6).pack(side="left", padx=(4, 16))
        ttk.Label(stock_thr_frame, text="«Мало продаж» до, шт/неделю:").pack(side="left")
        self.low_sales_var = tk.StringVar(value=str(DEFAULT_LOW_SALES_PER_WEEK))
        ttk.Entry(stock_thr_frame, textvariable=self.low_sales_var, width=6).pack(side="left", padx=(4, 16))

        action_frame = ttk.Frame(top)
        action_frame.pack(fill="x", pady=(0, 6))
        self.calc_button = ttk.Button(action_frame, text="Рассчитать", command=self.run_calculation)
        self.calc_button.pack(side="left")
        self.export_button = ttk.Button(action_frame, text="Экспортировать в Excel...", command=self.export, state="disabled")
        self.export_button.pack(side="left", padx=8)

        ttk.Label(action_frame, text="Показать:").pack(side="left", padx=(24, 4))
        self.filter_var = tk.StringVar(value="all")
        filter_combo = ttk.Combobox(
            action_frame, textvariable=self.filter_var, state="readonly", width=24,
            values=[label for _, label in FILTER_OPTIONS],
        )
        filter_combo.current(0)
        filter_combo.pack(side="left")
        filter_combo.bind("<<ComboboxSelected>>", self.on_filter_change)

        self.status_label = ttk.Label(top, text="Добавьте файлы и нажмите «Рассчитать».")
        self.status_label.pack(fill="x", pady=(0, 4))

        tree_frame = ttk.Frame(self.root, padding=(8, 0, 8, 8))
        tree_frame.pack(fill="both", expand=True)

        columns = [c[0] for c in TREE_COLUMNS]
        self.tree = ttk.Treeview(tree_frame, columns=columns, show="headings")
        for key, label, width in TREE_COLUMNS:
            self.tree.heading(key, text=label, command=lambda k=key: self._sort_by(k))
            self.tree.column(key, width=width, anchor="w")
        for tag, color in TAG_COLORS.items():
            self.tree.tag_configure(tag, background=color)

        vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(tree_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        tree_frame.rowconfigure(0, weight=1)
        tree_frame.columnconfigure(0, weight=1)

    # ------------------------------------------------------------ Файлы ---

    def add_co_file(self):
        if len(self.co_file_rows) >= MAX_CO_FILES:
            messagebox.showwarning("Лимит файлов", f"Можно загрузить не более {MAX_CO_FILES} файлов ЦО.")
            return
        path = filedialog.askopenfilename(
            title="Выберите файл Ценообразование",
            filetypes=[("Excel", "*.xlsm *.xlsx"), ("Все файлы", "*.*")],
        )
        if not path:
            return

        try:
            schemes_in_file = list_available_schemes(path)
        except Exception as e:  # noqa: BLE001
            messagebox.showerror("Ошибка чтения файла", f"Не удалось открыть файл:\n{e}")
            return

        if not schemes_in_file:
            messagebox.showwarning(
                "Листы не найдены",
                f"В файле {os.path.basename(path)} не найдено ни одного известного листа-схемы "
                f"(СП / FBS / Китай / Россия / Неликвид / Монеты). Файл не добавлен.",
            )
            return

        row_frame = ttk.Frame(self.co_files_container)
        row_frame.pack(fill="x", pady=2)
        ttk.Label(row_frame, text=os.path.basename(path), width=45, anchor="w").pack(side="left", padx=4)

        scheme_vars = {}
        for scheme in schemes_in_file:
            var = tk.BooleanVar(value=True)
            ttk.Checkbutton(row_frame, text=scheme, variable=var).pack(side="left", padx=4)
            scheme_vars[scheme] = var

        entry = {"frame": row_frame, "path": path, "vars": scheme_vars}
        ttk.Button(row_frame, text="Удалить", command=lambda: self._remove_co_file(entry)).pack(side="left", padx=8)
        self.co_file_rows.append(entry)

    def _remove_co_file(self, entry):
        entry["frame"].destroy()
        self.co_file_rows.remove(entry)

    def choose_promo_file(self):
        path = filedialog.askopenfilename(
            title="Выберите файл акции Ozon",
            filetypes=[("Excel", "*.xlsx *.xlsm"), ("Все файлы", "*.*")],
        )
        if not path:
            return
        self.promo_path = path
        self.promo_label.config(text=os.path.basename(path))

    # ------------------------------------------------------------ Расчёт ---

    def _get_thresholds(self):
        try:
            red_max = float(self.red_max_var.get().replace(",", ".")) / 100.0
            green_min = float(self.green_min_var.get().replace(",", ".")) / 100.0
            stock_days_high = float(self.stock_days_high_var.get().replace(",", "."))
            low_sales_per_week = float(self.low_sales_var.get().replace(",", "."))
        except ValueError:
            raise ValueError(
                "Пороги должны быть числами (например 40 и 50 для маржи, "
                "60 дней и 3 шт/нед для остатков)"
            )
        if not (0 <= red_max <= green_min <= 1):
            raise ValueError("Должно выполняться: 0 <= красная граница <= зелёная граница <= 100")
        if stock_days_high <= 0 or low_sales_per_week < 0:
            raise ValueError("Пороги по остаткам должны быть положительными числами")
        return Thresholds(
            red_max=red_max,
            green_min=green_min,
            stock_days_high=stock_days_high,
            low_sales_per_week=low_sales_per_week,
        )

    def run_calculation(self):
        if not self.co_file_rows:
            messagebox.showwarning("Нет файлов", "Добавьте хотя бы один файл ЦО.")
            return
        if not any(v.get() for row in self.co_file_rows for v in row["vars"].values()):
            messagebox.showwarning("Нет листов", "Отметьте хотя бы один лист (СП/FBS/Неликвид) для загрузки.")
            return
        if not self.promo_path:
            messagebox.showwarning("Нет файла акции", "Выберите файл акции Ozon.")
            return
        try:
            thresholds = self._get_thresholds()
        except ValueError as e:
            messagebox.showerror("Некорректные пороги", str(e))
            return

        self.last_thresholds = thresholds
        self.calc_button.config(state="disabled")
        self.export_button.config(state="disabled")
        self.status_label.config(text="Загрузка файлов и расчёт...")

        file_scheme_pairs = [
            (row["path"], [scheme for scheme, var in row["vars"].items() if var.get()])
            for row in self.co_file_rows
        ]
        promo_path = self.promo_path

        threading.Thread(
            target=self._worker, args=(file_scheme_pairs, promo_path, thresholds), daemon=True
        ).start()
        self.root.after(100, self._poll_queue)

    def _worker(self, file_scheme_pairs, promo_path, thresholds):
        try:
            co_catalog, co_warnings = load_co_files(file_scheme_pairs)
            promo_catalog, promo_name, promo_warnings = load_promo_file(promo_path)
            report_rows = build_report(co_catalog, promo_catalog, thresholds)
            self.result_queue.put(("ok", report_rows, promo_name, co_warnings + promo_warnings))
        except Exception as e:  # noqa: BLE001 - показываем пользователю любую ошибку загрузки
            self.result_queue.put(("error", str(e)))

    def _poll_queue(self):
        try:
            item = self.result_queue.get_nowait()
        except queue.Empty:
            self.root.after(100, self._poll_queue)
            return

        self.calc_button.config(state="normal")

        if item[0] == "error":
            self.status_label.config(text="Ошибка расчёта")
            messagebox.showerror("Ошибка", item[1])
            return

        _, report_rows, promo_name, warnings = item
        self.report_rows = report_rows
        self.promo_name = promo_name
        self.export_button.config(state="normal" if report_rows else "disabled")

        counts = Counter(r["status"] for r in report_rows)
        summary = ", ".join(f"{STATUS_LABELS[k]}: {v}" for k, v in counts.items())
        self.status_label.config(text=f"Готово. Всего товаров: {len(report_rows)}. {summary}")

        self._populate_tree()

        if warnings:
            shown = warnings[:30]
            text = "\n".join(shown)
            if len(warnings) > 30:
                text += f"\n... и ещё {len(warnings) - 30} предупреждений"
            messagebox.showwarning("Предупреждения при загрузке", text)

    # ------------------------------------------------------------ Таблица ---

    def _current_filter_key(self):
        label = self.filter_var.get()
        for key, lbl in FILTER_OPTIONS:
            if lbl == label:
                return key
        return "all"

    def on_filter_change(self, _event=None):
        self._populate_tree()

    def _row_matches_filter(self, row, filter_key):
        if filter_key == "all":
            return True
        if filter_key == "no_co":
            return row["status"] in ("no_co", "no_price_data", "error")
        return row["status"] == filter_key

    @staticmethod
    def _fmt(field, value):
        if value is None:
            return ""
        if field == "pct_of_cost":
            return f"{value:.0%}"
        if field in ("price_co", "price_promo_max", "price_used", "net_profit"):
            return f"{value:,.0f}".replace(",", " ")
        if field in ("ordered_qty_7d", "stock_ozon", "stock_own") and isinstance(value, (int, float)):
            return f"{value:,.0f}".replace(",", " ")
        return value

    def _populate_tree(self):
        self.tree.delete(*self.tree.get_children())
        filter_key = self._current_filter_key()
        for row in self.report_rows:
            if not self._row_matches_filter(row, filter_key):
                continue
            values = [self._fmt(key, row.get(key)) for key, _, _ in TREE_COLUMNS]
            tag = row.get("status") or "no_co"
            self.tree.insert("", "end", values=values, tags=(tag,))

    def _sort_by(self, key):
        reverse = self._sort_state.get(key, False)
        self.report_rows.sort(key=lambda r: (r.get(key) is None, r.get(key)), reverse=reverse)
        self._sort_state[key] = not reverse
        self._populate_tree()

    # ------------------------------------------------------------ Экспорт ---

    @staticmethod
    def _sanitize_filename(name):
        for ch in '\\/:*?"<>|':
            name = name.replace(ch, "")
        return name.strip()

    def _default_export_filename(self):
        base = "Отчет по акции"
        if self.promo_name:
            base += f" {self._sanitize_filename(self.promo_name)}"
        return base + ".xlsx"

    def export(self):
        if not self.report_rows:
            messagebox.showinfo("Нет данных", "Сначала выполните расчёт.")
            return
        path = filedialog.asksaveasfilename(
            title="Сохранить отчёт",
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx")],
            initialfile=self._default_export_filename(),
        )
        if not path:
            return
        try:
            export_report(self.report_rows, self.promo_name, self.last_thresholds, path)
        except Exception as e:  # noqa: BLE001
            messagebox.showerror("Ошибка экспорта", str(e))
            return
        messagebox.showinfo("Готово", f"Отчёт сохранён:\n{path}")


def run():
    root = tk.Tk()
    App(root)
    root.mainloop()
