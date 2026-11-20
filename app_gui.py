#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Конвертер Excel у формат АВК-5
Графічна програма для автоматичної підготовки складських відомостей та прайсів для АВК-5.
"""

import os
import sys
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import converter

class AVKConverterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Конвертер матеріалів в АВК-5")
        self.root.geometry("820x640")
        self.root.minsize(700, 500)

        # Apply standard style
        self.style = ttk.Style()
        self.style.theme_use('clam')
        
        self.current_filepath = None
        self.parsed_items = []
        
        self.setup_ui()

    def setup_ui(self):
        # Top banner
        header_frame = tk.Frame(self.root, bg="#1F497D", padx=15, pady=12)
        header_frame.pack(fill=tk.X)
        
        title_label = tk.Label(
            header_frame, 
            text="Автоматична підготовка матеріалів для АВК-5", 
            font=("Arial", 14, "bold"), 
            fg="white", 
            bg="#1F497D"
        )
        title_label.pack(anchor="w")
        
        subtitle_label = tk.Label(
            header_frame, 
            text="Перетворює складські відомості з Excel/CSV у формат для імпорту в Локальний кошторис", 
            font=("Arial", 9), 
            fg="#DCE6F1", 
            bg="#1F497D"
        )
        subtitle_label.pack(anchor="w")

        # Main container
        main_frame = ttk.Frame(self.root, padding="15")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # File selection box
        file_box = ttk.LabelFrame(main_frame, text=" 1. Оберіть файл Excel або CSV ", padding="10")
        file_box.pack(fill=tk.X, pady=(0, 10))

        self.file_label = ttk.Label(file_box, text="Файл не вибрано", font=("Arial", 10))
        self.file_label.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(5, 10))

        choose_btn = ttk.Button(file_box, text="Огляд файлу...", command=self.choose_file)
        choose_btn.pack(side=tk.RIGHT)

        # Options box
        options_box = ttk.LabelFrame(main_frame, text=" 2. Налаштування шифру АВК ", padding="10")
        options_box.pack(fill=tk.X, pady=(0, 10))

        ttk.Label(options_box, text="Префікс шифру матеріалу:").pack(side=tk.LEFT, padx=(5, 5))
        self.prefix_var = tk.StringVar(value="С1-ПР1-")
        prefix_entry = ttk.Entry(options_box, textvariable=self.prefix_var, width=15)
        prefix_entry.pack(side=tk.LEFT, padx=(0, 15))

        ttk.Label(options_box, text="(Наприклад, С1-ПР1-1, С1-ПР1-2... як у кошторисі)").pack(side=tk.LEFT)

        # Statistics & Status
        self.status_var = tk.StringVar(value="Оберіть файл для початку обробки.")
        status_label = ttk.Label(main_frame, textvariable=self.status_var, font=("Arial", 10, "bold"), foreground="#1F497D")
        status_label.pack(anchor="w", pady=(0, 5))

        # Preview table
        table_frame = ttk.Frame(main_frame)
        table_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))

        columns = ("cipher", "name", "unit", "qty", "price", "sum")
        self.tree = ttk.Treeview(table_frame, columns=columns, show="headings", height=10)

        self.tree.heading("cipher", text="Шифр")
        self.tree.heading("name", text="Найменування товару")
        self.tree.heading("unit", text="Од.")
        self.tree.heading("qty", text="Кіл-сть")
        self.tree.heading("price", text="Ціна, грн")
        self.tree.heading("sum", text="Сума, грн")

        self.tree.column("cipher", width=85, anchor="center")
        self.tree.column("name", width=340, anchor="w")
        self.tree.column("unit", width=55, anchor="center")
        self.tree.column("qty", width=75, anchor="e")
        self.tree.column("price", width=85, anchor="e")
        self.tree.column("sum", width=95, anchor="e")

        scrollbar = ttk.Scrollbar(table_frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscroll=scrollbar.set)

        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # Action buttons
        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(fill=tk.X, pady=(5, 0))

        self.save_btn = ttk.Button(
            btn_frame, 
            text=" 💾 Зберегти готовий файл для АВК-5 (.xlsx) ", 
            command=self.save_for_avk
        )
        self.save_btn.pack(side=tk.LEFT, padx=(0, 10))
        self.save_btn.config(state=tk.DISABLED)

        self.save_rasch_btn = ttk.Button(
            btn_frame, 
            text=" 💾 Зберегти під Rasch_290 (.xlsx) ", 
            command=self.save_for_rasch
        )
        self.save_rasch_btn.pack(side=tk.LEFT)
        self.save_rasch_btn.config(state=tk.DISABLED)

    def choose_file(self):
        filetypes = [
            ("Таблиці Excel та CSV", "*.xlsx;*.xls;*.csv"),
            ("Файли Excel (.xlsx)", "*.xlsx"),
            ("Файли Excel 97-2003 (.xls)", "*.xls"),
            ("Файли CSV (.csv)", "*.csv"),
            ("Усі файли", "*.*")
        ]
        path = filedialog.askopenfilename(title="Оберіть файл зі списком товарів", filetypes=filetypes)
        if not path:
            return

        self.current_filepath = path
        self.file_label.config(text=os.path.basename(path))
        self.process_file(path)

    def process_file(self, filepath):
        try:
            self.parsed_items = converter.read_file(filepath)
            
            # Clear previous table
            for row in self.tree.get_children():
                self.tree.delete(row)

            total_sum = sum(item['sum'] for item in self.parsed_items)
            prefix = self.prefix_var.get().strip()

            for idx, item in enumerate(self.parsed_items, 1):
                cipher = f"{prefix}{idx}" if prefix else str(idx)
                self.tree.insert("", tk.END, values=(
                    cipher,
                    item['name'],
                    item['unit'],
                    f"{item['qty']:g}",
                    f"{item['price']:.2f}",
                    f"{item['sum']:.2f}"
                ))

            self.status_var.set(f" Успішно розпізнано позицій: {len(self.parsed_items)} | Загальна сума: {total_sum:,.2f} грн".replace(',', ' '))
            self.save_btn.config(state=tk.NORMAL)
            self.save_rasch_btn.config(state=tk.NORMAL)

        except Exception as e:
            messagebox.showerror("Помилка обробки файлу", f"Не вдалося зчитати файл:\n{e}")
            self.status_var.set("Помилка під час зчитування.")
            self.save_btn.config(state=tk.DISABLED)
            self.save_rasch_btn.config(state=tk.DISABLED)

    def save_for_avk(self):
        if not self.parsed_items:
            return
        
        default_name = "АВК_Матеріали.xlsx"
        path = filedialog.asksaveasfilename(
            title="Зберегти файл для АВК-5",
            defaultextension=".xlsx",
            initialfile=default_name,
            filetypes=[("Книга Excel", "*.xlsx")]
        )
        if not path:
            return

        try:
            prefix = self.prefix_var.get().strip()
            converter.export_to_avk_excel(self.parsed_items, path, cipher_prefix=prefix)
            messagebox.showinfo("Успіх", f"Файл успішно збережено!\n{path}\n\nТепер в АВК-5 відкрийте Локальний кошторис та оберіть 'Введення інформації з файлів Excel'.")
        except Exception as e:
            messagebox.showerror("Помилка збереження", f"Не вдалося зберегти файл:\n{e}")

    def save_for_rasch(self):
        if not self.parsed_items:
            return
            
        import openpyxl
        default_name = "Для_Rasch_290.xlsx"
        path = filedialog.asksaveasfilename(
            title="Зберегти файл під Rasch_290",
            defaultextension=".xlsx",
            initialfile=default_name,
            filetypes=[("Книга Excel", "*.xlsx")]
        )
        if not path:
            return

        try:
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Лист1"
            # Headers that Rasch_290 expects
            ws.append(["Шифр", "Наименование", "Ед. измерения", "Количество", "Материалы"])
            prefix = self.prefix_var.get().strip()
            for idx, item in enumerate(self.parsed_items, 1):
                cipher = f"{prefix}{idx}" if prefix else str(idx)
                ws.append([cipher, item['name'], item['unit'], item['qty'], item['price']])
            wb.save(path)
            messagebox.showinfo("Успіх", f"Файл збережено під вимоги Rasch_290!\n{path}")
        except Exception as e:
            messagebox.showerror("Помилка збереження", f"Не вдалося зберегти файл:\n{e}")

def main():
    root = tk.Tk()
    app = AVKConverterApp(root)
    root.mainloop()

if __name__ == '__main__':
    main()
