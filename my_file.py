
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from tkcalendar import DateEntry
import pandas as pd
import os
from docx import Document

class ExcelEntryApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Форма заполнения реквизитов")
        self.root.geometry("750x600")
        self.entries = {}
        self.excel_file = None
        self.word_template = "Шаблон_договора.docx"

        # Вкладки
        notebook = ttk.Notebook(root)
        notebook.pack(expand=True, fill='both', padx=10, pady=10)

        self.frames = {
            "Основное": ttk.Frame(notebook),
            "Банк и документы": ttk.Frame(notebook),
            "Даты и номер": ttk.Frame(notebook)
        }

        for name, frame in self.frames.items():
            notebook.add(frame, text=name)

        self.fields_main = [
            "Полное_название", "ФИО_родительный", "ФИО_сокращ",
            "полный_тип_объекта", "Короткое_название",
            "Юридический_адрес", "Фактический_адрес"
        ]
        self.fields_bank = [
            "ИНН", "ОГРН", "КПП", "Банк", "Кор_счет", "БИК", "ОКПО", "Контактная_персона"
        ]
        self.fields_date = [
            "Номер", "Дата_аренды", "Дата"
        ]
        bank_options = ["ПАО СБЕРБАНК", "ВТБ", "Газпромбанк", "Альфа-Банк", "Тинькофф"]
        person_options = ["Генеральный директор", "Помощник", "Ответственный менеджер"]

        for field in self.fields_main:
            self.create_entry(self.frames["Основное"], field)

        for field in self.fields_bank:
            if field == "Банк":
                self.create_combobox(self.frames["Банк и документы"], field, bank_options)
            elif field == "Контактная_персона":
                self.create_combobox(self.frames["Банк и документы"], field, person_options)
            else:
                self.create_entry(self.frames["Банк и документы"], field)

        for field in self.fields_date:
            if "Дата" in field:
                self.create_dateentry(self.frames["Даты и номер"], field)
            else:
                self.create_entry(self.frames["Даты и номер"], field)

        button_frame = ttk.Frame(root)
        button_frame.pack(pady=10)

        ttk.Button(button_frame, text="📂 Выбрать Excel-файл", command=self.choose_file).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="💾 Сохранить как...", command=self.save_as).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="✅ Добавить и создать Word", command=self.save_and_generate_word).pack(side=tk.LEFT, padx=5)

    def create_entry(self, parent, field):
        label = ttk.Label(parent, text=field.replace("_", " "))
        label.pack(anchor='w', padx=10)
        entry = ttk.Entry(parent, width=85)
        entry.pack(padx=10, pady=3)
        self.entries[field] = entry

    def create_combobox(self, parent, field, options):
        label = ttk.Label(parent, text=field.replace("_", " "))
        label.pack(anchor='w', padx=10)
        combo = ttk.Combobox(parent, values=options, width=83)
        combo.pack(padx=10, pady=3)
        self.entries[field] = combo

    def create_dateentry(self, parent, field):
        label = ttk.Label(parent, text=field.replace("_", " "))
        label.pack(anchor='w', padx=10)
        date_entry = DateEntry(parent, width=83, date_pattern="dd.mm.yyyy")
        date_entry.pack(padx=10, pady=3)
        self.entries[field] = date_entry

    def choose_file(self):
        path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if path:
            self.excel_file = path

    def save_as(self):
        path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if path:
            self.excel_file = path

    def save_and_generate_word(self):
        if not self.excel_file:
            messagebox.showwarning("Файл не выбран", "Пожалуйста, выберите или создайте Excel-файл перед сохранением.")
            return

        new_data = {k: v.get() for k, v in self.entries.items()}
        new_row = pd.DataFrame([new_data])

        if os.path.exists(self.excel_file):
            try:
                df = pd.read_excel(self.excel_file)
                df = pd.concat([df, new_row], ignore_index=True)
            except:
                df = new_row
        else:
            df = new_row

        df.to_excel(self.excel_file, index=False)

        if os.path.exists(self.word_template):
            doc = Document(self.word_template)
            for p in doc.paragraphs:
                for key, val in new_data.items():
                    if f"{{{{{key}}}}}" in p.text:
                        p.text = p.text.replace(f"{{{{{key}}}}}", val)

            word_output = self.excel_file.replace(".xlsx", "_документ.docx")
            doc.save(word_output)
            os.startfile(word_output)

        for entry in self.entries.values():
            entry.delete(0, tk.END)

if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelEntryApp(root)
    root.mainloop()
