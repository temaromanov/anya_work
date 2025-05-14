
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from tkcalendar import DateEntry
import pandas as pd
import os

class ExcelEntryApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Добавление данных в Excel")
        self.entries = {}
        self.excel_file = None

        # Поля формы
        fields = [
            "Полное_название", "ФИО_родительный", "ФИО_сокращ", "полный_тип_объекта",
            "Короткое_название", "Юридический_адрес", "Фактический_адрес", "ИНН", "ОГРН",
            "КПП", "Банк", "Кор_счет", "БИК", "ОКПО", "Контактная_персона", "Номер", 
            "Дата_аренды", "Дата"
        ]

        bank_options = ["ПАО СБЕРБАНК", "ВТБ", "Газпромбанк", "Альфа-Банк", "Тинькофф"]
        person_options = ["Генеральный директор", "Помощник", "Ответственный менеджер"]

        for field in fields:
            label = tk.Label(root, text=field.replace("_", " "))
            label.pack()

            if field == "Банк":
                entry = ttk.Combobox(root, values=bank_options, width=67)
            elif field == "Контактная_персона":
                entry = ttk.Combobox(root, values=person_options, width=67)
            elif "Дата" in field:
                entry = DateEntry(root, width=67, date_pattern="dd.mm.yyyy")
            else:
                entry = tk.Entry(root, width=70)

            entry.pack()
            self.entries[field] = entry

        # Кнопки управления
        btn_frame = tk.Frame(root)
        btn_frame.pack(pady=10)

        tk.Button(btn_frame, text="📂 Выбрать Excel-файл", command=self.choose_file).pack(side=tk.LEFT, padx=5)
        tk.Button(btn_frame, text="💾 Сохранить как...", command=self.save_as).pack(side=tk.LEFT, padx=5)
        tk.Button(btn_frame, text="✅ Добавить запись", command=self.save_to_excel, bg="lightgreen").pack(side=tk.LEFT, padx=5)

    def choose_file(self):
        path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if path:
            self.excel_file = path
            messagebox.showinfo("Файл выбран", f"Текущий файл:{self.excel_file}")

    def save_as(self):
        path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if path:
            self.excel_file = path
            messagebox.showinfo("Файл сохранения выбран", f"Будет сохранено в:{self.excel_file}")

    def save_to_excel(self):
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
        os.startfile(self.excel_file)
        messagebox.showinfo("Готово", f"Данные добавлены в файл:{self.excel_file}")

        for entry in self.entries.values():
            entry.delete(0, tk.END)

if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelEntryApp(root)
    root.mainloop()
