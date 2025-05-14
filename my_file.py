
import tkinter as tk
from tkinter import messagebox
import pandas as pd
import os

EXCEL_FILE = "Реквизиты_аренда_акт_возврата.xlsx"

fields = [
    "Полное_название", "ФИО_родительный", "ФИО_сокращ", "полный_тип_объекта",
    "Короткое_название", "Юридический_адрес", "Фактический_адрес", "ИНН", "ОГРН",
    "КПП", "Банк", "Кор_счет", "БИК", "ОКПО", "Контактная_персона", "Номер", 
    "Дата_аренды", "Дата"
]

class ExcelEntryApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Добавление данных в Excel")
        self.entries = {}

        for field in fields:
            label = tk.Label(root, text=field.replace("_", " "))
            label.pack()
            entry = tk.Entry(root, width=70)
            entry.pack()
            self.entries[field] = entry

        save_button = tk.Button(root, text="Сохранить в Excel", command=self.save_to_excel, bg="lightgreen")
        save_button.pack(pady=10)

    def save_to_excel(self):
        new_data = {k: v.get() for k, v in self.entries.items()}
        new_row = pd.DataFrame([new_data])

        if os.path.exists(EXCEL_FILE):
            df = pd.read_excel(EXCEL_FILE)
            df = pd.concat([df, new_row], ignore_index=True)
        else:
            df = new_row

        df.to_excel(EXCEL_FILE, index=False)
        messagebox.showinfo("Успех", f"Строка добавлена в файл: {EXCEL_FILE}")
        for entry in self.entries.values():
            entry.delete(0, tk.END)

if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelEntryApp(root)
    root.mainloop()
