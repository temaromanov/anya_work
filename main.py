import tkinter as tk
from form_ooo import ExcelEntryAppOOO
from form_ip import ExcelEntryAppIP

class StartMenu(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Выберите тип организации")
        self.geometry("320x200")
        self.resizable(False, False)
        label = tk.Label(self, text="Выберите организацию", font=("Segoe UI", 14, "bold"))
        label.pack(pady=24)
        tk.Button(self, text="ООО", width=15, height=2, command=self.open_ooo).pack(pady=8)
        tk.Button(self, text="ИП", width=15, height=2, command=self.open_ip).pack(pady=8)

    def open_ooo(self):
        self.destroy()
        root = tk.Tk()
        ExcelEntryAppOOO(root)
        root.mainloop()

    def open_ip(self):
        self.destroy()
        root = tk.Tk()
        ExcelEntryAppIP(root)
        root.mainloop()

if __name__ == "__main__":
    StartMenu().mainloop()
