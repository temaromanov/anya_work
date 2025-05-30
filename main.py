import tkinter as tk
from tkinter import ttk
from form_ooo import ExcelEntryAppOOO
from form_ip import ExcelEntryAppIP
from uslugi_ooo import ExcelEntryAppUslugiOOO

class MainApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Выберите тип организации")
        self.geometry("820x700")
        self.resizable(False, False)
        self.configure(bg="#f3f6fa")

        # Настроим стили для красоты!
        style = ttk.Style()
        style.theme_use('clam')
        style.configure("TButton", font=("Segoe UI", 12, "bold"), background="#4287f5", foreground="#fff", padding=10)
        style.configure("Menu.TButton", font=("Segoe UI", 13, "bold"), background="#3bb273", foreground="#fff", padding=14)
        style.configure("TLabel", font=("Segoe UI", 12), background="#f3f6fa", foreground="#2255AA")
        style.configure("TNotebook.Tab", font=("Segoe UI", 11, "bold"), padding=[18, 10])

        self.menu_frame = ttk.Frame(self, padding=30, style="TFrame")
        self.menu_frame.pack(expand=True, fill="both")

        label = ttk.Label(self.menu_frame, text="Выберите тип организации", font=("Segoe UI", 17, "bold"))
        label.pack(pady=(38,18))

        ttk.Button(self.menu_frame, text="ООО", width=18, style="Menu.TButton", command=self.show_ooo).pack(pady=10)
        ttk.Button(self.menu_frame, text="ИП", width=18, style="Menu.TButton", command=self.show_ip).pack(pady=10)

        # ---- Вот сюда вставляем кнопку Оказание услуг (ООО):
        ttk.Button(
            self.menu_frame,
            text="Оказание услуг (ООО)",
            width=22,
            style="Menu.TButton",
            command=self.show_uslugi_ooo
        ).pack(pady=10)
        # ---------------------------------------

        self.ooo_frame = None
        self.ip_frame = None
        self.uslugi_ooo_frame = None

    def show_ooo(self):
        self.menu_frame.pack_forget()
        if self.ooo_frame is None:
            self.ooo_frame = ttk.Frame(self, padding=12)
            ExcelEntryAppOOO(self.ooo_frame, go_back=self.back_to_menu)
        self.ooo_frame.pack(expand=True, fill="both")

    def show_uslugi_ooo(self):
        self.menu_frame.pack_forget()
        if self.uslugi_ooo_frame is None:
            self.uslugi_ooo_frame = ttk.Frame(self, padding=12)
            ExcelEntryAppUslugiOOO(self.uslugi_ooo_frame, go_back=self.back_to_menu)
        self.uslugi_ooo_frame.pack(expand=True, fill="both")

    def show_ip(self):
        self.menu_frame.pack_forget()
        if self.ip_frame is None:
            self.ip_frame = ttk.Frame(self, padding=12)
            ExcelEntryAppIP(self.ip_frame, go_back=self.back_to_menu)
        self.ip_frame.pack(expand=True, fill="both")

    def back_to_menu(self):
        if self.ooo_frame: self.ooo_frame.pack_forget()
        if self.ip_frame: self.ip_frame.pack_forget()
        if self.uslugi_ooo_frame: self.uslugi_ooo_frame.pack_forget()
        self.menu_frame.pack(expand=True, fill="both")

if __name__ == "__main__":
    MainApp().mainloop()

