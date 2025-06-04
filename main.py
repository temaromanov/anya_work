import tkinter as tk
from tkinter import ttk
from form_ooo import ExcelEntryAppOOO
from form_ip import ExcelEntryAppIP
from uslugi_ooo import ExcelEntryAppUslugiOOO
from uslugi_ip import ExcelEntryAppUslugiIP
from act_vozvrata_ooo import ExcelEntryAppActVozvrataOOO
from peredatochn_act_ooo import ExcelPeredActAppOOO

class MainApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Выберите тип организации")
        self.geometry("820x700")
        self.resizable(False, False)
        self.configure(bg="#f3f6fa")

        # Стили
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

        ttk.Button(self.menu_frame, text="Договор аренды (ООО)", width=22, style="Menu.TButton", command=self.show_ooo).pack(pady=8)
        ttk.Button(self.menu_frame, text="Договор аренды (ИП)", width=22, style="Menu.TButton", command=self.show_ip).pack(pady=8)
        ttk.Button(self.menu_frame, text="Оказание услуг (ООО)", width=22, style="Menu.TButton", command=self.show_uslugi_ooo).pack(pady=8)
        ttk.Button(self.menu_frame, text="Оказание услуг (ИП)", width=22, style="Menu.TButton", command=self.show_uslugi_ip).pack(pady=8)
        ttk.Button(self.menu_frame, text="Акт возврата (ООО)", width=22, style="Menu.TButton", command=self.show_act_vozvrata_ooo).pack(pady=8)
        ttk.Button(self.menu_frame, text="Передаточный акт (ООО)", width=22, style="Menu.TButton", command=self.show_peredatoch_act_ooo).pack(pady=8)

        self.ooo_frame = None
        self.ip_frame = None
        self.uslugi_ooo_frame = None
        self.uslugi_ip_frame = None
        self.act_vozvrata_frame = None
        self.peredatoch_act_frame = None
        

    def show_ooo(self):
        self.menu_frame.pack_forget()
        if self.ooo_frame is None:
            self.ooo_frame = ttk.Frame(self, padding=12)
            ExcelEntryAppOOO(self.ooo_frame, go_back=self.back_to_menu)
        self.ooo_frame.pack(expand=True, fill="both")

    def show_ip(self):
        self.menu_frame.pack_forget()
        if self.ip_frame is None:
            self.ip_frame = ttk.Frame(self, padding=12)
            ExcelEntryAppIP(self.ip_frame, go_back=self.back_to_menu)
        self.ip_frame.pack(expand=True, fill="both")

    def show_uslugi_ooo(self):
        self.menu_frame.pack_forget()
        if self.uslugi_ooo_frame is None:
            self.uslugi_ooo_frame = ttk.Frame(self, padding=12)
            ExcelEntryAppUslugiOOO(self.uslugi_ooo_frame, go_back=self.back_to_menu)
        self.uslugi_ooo_frame.pack(expand=True, fill="both")

    def show_uslugi_ip(self):
        self.menu_frame.pack_forget()
        if self.uslugi_ip_frame is None:
            self.uslugi_ip_frame = ttk.Frame(self, padding=12)
            ExcelEntryAppUslugiIP(self.uslugi_ip_frame, go_back=self.back_to_menu)
        self.uslugi_ip_frame.pack(expand=True, fill="both")

    def show_act_vozvrata_ooo(self):
        self.menu_frame.pack_forget()
        if self.act_vozvrata_frame is None:
            self.act_vozvrata_frame = ttk.Frame(self, padding=12)
        ExcelEntryAppActVozvrataOOO(self.act_vozvrata_frame, go_back=self.back_to_menu)
        self.act_vozvrata_frame.pack(expand=True, fill="both")

    def show_peredatoch_act_ooo(self):
        self.menu_frame.pack_forget()
        if self.peredatoch_act_frame is None:
            self.peredatoch_act_frame = ttk.Frame(self, padding=12)
        ExcelPeredActAppOOO(self.peredatoch_act_frame, go_back=self.back_to_menu)
        self.peredatoch_act_frame.pack(expand=True, fill="both")

    def back_to_menu(self):
        for frame in [
            self.ooo_frame, self.ip_frame, self.uslugi_ooo_frame,
            self.uslugi_ip_frame, self.act_vozvrata_frame, self.peredatoch_act_frame
        ]:
            if frame: frame.pack_forget()
        self.menu_frame.pack(expand=True, fill="both")

if __name__ == "__main__":
    MainApp().mainloop()


