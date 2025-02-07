import tkinter as tk

def create_info_tab(notebook):
    tab = tk.Frame(notebook)

    info_text = (
        "Приложение для автоматизации проверок на Atomskills R94 Инженерное проектирование.\n"
        "1. На вкладке 'Профиль отчета' создается профиль отчета для проверки в CadLib на основе приложения D.\n"
        "2. Вкладка 'О программе' содержит описание функционала."
    )

    tk.Label(tab, text=info_text, font=("Arial", 10), justify=tk.LEFT, wraplength=600).pack(padx=10, pady=10)

    return tab
