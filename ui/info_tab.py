import tkinter as tk

def create_info_tab(notebook):
    tab = tk.Frame(notebook)

    info_text = (
        "Приложение для автоматизации проверок на Atomskills R94 Инженерное проектирование.\n"
        "1. Вкладка 'Работа с файлами' позволяет загружать текстовые файлы.\n"
        "2. Вкладка 'О программе' содержит описание функционала."
    )

    tk.Label(tab, text=info_text, font=("Arial", 10), justify=tk.LEFT).pack(padx=10, pady=10)

    return tab
