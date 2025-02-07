from tkinter import ttk
from ui.file_tab import create_report_profile_tab

from ui.info_tab import create_info_tab

def create_notebook(root):
    notebook = ttk.Notebook(root)

    # Добавление вкладок
    file_tab = create_report_profile_tab(notebook)
    info_tab = create_info_tab(notebook)

    notebook.add(file_tab, text="Профиль отчета")
    notebook.add(info_tab, text="О программе")

    return notebook
