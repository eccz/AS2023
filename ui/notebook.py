from tkinter import ttk
from ui.report_profile_tab import create_report_profile_tab
from ui.parameters_tab import create_parameters_tab
from ui.ifc_import_tab import create_ifc_import_tab
from ui.info_tab import create_info_tab

def create_notebook(root):
    notebook = ttk.Notebook(root)

    # Добавление вкладок
    file_tab = create_report_profile_tab(notebook)
    parameters_tab = create_parameters_tab(notebook)
    ifc_import_tab = create_ifc_import_tab(notebook)
    info_tab = create_info_tab(notebook)

    notebook.add(file_tab, text="Профиль отчета")
    notebook.add(parameters_tab, text="Импорт параметров")
    notebook.add(ifc_import_tab, text="Профиль импорта IFC")
    notebook.add(info_tab, text="О программе")

    return notebook
