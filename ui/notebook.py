from tkinter import ttk
from ui.tabs.report_profile_tab import create_report_profile_tab
from ui.tabs.parameters_tab import create_parameters_tab
from ui.tabs.ifc_import_tab import create_ifc_import_tab
from ui.tabs.ifc_export_tab import create_ifc_export_tab
from ui.tabs.info_tab import create_info_tab
from ui.tabs.specificator_tab import create_specificator_tab

def create_notebook(root):
    notebook = ttk.Notebook(root)

    # Добавление вкладок
    file_tab = create_report_profile_tab(notebook)
    parameters_tab = create_parameters_tab(notebook)
    ifc_import_tab = create_ifc_import_tab(notebook)
    ifc_export_tab = create_ifc_export_tab(notebook)
    specificator_tab = create_specificator_tab(notebook)
    info_tab = create_info_tab(notebook)

    notebook.add(file_tab, text="Профиль отчета")
    notebook.add(parameters_tab, text="Импорт параметров")
    notebook.add(ifc_import_tab, text="Профиль импорта IFC")
    notebook.add(ifc_export_tab, text="Профиль экспорта IFC")
    notebook.add(specificator_tab, text="Спецификатор")
    notebook.add(info_tab, text="О программе")

    return notebook
