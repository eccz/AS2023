import openpyxl
import tkinter as tk
from tkinter import messagebox

from ui.tk_file_utils import choose_file
from ifc_export.ifc_export_full import ifc_export_xml_build, ifc_export_xml_output
from d_parser.d_parser import parse
from ui.tk_file_utils import to_xml_tk


def process_file(file_path, entry, choice_mapping, choice_class):
    if not file_path.get():
        messagebox.showwarning("Ошибка", "Выберите файл для обработки")
        return

    try:
        res = None
        workbook = openpyxl.load_workbook(file_path.get())
        src = parse(workbook, to_term=False)

        pset_name = entry.get()
        ch1, ch2 = True, True
        if choice_mapping.get() == 'one':
            ch1 = True
        elif choice_mapping.get() == 'two':
            ch1 = False
        if choice_class.get() == 'one':
            ch2 = True
        if choice_class.get() == 'two':
            ch2 = False

        res = ifc_export_xml_build(pset_name, src, ms_mapping=ch1, class2025=ch2)
        to_xml_tk(res, ifc_export_xml_output, filename='ifc_export')

    except Exception as e:
        messagebox.showerror("Ошибка", f"Не удалось обработать файл: {e}")


def create_ifc_export_tab(notebook):
    tab = tk.Frame(notebook)
    file_path = tk.StringVar()

    frame_0 = tk.Frame(tab, padx=0, pady=0)
    frame_0.pack(padx=0, pady=10)

    frame_1 = tk.Frame(tab, padx=0, pady=0)
    frame_1.pack(padx=0, pady=0)

    frame_2 = tk.Frame(tab, padx=0, pady=0)
    frame_2.pack(padx=0, pady=0)

    frame_3 = tk.Frame(tab, padx=0, pady=0)
    frame_3.pack(padx=0, pady=0)
    choice_mapping = tk.StringVar(value="one")
    choice_class = tk.StringVar(value="one")

    info_text = (
        "Формирование xml-профиля экспорта IFC для CadLib на основе доработанного приложения Д."
    )
    tk.Label(frame_0, text=info_text, font=("Arial", 10, 'bold'), wraplength=600).grid(row=2, column=0, pady=0)

    tk.Label(frame_1, text="Наименование property_set:").grid(row=0, column=0, pady=0)
    pset_name_entry = tk.Entry(frame_1, width=40)
    pset_name_entry.grid(row=0, column=1, padx=5, pady=5, sticky='w')
    pset_name_entry.insert(0, "EngeneeringDesign")  # Значение по умолчанию

    tk.Label(frame_2, text="Выберите файл:").grid(row=0, column=0, sticky="w")
    file_entry = tk.Entry(frame_2, textvariable=file_path, width=60)
    file_entry.grid(row=0, column=1, padx=5, pady=5, sticky='w')
    tk.Button(frame_2, text="Обзор", command=lambda n=file_path: choose_file(n)).grid(row=0, column=2, padx=5, pady=5,
                                                                                      sticky="w")

    tk.Button(frame_2, text='Создать профиль экспорта IFC',
              command=lambda m=file_path,
                             n=pset_name_entry,
                             q=choice_mapping,
                             d=choice_class:
                        process_file(m, n, q, d)).grid(row=1,column=0,columnspan=3, pady=2)

    tk.Label(frame_3, text="Выберите опцию по использованию столбца маппинга с атрибутами MS (по умолчанию 'True'):", font=("Arial", 8, 'bold')).grid(row=1, column=0, sticky="w")
    tk.Radiobutton(frame_3, text="Использовать столбец маппинга MS", variable=choice_mapping, value="one").grid(row=2, column=0, pady=0, sticky="w")
    tk.Radiobutton(frame_3, text="Не использовать столбец маппинга MS", variable=choice_mapping, value="two").grid(row=3, column=0, pady=0, sticky="w")
    tk.Label(frame_3, text="Выберите опцию по использованию классификатора AS2025(по умолчанию 'True'):", font=("Arial", 8, 'bold')).grid(row=4, column=0, sticky="w")
    tk.Radiobutton(frame_3, text="Использовать классификатор AS2025", variable=choice_class, value="one").grid(row=5, column=0, pady=0, sticky="w")
    tk.Radiobutton(frame_3, text="Не использовать классификатор AS2025", variable=choice_class, value="two").grid(row=6, column=0, pady=0, sticky="w")

    return tab
