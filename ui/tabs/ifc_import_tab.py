import openpyxl
import tkinter as tk
from tkinter import messagebox

from utils.tk_file_utils import choose_file
from src.ifc_import.ifc_import import ifc_import_maker_no_interface_d,ifc_import_maker_no_interface_full, ifc_import_xml_output
from src.d_parser import parse
from utils.tk_file_utils import to_xml_tk


def process_file(file_path, entry, choice):
    if not file_path.get():
        messagebox.showwarning("Ошибка", "Выберите файл для обработки")
        return

    try:
        res = None
        workbook = openpyxl.load_workbook(file_path.get())
        src = parse(workbook, to_term=True)

        pset_name = entry.get()
        if choice.get() == "one":
            res = ifc_import_maker_no_interface_d(src, property_set=pset_name)
        if choice.get() == "two":
            res = ifc_import_maker_no_interface_full(src, property_set=pset_name, pset_mapping=False)
        if choice.get() == "three":
            res = ifc_import_maker_no_interface_full(src, property_set=pset_name, pset_mapping=True)

        to_xml_tk(res, ifc_import_xml_output, filename='ifc_import')

    except Exception as e:
        messagebox.showerror("Ошибка", f"Не удалось обработать файл: {e}")


def create_ifc_import_tab(notebook):
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
    choice = tk.StringVar(value="one")

    info_text = (
        "Формирование xml-профиля импорта IFC для CadLib на основе доработанного приложения Д."
    )
    tk.Label(frame_0, text=info_text, font=("Arial", 10, 'bold'), wraplength=600).grid(row=2, column=0, pady=0)

    tk.Label(frame_1, text="Наименование property_set:").grid(row=0, column=0, pady=0)
    pset_name_entry = tk.Entry(frame_1, width=40)
    pset_name_entry.grid(row=0, column=1, padx=5, pady=5, sticky='w')
    pset_name_entry.insert(0, "EngeneeringDesign")  # Значение по умолчанию

    tk.Label(frame_2, text="Выберите файл:").grid(row=0, column=0, sticky="w")
    file_entry = tk.Entry(frame_2, textvariable=file_path, width=60)
    file_entry.grid(row=0, column=1, padx=5, pady=5, sticky='w')
    tk.Button(frame_2,
              text="Обзор",
              command=lambda n=file_path:
              choose_file(n)).grid(row=0, column=2, padx=5, pady=5, sticky="w")

    tk.Button(frame_2,
              text='Создать профиль импорта IFC',
              command=lambda m=file_path, n=pset_name_entry, q=choice:
              process_file(m, n, q)).grid(row=1, column=0, columnspan=3, pady=2)

    tk.Label(frame_3, text="Выберите опцию (по умолчанию 'На основе элементов'):", font=("Arial", 8, 'bold')).grid(row=0, column=0, sticky="w")
    tk.Radiobutton(frame_3, text="На основе элементов", variable=choice, value="one").grid(row=1, column=0, pady=0, sticky="w")
    tk.Radiobutton(frame_3, text="На основе листа Attributes", variable=choice, value="two").grid(row=2, column=0, pady=0, sticky="w")
    tk.Radiobutton(frame_3, text="На основе листа Attributes c учетом маппинга property set-ов", variable=choice, value="three").grid(row=3, column=0, pady=0, sticky="w")

    return tab
