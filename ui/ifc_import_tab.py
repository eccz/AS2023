import tkinter as tk
import openpyxl
from tkinter import messagebox
from ui.file_handler import choose_file
from ifc_import.ifc_import import ifc_import_maker_no_interface
from d_parser.d_parser import parse


def process_file(file_path):
    if not file_path.get():
        messagebox.showwarning("Ошибка", "Выберите файл для обработки")
        return

    try:
        workbook = openpyxl.load_workbook(file_path.get())
        src = parse(workbook, to_term=True)
        ifc_import_maker_no_interface(src)
        messagebox.showinfo('Успех', 'Файл импорта параметров создан!')

    except Exception as e:
        messagebox.showerror("Ошибка", f"Не удалось обработать файл: {e}")


def create_ifc_import_tab(notebook):
    tab = tk.Frame(notebook)
    file_path = tk.StringVar()

    frame_0 = tk.Frame(tab, padx=0, pady=0)
    frame_0.pack(padx=0, pady=10)

    frame = tk.Frame(tab, padx=0, pady=0)
    frame.pack(padx=0, pady=0)

    info_text = (
        "Формирование xml импорта IFC для CadLib на основе доработанного приложения Д."
    )

    tk.Label(frame_0, text=info_text, font=("Arial", 10), justify=tk.LEFT).grid(row=2, column=0, pady=0)

    tk.Label(frame, text="Выберите файл:").grid(row=0, column=0, sticky="w")
    file_entry = tk.Entry(frame, textvariable=file_path, width=60)
    file_entry.grid(row=0, column=1, padx=5, pady=5, sticky='w')
    tk.Button(frame, text="Обзор", command=lambda n=file_path: choose_file(n)).grid(row=0, column=2, padx=5, pady=5,
                                                                                    sticky="w")

    tk.Button(frame, text='Создать файл импорта IFC', command=lambda m=file_path: process_file(m)).grid(row=1,
                                                                                                               column=0,
                                                                                                               columnspan=3,
                                                                                                               pady=2)

    return tab
