import tkinter as tk
import openpyxl
from tkinter import messagebox
from ui.file_handler import choose_file
from parameters.parameters import parameters_maker_no_interface
from d_parser.d_parser import parse


def process_file(file_path, entry):
    if not file_path.get():
        messagebox.showwarning("Ошибка", "Выберите файл для обработки")
        return

    try:
        workbook = openpyxl.load_workbook(file_path.get())
        src = parse(workbook, to_term=True)

        param_group_name = entry.get()
        parameters_maker_no_interface(src, param_group_name)
        messagebox.showinfo('Успех', 'Файл импорта параметров создан!')

    except Exception as e:
        messagebox.showerror("Ошибка", f"Не удалось обработать файл: {e}")


def create_parameters_tab(notebook):
    tab = tk.Frame(notebook)
    file_path = tk.StringVar()

    frame_0 = tk.Frame(tab, padx=0, pady=0)
    frame_0.pack(padx=0, pady=10)

    frame_1 = tk.Frame(tab, padx=0, pady=0)
    frame_1.pack(padx=0, pady=0)

    frame_2 = tk.Frame(tab, padx=0, pady=0)
    frame_2.pack(padx=0, pady=0)

    info_text = (
        "Формирование xml импорта параметров для CadLib на основе доработанного приложения Д."
    )

    tk.Label(frame_0, text=info_text, font=("Arial", 10)).grid(row=0, column=0, pady=0)

    tk.Label(frame_1, text="Наименование группы параметров:").grid(row=0, column=0, pady=0)
    param_gr_name_entry = tk.Entry(frame_1, width=40)
    param_gr_name_entry.grid(row=0, column=1, padx=5, pady=5, sticky='w')
    param_gr_name_entry.insert(0, "YS_2025")  # Значение по умолчанию

    tk.Label(frame_2, text="Выберите файл:").grid(row=0, column=0, sticky="w")
    file_entry = tk.Entry(frame_2, textvariable=file_path, width=60)
    file_entry.grid(row=0, column=1, padx=5, pady=5, sticky='w')
    tk.Button(frame_2, text="Обзор", command=lambda n=file_path: choose_file(n)).grid(row=0, column=2, padx=5, pady=5,
                                                                                      sticky="w")
    tk.Button(frame_2,
              text='Создать файл импорта параметров',
              command=lambda m=file_path, n=param_gr_name_entry: process_file(m, n)
              ).grid(row=1,column=0, columnspan=3,pady=2)

    return tab
