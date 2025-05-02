import openpyxl
import tkinter as tk
from tkinter import messagebox, filedialog, ttk

from utils.tk_file_utils import choose_file
from src.specificator.spec import specificator_no_interface, specificator_xml_output
from src.d_parser import parse
import os


def process_file(file_path: tk.StringVar, type_entry: tk.Entry, spec_entry: tk.Entry, choice_mapping: tk.StringVar) -> None:
    """
    Обрабатывает выбранный Excel-файл и создает XML-файлы для Спецификатора.
    """
    if not file_path.get():
        messagebox.showwarning("Ошибка", "Выберите файл для обработки")
        return

    try:
        workbook = openpyxl.load_workbook(file_path.get())
        src = parse(workbook, to_term=False)

        type_param_name = type_entry.get()
        spec_param_name = spec_entry.get()

        ch1 = choice_mapping.get() == 'one'

        res = specificator_no_interface(src, type_param_name, spec_param_name, ms_mapping=ch1)

        directory_path = filedialog.askdirectory(title="Выберите папку для сохранения файлов")
        if not directory_path:
            messagebox.showwarning("Внимание", "Папка не выбрана. Операция отменена.")
            return

        try:
            for name, data in res.items():
                output_path = os.path.join(directory_path, f'SPEC_{name}.xml')
                specificator_xml_output(data, output_path)

            messagebox.showinfo("Успех", "Файлы успешно сохранены!")
        except Exception as e:
            messagebox.showerror("Ошибка при сохранении файлов", str(e))


    except Exception as e:
        messagebox.showerror("Ошибка обработки файла", str(e))


def create_specificator_tab(notebook: ttk.Notebook) -> tk.Frame:
    """
    Создает вкладку в интерфейсе для создания XML-файлов Спецификатора.
    """
    tab = tk.Frame(notebook)
    file_path = tk.StringVar()
    choice_mapping = tk.StringVar(value="one")

    # --- Информационный блок ---
    frame_info = tk.Frame(tab)
    frame_info.pack(pady=10)

    info_text = (
        "Формирование xml-профилей для Спецификатора ModelStudioCS на основе доработанного приложения Д."
    )
    tk.Label(frame_info, text=info_text, font=("Arial", 10, 'bold'), wraplength=600).pack()

    # --- Поле фильтрации по типу ---
    frame_type = tk.Frame(tab)
    frame_type.pack()

    tk.Label(frame_type, text="Наименование параметра фильтрации по типу:").grid(row=0, column=0, sticky="w")
    type_name_entry = tk.Entry(frame_type, width=40)
    type_name_entry.grid(row=0, column=1, padx=5, pady=5, sticky="w")
    type_name_entry.insert(0, "TYPE")

    # --- Поле фильтрации по специализации ---
    frame_spec = tk.Frame(tab)
    frame_spec.pack()

    tk.Label(frame_spec, text="Наименование параметра фильтрации по специализации:").grid(row=0, column=0, sticky="w")
    spec_name_entry = tk.Entry(frame_spec, width=40)
    spec_name_entry.grid(row=0, column=1, padx=5, pady=5, sticky="w")
    spec_name_entry.insert(0, "SPECIALITY")

    # --- Выбор файла приложения Д ---
    frame_file = tk.Frame(tab)
    frame_file.pack()

    tk.Label(frame_file, text="Выберите файл приложения Д:").grid(row=0, column=0, sticky="w")
    file_entry = tk.Entry(frame_file, textvariable=file_path, width=60)
    file_entry.grid(row=0, column=1, padx=5, pady=5, sticky="w")
    tk.Button(frame_file, text="Обзор", command=lambda: choose_file(file_path)).grid(row=0, column=2, padx=5, pady=5)

    # --- Кнопка обработки файла ---
    tk.Button(frame_file, text='Создать xml-файлы для Спецификатора',
              command=lambda: process_file(file_path, type_name_entry, spec_name_entry, choice_mapping)
             ).grid(row=1, column=0, columnspan=3, pady=10)

    # --- Выбор использования маппинга ---
    frame_mapping = tk.Frame(tab, padx=0, pady=0)
    frame_mapping.pack(padx=0, pady=0)

    tk.Label(frame_mapping, text="Выберите опцию по использованию столбца маппинга с атрибутами MS (по умолчанию 'True'):", font=("Arial", 8, 'bold')).grid(row=0, column=0, pady=0, sticky="w")
    tk.Radiobutton(frame_mapping, text="Использовать столбец маппинга MS", variable=choice_mapping, value="one").grid(row=1, column=0, pady=0, sticky="w")
    tk.Radiobutton(frame_mapping, text="Не использовать столбец маппинга MS", variable=choice_mapping, value="two").grid(row=2, column=0, pady=0, sticky="w")
    return tab
