import tkinter as tk
from tkinter import filedialog, messagebox
import openpyxl

from report_profile.report_profile import xml_report_profile_build
from d_parser.d_parser import parse


def process_file(file_path):
    if not file_path.get():
        messagebox.showwarning("Ошибка", "Выберите файл для обработки")
        return

    try:
        workbook = openpyxl.load_workbook(file_path.get())
        src = parse(workbook, to_term=True)
        xml_report_profile_build(src)
        messagebox.showinfo('Успех', 'Файл создан!')

    except Exception as e:
        messagebox.showerror("Ошибка", f"Не удалось обработать файл: {e}")


def create_report_profile_tab(notebook):
    def choose_file():
        filename = filedialog.askopenfilename(
            title="!Выберите файл",
            filetypes=(("Файлы excel", "*.xlsx"), ("Файлы excel", "*.xlsm*"))
        )
        file_path.set(filename)

    tab = tk.Frame(notebook)
    file_path = tk.StringVar()
    frame = tk.Frame(notebook, padx=10, pady=10)
    frame.pack(padx=0, pady=50)

    tk.Label(frame, text="Выберите файл:").grid(row=0, column=0, sticky="w")
    file_entry = tk.Entry(frame, textvariable=file_path, width=60)
    file_entry.grid(row=0, column=1, padx=5, pady=5, sticky='w')
    tk.Button(frame, text="Обзор", command=choose_file).grid(row=0, column=2, padx=5, pady=5, sticky="w")

    tk.Button(frame, text='Создать профиль отчета', command=lambda m=file_path: process_file(m)).grid(row=1,
                                                                                                      column=0,
                                                                                                      columnspan=3,
                                                                                                      pady=2)

    return tab
