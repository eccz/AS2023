import openpyxl
import tkinter as tk
from tkinter import messagebox

from utils.tk_file_utils import choose_file
from src.report_profile.report_profile import xml_report_profile_build, report_profile_xml_output
from src.d_parser import parse
from utils.tk_file_utils import to_xml_tk


def process_file(file_path):
    if not file_path.get():
        messagebox.showwarning("Ошибка", "Выберите файл для обработки")
        return

    try:
        workbook = openpyxl.load_workbook(file_path.get())
        src = parse(workbook, to_term=True)
        res = xml_report_profile_build(src)
        to_xml_tk(res, report_profile_xml_output, filename='report_profile')

    except Exception as e:
        messagebox.showerror("Ошибка", f"Не удалось обработать файл: {e}")


def create_report_profile_tab(notebook):
    tab = tk.Frame(notebook)
    file_path = tk.StringVar()

    frame_0 = tk.Frame(tab, padx=0, pady=0)
    frame_0.pack(padx=0, pady=10)

    frame = tk.Frame(tab, padx=0, pady=0)
    frame.pack(padx=0, pady=0)

    info_text = (
        "Формирование профиля отчета для CadLib на основе доработанного приложения Д."
    )

    tk.Label(frame_0, text=info_text, font=("Arial", 10, 'bold'), wraplength=600).grid(row=2, column=0, pady=0)

    tk.Label(frame, text="Выберите файл:").grid(row=0, column=0, sticky="w")
    file_entry = tk.Entry(frame, textvariable=file_path, width=60)
    file_entry.grid(row=0, column=1, padx=5, pady=5, sticky='w')
    tk.Button(frame, text="Обзор", command=lambda n=file_path: choose_file(n)).grid(row=0, column=2, padx=5, pady=5,
                                                                                    sticky="w")

    tk.Button(frame, text='Создать профиль отчета', command=lambda m=file_path: process_file(m)).grid(row=1,
                                                                                                      column=0,
                                                                                                      columnspan=3,
                                                                                                      pady=2)

    return tab
