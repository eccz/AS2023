from tkinter import filedialog, messagebox
import tkinter as tk
import openpyxl

def load_excel_file():
    """Открывает диалог для выбора файла и возвращает его содержимое."""
    file_path = filedialog.askopenfilename(
        title="Выберите файл",
        filetypes=(("Файлы excel", "*.xlsx"), ("Файлы excel", "*.xlsm*"))
    )

    if not file_path:
        return None
    try:
        workbook = openpyxl.load_workbook(file_path)
        return workbook
    except Exception as e:
        messagebox.showerror("Ошибка", f"Не удалось открыть файл: {e}")
        return None
