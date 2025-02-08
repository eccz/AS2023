from tkinter import filedialog, messagebox
import openpyxl
from datetime import datetime


def choose_file(filepath):
    filename = filedialog.askopenfilename(
        title="!Выберите файл",
        filetypes=(("Файлы excel", "*.xlsx"), ("Файлы excel", "*.xlsm*"))
    )
    filepath.set(filename)


def to_xml_tk(el, output_func, filename):
    current_date = datetime.now().strftime("%d-%m-%Y_%H-%M-%S")

    default_filename = f"{filename}_{current_date}.xml"
    file_path = filedialog.asksaveasfilename(defaultextension=".xml",
                                             filetypes=[("xml файлы", "*.xml")],
                                             initialfile=default_filename)
    if not file_path:
        return

    try:
        output_func(el, file_path)
        messagebox.showinfo("Успех", "Файл успешно сохранен!")
    except Exception as e:
        messagebox.showerror("Ошибка", f"Не удалось сохранить файл: {e}")


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
