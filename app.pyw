import tkinter as tk
from ui.notebook import create_notebook
import datetime


def main():
    root = tk.Tk()
    root.title("Atomskills Automation v1.0")
    root.geometry("700x350")
    root.resizable(False, False)

    # Создаем вкладки через функцию из модуля
    notebook = create_notebook(root)
    notebook.pack(fill="both", expand=True)

    root.mainloop()


if __name__ == "__main__":
    main() # if str(datetime.datetime.now()).startswith('2025-03-03') else None
