import tkinter as tk
from ui.notebook import create_notebook

def main():
    root = tk.Tk()
    root.title("Atomskills")
    root.geometry("600x200")

    # Создаем вкладки через функцию из модуля
    notebook = create_notebook(root)
    notebook.pack(fill="both", expand=True)

    root.mainloop()

if __name__ == "__main__":
    main()
