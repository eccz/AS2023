import tkinter as tk
from tkinter import messagebox, ttk
import pyodbc

def connect_to_db():
    server = server_entry.get()
    database = db_combobox.get()
    username = user_entry.get()
    password = password_entry.get()
    auth_mode = auth_var.get()

    try:
        if auth_mode == "Windows":
            return pyodbc.connect(f'DRIVER={{SQL Server}};SERVER={server};DATABASE={database};Trusted_Connection=yes')
        else:
            return pyodbc.connect(f'DRIVER={{SQL Server}};SERVER={server};DATABASE={database};UID={username};PWD={password}')

    except Exception as err:
            messagebox.showerror("Ошибка", f"Ошибка подключения к sql-server: {err}")


def fetch_databases():
    conn = connect_to_db()

    try:
        cursor = conn.cursor()
        cursor.execute("SELECT name FROM sys.databases")
        databases = [row[0] for row in cursor.fetchall()]
        db_combobox['values'] = databases
        print(db_combobox['values'])
        if databases:
            db_combobox.current(0)

        conn.close()
    except Exception as err:
        messagebox.showerror("Ошибка", f"Ошибка получения баз данных: {err}")


def clear_parameters():
    conn = connect_to_db()
    try:
        cursor = conn.cursor()
        cursor.execute("exec dbo.spParamDefsClear")
        conn.commit()

        messagebox.showinfo("Успех", "Параметры возвращены к default!")
        conn.close()
    except Exception as err:
        messagebox.showerror("Ошибка", f"Ошибка выполнения очистки параметров: {err}")


def fix_parameters_counter():
    conn = connect_to_db()
    try:
        cursor = conn.cursor()

        query = "SELECT TOP 1 * FROM dbo.ParamDefs ORDER BY idParamDef DESC"
        cursor.execute(query)

        max_param_id = (cursor.fetchone()[0])

        update_query = f"UPDATE dbo.ParamDefs_IDENTITY SET id = {max_param_id}"
        cursor.execute(update_query)
        conn.commit()

        messagebox.showinfo("Успех", "Счетчик параметров исправлен")
        cursor.close()
        conn.close()

    except Exception as err:
        messagebox.showerror("Ошибка", f"Ошибка выполнения исправления счетчика параметров: {err}")


# Создание основного окна
root = tk.Tk()
root.title("Подключение к MS SQL Server")
root.geometry("400x800")

# Поля для ввода данных
tk.Label(root, text="Сервер:").pack(pady=5)
server_entry = tk.Entry(root)
server_entry.pack(pady=5)
server_entry.insert(0, "(local)\\SQLEXPRESS")

tk.Label(root, text="Аутентификация:").pack(pady=5)
auth_var = tk.StringVar(value="SQL")
radio_sql = tk.Radiobutton(root, text="SQL Server", variable=auth_var, value="SQL")
radio_sql.pack()
radio_windows = tk.Radiobutton(root, text="Windows", variable=auth_var, value="Windows")
radio_windows.pack()

tk.Label(root, text="Пользователь:").pack(pady=5)
user_entry = tk.Entry(root)
user_entry.pack(pady=5)
user_entry.insert(0, "sa")

tk.Label(root, text="Пароль:").pack(pady=5)
password_entry = tk.Entry(root, show="*")
password_entry.pack(pady=5)
password_entry.insert(0, "123")

# Кнопка для загрузки списка баз данных
fetch_db_button = tk.Button(root, text="Получить базы", command=fetch_databases)
fetch_db_button.pack(pady=5)

tk.Label(root, text="База данных:").pack(pady=5)
db_combobox = ttk.Combobox(root, state="readonly")
db_combobox.pack(pady=5)

# Кнопка очистки параметров CadLib
connect_button = tk.Button(root, text="Очистить параметры", command=clear_parameters)
connect_button.pack(pady=20)

connect_button = tk.Button(root, text="Исправить счетчик параметров", command=fix_parameters_counter)
connect_button.pack()

root.mainloop()
