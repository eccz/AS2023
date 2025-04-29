import tkinter as tk

def create_info_tab(notebook):
    tab = tk.Frame(notebook)

    text_widget = tk.Text(
        tab,
        wrap="word",
        font=("Arial", 10),
        height=20,
        width=90,
        bd=0,
        highlightthickness=0,
        relief="flat"
    )
    text_widget.pack(padx=10, pady=10)

    # Теги
    text_widget.tag_configure("bold", font=("Arial", 10, "bold"))
    text_widget.tag_configure("paragraph", spacing1=4)  # Полуторный интервал

    # Вставка жирного заголовка
    text_widget.insert("1.0", "Приложение для автоматизации проверок на Atomskills R94 Инженерное проектирование.\n", ("bold", "paragraph"))

    remaining_text = (
        '1. На вкладке "Профиль отчета" создается профиль отчета для проверки в CadLib на основе приложения Д.\n'
        '2. На вкладке "Импорт параметров" создается профиль импорта параметров для CadLib на основе приложения Д. Доступны опции с полным списком параметров с листа Attributes, или только с используемыми в приложении Д.\n'
        '3. На вкладке "Профиль импорта IFC" создается профиль импорта IFC в CadLib на основе приложения Д. Доступны опции с полным списком параметров с листа Attributes, или только с используемыми в приложении Д.\n'
        '4. На вкладке "Профиль экспорта IFC" создается профиль экспорта в IFC из CadLib на основе приложения Д. Доступны опции с полным списком параметров с листа Attributes, или только с используемыми в приложении Д.'
        'Профиль экспорта учитывает назначение класса IFC в зависимости от групп.\n'
        '5. На вкладке Спецификатор создаются профили создания спецификатора по TYPE, SPECIALITY (3 профиля по специализациям АС, ТХ, ЭОМ)\n'
        '6. Вкладка "О программе" содержит описание функционала.'
    )
    text_widget.insert("end", remaining_text, "paragraph")

    text_widget.config(state="disabled")

    return tab
