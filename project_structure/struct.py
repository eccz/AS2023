import csv
import xml.etree.ElementTree as ET
import os


def csv_input(filepath):
    result = []

    with open(filepath, 'r', encoding='cp1251', newline='') as file:
        reader = csv.reader(file, delimiter=';')
        [result.append(row[2].strip()) for row in reader if row[2].strip() not in result]
    return result


def struct_root_element_build(name):
    element_args = dict(name=f"{name}", id="0", uid="0")

    return ET.Element('Element', **element_args)


def struct_parameters_build():
    parameters = ET.Element('Parameters')
    parameter_1_args = dict(name="PROJECT_STRUCT_LEVEL", value="1.  Стадия", caption="Уровень иерархии", comment="")
    parameter_2_args = dict(name="SYS_CATEGORY_GROUP", value="Разделы проекта", caption="Группа данных",
                            comment="DisciplinesHierarchy")
    parameter_3_args = dict(name="USER_ACCESS_GROUPS", value="", caption="Доступ", comment="")

    parameters.append(ET.Element('Parameter', **parameter_1_args))
    parameters.append(ET.Element('Parameter', **parameter_2_args))
    parameters.append(ET.Element('Parameter', **parameter_3_args))

    return parameters


def struct_el_element_build(name, id_n):
    element = ET.Element('Element', name=f"{name}", id=f'{id_n}', uid="0")

    parameters = ET.Element('Parameters')
    params = []
    params.append(dict(name="INFOMODEL_CODE", value="*", caption="Шифр",
                       comment='=if ([INFOMODEL_TAG]<>"", [INFOMODEL_TAG],"")'))
    params.append(dict(name="INFOMODEL_NAME", value="", caption="Наименование", comment=""))
    params.append(dict(name="INFOMODEL_TAG", value="", caption="Обозначение", comment=""))
    params.append(dict(name="PROJECT_STRUCT_LEVEL", value="2. Разделы проекта", caption="Уровень иерархии", comment=""))
    params.append(dict(name="SYS_CATEGORY_GROUP", value="Разделы проекта", caption="Группа данных",
                       comment="DisciplinesHierarchy"))
    params.append(dict(name="USER_ACCESS_GROUPS", value="", caption="Доступ", comment=""))

    [parameters.append(ET.Element('Parameter', **param)) for param in params]
    element.append(parameters)

    return element


def struct_xml_build(file, name):
    src = csv_input(file)

    elements = ET.Element('Elements')
    for i, e in enumerate(src):
        elements.append(struct_el_element_build(e, i + 1))

    parameters = struct_parameters_build()
    root = struct_root_element_build(name)

    root.append(parameters)
    root.append(elements)

    return root


def struct_file_creation(mydata, name):
    with open(fr'work/results/{name}.xml', 'w', encoding="utf-8") as f:
        f.write('<?xml version="1.0" ?>\n')
        f.write(mydata.decode(encoding="utf-8"))


def main():
    while True:
        print(f'{"#" * 30}\nCкрипт для создания xml-профиля структуры проекта\n{"#" * 30}\n')

        print(f'Убедись, что в папке присутствуют необходимые csv-файлы по специальностям\n'
              f'Список файлов в папке "AS2023/src":')
        [print(f'{i + 1}. {e}') for i, e in enumerate(os.listdir('../src')) if e.endswith('.csv')]

        res_print = []
        for fname in os.listdir('../src'):
            if fname.endswith('.csv'):
                print(f'Введи наименование категории для файла: {fname}, например Строительная_часть')
                category = input()
            else:
                continue

            mydata = ET.tostring(struct_xml_build(
                file=f'../src/{fname}',
                name=f'{category}'),
                encoding="utf-8",
                method="xml")

            struct_file_creation(mydata, f'{fname.split(".")[0].strip()}_structure')
            res_print.append(f'{fname.split(".")[0].strip()}_structure.xml')

        [print(f'Файл {i} создан в папке results') for i in res_print]
        print('\nОкно нужно закрыть')

        while True:
            input()


if __name__ == '__main__':
    main()
