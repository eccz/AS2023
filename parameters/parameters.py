import xml.etree.ElementTree as ET
import csv
import os
import openpyxl

from d_parser.d_parser import parse
from utils import indent


def csv_input():
    result = dict()
    filepath = 'zadanie.csv'
    with open(filepath, 'r', encoding='cp1251', newline='') as file:
        reader = csv.reader(file, delimiter=';')
        for row in reader:
            result[row[1]] = row[0]
    return result


def new_parameter(attr_name, attr_comment, cat_name):
    parameter = ET.Element('PARAMETER', type="1", prescision="-1", readonly="0", accuracy="-1", valueType="1",
                           sysNameMeasureUnitBase="", sysNameMeasureUnit="")

    name = ET.Element('NAME')
    name.text = attr_name
    parameter.append(name)

    comment = ET.Element('COMMENT')
    comment.text = attr_comment
    parameter.append(comment)

    default = ET.Element('DEFAULT')
    parameter.append(default)

    categories = ET.Element('CATEGORIES')
    category = ET.Element('CATEGORY', order="1", categoryOrder="2147483647")
    category.text = cat_name
    categories.append(category)
    parameter.append(categories)

    return parameter


def parameters_root_build():
    return ET.Element('PARAMETERS')


def parameters_build(source: dict, c_name):
    if not c_name:
        c_name = PARAMETER_GROUP_NAME

    root = parameters_root_build()

    for key, value in source.items():
        root.append(new_parameter(f'{key}', f'{key}', c_name))

    root.append(ET.Element('MEASUREMENTS'))

    return root


def parameters_xml_output(el, filename):
    indent(el)
    mydata = ET.tostring(el, encoding="utf-8", method="xml")

    with open(filename, 'w', encoding="utf-8") as f:
        f.write('<?xml version="1.0" encoding="utf-8"?>\n')
        f.write(mydata.decode(encoding="utf-8"))


def parameters_maker_console():
    # с интерфейсом берет данные из файла zadanie.csv
    while True:
        print(f'{"#" * 30}\nCкрипт для создания xml-профиля параметров для CADLib\n{"#" * 30}\n')

        print(f'Убедись, что в папке присутствует файл zadanie.csv с именами атрибутов\n'
              f'Список файлов в папке "/parameters":')
        [print(f'{i + 1}. {e}') for i, e in enumerate(os.listdir())]
        input()

        print(f'Введите наименование группы параметров\n'
              f'По умолчанию AS_2024')

        c_name = input()

        source = csv_input()
        root = parameters_build(source, c_name)
        output_filename = 'param_profile_1.xml'
        parameters_xml_output(root, output_filename)

        print(r'Файл param_profile.xml создан по адресу "/parameters"')
        print('Окно нужно закрыть')

        while True:
            input()


def parameters_maker_no_interface(source, param_group_name, output_filename='param_profile_1.xml'):
    # без интерфейса работает на основе парсинга приложения Д, см d_parser.py
    res = dict()
    inp = [j['el_attr_list'] for j in source.values()]


    for m in inp:
        for n in m:
            res.update({n[0]: n[1]})

    print(res)

    root = parameters_build(res, param_group_name)
    output_filename = output_filename
    parameters_xml_output(root, output_filename)


if __name__ == '__main__':
    workbook = openpyxl.load_workbook("../src/add_D.xlsx")
    src = parse(workbook, to_term=True, to_json=False)
    parameters_maker_no_interface(src, 'AS_2036')
