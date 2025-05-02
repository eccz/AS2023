import xml.etree.ElementTree as ET
import csv
import os

from src.d_parser import parse
from utils.utils import indent


def csv_input():
    result = dict()
    filepath = 'zadanie_YS2025.csv'
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


def parameters_build(source: dict, c_name, pset_mapping_dict=None):
    if not c_name:
        c_name = 'AS_2024'

    root = parameters_root_build()

    for key, value in source.items():
        if pset_mapping_dict:
            root.append(new_parameter(f'{value}', f'{value}', pset_mapping_dict.get(value)))
        else:
            root.append(new_parameter(f'{value}', f'{value}', c_name))

    root.append(ET.Element('MEASUREMENTS'))

    return root


def parameters_xml_output(el, filepath):
    indent(el)
    mydata = ET.tostring(el, encoding="utf-8", method="xml")

    with open(filepath, 'w', encoding="utf-8") as f:
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
        output_path = '../../data/parameters/param_profile_console.xml'
        parameters_xml_output(root, output_path)

        print(r'Файл param_profile_console.xml создан по адресу "/parameters"')
        print('Окно нужно закрыть')

        while True:
            input()


def parameters_maker_no_interface_d(source, param_group_name):
    # без интерфейса работает на основе парсинга приложения Д, см d_parser.py
    res = dict()
    inp = [j.get('el_attr_list') for j in source.values() if j.get('el_attr_list')]

    for m in inp:

        for n in m:
            res.update({n[1]: n[0]})

    return parameters_build(res, param_group_name)


def parameters_maker_no_interface_full(source, param_group_name, pset_mapping):
    # без интерфейса работает на основе парсинга приложения Д, см d_parser.py
    if pset_mapping:
        pset_mapping_dict = source['pset_mapping']
    else:
        pset_mapping_dict = None
    params = source['full_attr_list']

    return parameters_build(params, param_group_name, pset_mapping_dict)


if __name__ == '__main__':
    import openpyxl

    workbook = openpyxl.load_workbook("../../data/ADD_D_AS_2025_cleaned.xlsx")
    src = parse(workbook, to_term=True, to_json=False)
    # res_xml = parameters_maker_console()
    # parameters_xml_output(parameters_maker_no_interface_d(src, 'AS_2025'), '../../data/parameters/parameters_d.xml')
    parameters_xml_output(parameters_maker_no_interface_full(src, 'AS_2025', pset_mapping=False), '../../data/parameters/parameters_full.xml')
