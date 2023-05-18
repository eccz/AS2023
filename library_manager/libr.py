import xml.etree.ElementTree as ET
import csv
import os


def indent(elem, level=0):
    i = "\n" + level * "  "
    if len(elem):
        if not elem.text or not elem.text.strip():
            elem.text = i + "  "
        if not elem.tail or not elem.tail.strip():
            elem.tail = i
        for elem in elem:
            indent(elem, level + 1)
        if not elem.tail or not elem.tail.strip():
            elem.tail = i
    else:
        if level and (not elem.tail or not elem.tail.strip()):
            elem.tail = i


def new_parameter(attr_name, attr_comment):
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
    category = ET.Element('CATEGORY', order="2147483647", categoryOrder="1")
    category.text = 'AS2023'
    categories.append(category)
    parameter.append(categories)

    return parameter


def csv_input(name):
    filepath = fr'work/src/{name}'
    result = dict()

    with open(filepath, 'r', encoding='cp1251', newline='') as file:
        reader = csv.reader(file, delimiter=';')
        for row in reader:
            result[row[1]] = row[0]

    return result


if __name__ == '__main__':
    while True:
        print(f'{"#"*30}\nCкрипт для создания xml для Менеджера библиотек\n{"#"*30}\n')

        print(f'Выбери csv-файл для работы\n'
              f'Список файлов в папке "/library_manager/work/src":')
        [print(f'{i+1}. {e}') for i, e in enumerate(os.listdir('work/src'))]

        print(f'\nВведи порядковый номер нужного файла:')
        src_csv_name = os.listdir('work/src')[int(input())-1]
        print(f'Выбран файл "{src_csv_name}"\n')

        print(f'Выбери исходную xml-базу параметров\n'
              f'Список файлов в папке "/library_manager/work":')
        [print(f'{i+1}. {e}') for i, e in enumerate(os.listdir('work'))]

        print(f'\nВведи порядковый номер нужного файла:')
        src_xml_name = os.listdir('work')[int(input())-1]
        print(f'Выбран файл "{src_xml_name}"\n')

        print('Нажми enter, чтобы создать xml')
        input()

        tree = ET.parse(fr'work/{src_xml_name}')
        root = tree.getroot()

        parameter_dict = csv_input(src_csv_name)
        for par in parameter_dict:
            res = new_parameter(par, parameter_dict[par])
            root.append(res)

        tree.write(fr'work/results/{src_csv_name.split(".")[0]}.xml', encoding='utf-8')

        print(fr'Файл {src_csv_name.split(".")[0]}.xml создан по адресу "/library_manager/work/results/"')
        print('Окно нужно закрыть')

        while True:
            input()
