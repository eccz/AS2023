import xml.etree.ElementTree as ET
import csv
import os


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
    category = ET.Element('CATEGORY', order="2147483647", categoryOrder="1")
    category.text = cat_name
    categories.append(category)
    parameter.append(categories)

    return parameter


def csv_input(name):
    filepath = fr'../../data/{name}'
    result = dict()

    with open(filepath, 'r', encoding='cp1251', newline='') as file:
        reader = csv.reader(file, delimiter=';')
        for row in reader:
            result[row[1].strip()] = row[0].strip()

    return result


def main():
    while True:
        print(f'{"#" * 30}\nCкрипт для создания xml для Менеджера библиотек\n{"#" * 30}\n')

        print(f'Выбери csv-файл для работы\n'
              f'Список файлов в папке "AS2023/data":')
        [print(f'{i + 1}. {e}') for i, e in enumerate(os.listdir('../../data'))]

        print(f'\nВведи порядковый номер нужного файла:')
        src_csv_name = os.listdir('../../data')[int(input().strip()) - 1]
        print(f'Выбран файл "{src_csv_name}"\n')

        print(f'Выбери исходную xml-базу параметров\n'
              f'Список файлов в папке "/library_manager/":')
        [print(f'{i + 1}. {e}') for i, e in enumerate(os.listdir())]

        print(f'\nВведи порядковый номер нужного файла:')
        src_xml_name = os.listdir()[int(input().strip()) - 1]
        print(f'Выбран файл "{src_xml_name}"\n')

        print(f'Введи наименование категории в соответствии с заданием (например AS2023):')
        cat_name = input().strip()
        print(f'Выбрано наименование "{cat_name.strip()}"\n')

        print('Нажми enter, чтобы создать xml')
        input()

        tree = ET.parse(fr'{src_xml_name}')
        root = tree.getroot()

        parameter_dict = csv_input(src_csv_name)
        for par in parameter_dict:
            res = new_parameter(par, parameter_dict[par], cat_name)
            root.append(res)

        tree.write(fr'../../data/lib_manager/{src_csv_name.split(".")[0]}_libr.xml', encoding='utf-8')

        print(fr'Файл {src_csv_name.split(".")[0]}_libr.xml создан по адресу "../../data/lib_manager"')
        print('Окно нужно закрыть')

        while True:
            input()


if __name__ == '__main__':
    main()
