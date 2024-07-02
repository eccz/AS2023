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


def csv_input():
    result = dict()
    filepath = 'zadanie.csv'
    with open(filepath, 'r', encoding='cp1251', newline='') as file:
        reader = csv.reader(file, delimiter=';')
        for row in reader:
            result[row[0]] = row[1]
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


def ifc_root_build():

    return ET.Element('PARAMETERS')


def main():
    while True:
        print(f'{"#" * 30}\nCкрипт для создания xml-профиля параметров для CADLib\n{"#" * 30}\n')

        print(f'Убедись, что в папке присутствует файл zadanie.csv с именами атрибутов\n'
              f'Список файлов в папке "/parameters":')
        [print(f'{i + 1}. {e}') for i, e in enumerate(os.listdir())]
        input()

        print(f'Введите наименование группы параметров\n'
              f'По умолчанию AS_2024')

        c_name = input()
        if not c_name:
            c_name = 'AS_2024'

        base_xml = ifc_root_build()

        for key, value in csv_input().items():
            base_xml.append(new_parameter(f'{value}', f'{value}', c_name))

        base_xml.append(ET.Element('MEASUREMENTS'))

        indent(base_xml)
        mydata = ET.tostring(base_xml, encoding="utf-8", method="xml")

        with open(r'param_profile.xml', 'w', encoding="utf-8") as f:
            f.write('<?xml version="1.0" encoding="utf-8"?>\n')
            f.write(mydata.decode(encoding="utf-8"))

        print(r'Файл param_profile.xml создан по адресу "/parameters"')
        print('Окно нужно закрыть')

        while True:
            input()


if __name__ == '__main__':
    main()
