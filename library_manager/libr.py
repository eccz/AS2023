import xml.etree.ElementTree as ET
import csv


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


def csv_input():  # указать в дальнейшем переменную
    filepath = r'../work/src/stroika.csv'
    result = []

    with open(filepath, 'r', encoding='cp1251', newline='') as file:
        reader = csv.reader(file, delimiter=';')
        for row in reader:
            result.append(row)

    return result


if __name__ == '__main__':
    tree = ET.parse(r'../work/LIBR_DEFAULT.xml')
    root = tree.getroot()

    for par in csv_input():
        res = new_parameter(par[0], par[1])
        root.append(res)

    tree.write('res.xml', encoding='utf-8')
