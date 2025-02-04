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


def ifc_root_build():
    xsd = "http://www.w3.org/2001/XMLSchema"
    xsi = "http://www.w3.org/2001/XMLSchema-instance"
    ns = {"xmlns:xsd": xsd, "xmlns:xsi": xsi}
    return ET.Element('IfcImportProfileXML', **ns)


def ifc_par_record_maker(param, ifc_property_set):
    ifc_par_record = ET.Element('CADLibParRecord')
    name = ET.Element('Name')
    name.text = f'{param}'

    caption = ET.Element('Caption')
    caption.text = f'{param}'

    value = ET.Element('Value')
    value.text = f'[{ifc_property_set}.{param}])'
    # value.text = f'if([{ifc_property_set}.{param}]="[{param}]", "", [{ifc_property_set}.{param}])'

    comment = ET.Element('Comment')
    param_category = ET.Element('ParamCategory')

    value_is_function = ET.Element('ValueIsFunction')
    value_is_function.text = 'true'

    value_is_parameter = ET.Element('ValueIsParam')
    value_is_parameter.text = 'false'

    ifc_par_record.append(name)
    ifc_par_record.append(caption)
    ifc_par_record.append(value)
    ifc_par_record.append(comment)
    ifc_par_record.append(param_category)
    ifc_par_record.append(value_is_function)
    ifc_par_record.append(value_is_parameter)

    return ifc_par_record


def ifc_import_profile_build(ifc_property_set):
    root = ifc_root_build()
    profile = ET.Element('Profile')
    item = ET.Element('item')
    key = ET.Element('key')

    string = ET.Element('string')
    string.text = 'AnyIfcClass'

    value = ET.Element('value')

    array_of_cadlib_par_records = ET.Element('ArrayOfCADLibParRecord')

    for k, v in csv_input().items():
        array_of_cadlib_par_records.append(ifc_par_record_maker(v, f'{ifc_property_set}'))

    key.append(string)
    item.append(key)
    value.append(array_of_cadlib_par_records)
    item.append(value)
    profile.append(item)
    root.append(profile)

    return root


def main():
    while True:
        print(f'{"#" * 30}\nCкрипт для создания xml-профиля для импорта IFC\n{"#" * 30}\n')

        print(f'Убедись, что в папке присутствует файл zadanie.csv с именами атрибутов\n'
              f'Список файлов в папке "AS2023/ifc_import":')
        [print(f'{i + 1}. {e}') for i, e in enumerate(os.listdir())]
        input()

        print(f'Введите наименование PropertySet\n'
              f'По умолчанию EngeneeringDesign')

        property_set = input()
        if not property_set:
            property_set = 'EngeneeringDesign'

        print(f'Prettify?\n'
              f'True/False')

        prettify = input()

        res = ifc_import_profile_build(ifc_property_set=property_set)

        if prettify:
            indent(res)

        mydata = ET.tostring(res, encoding="utf-8", method="xml")

        with open(r'ifc_import_profile.xml', 'w', encoding="utf-8") as f:
            f.write('<?xml version="1.0" ?>\n')
            f.write(mydata.decode(encoding="utf-8"))

        print(r'Файл ifc_import_profile.xml создан по адресу "/ifc_import"')
        print('Окно нужно закрыть')

        while True:
            input()


if __name__ == '__main__':
    main()
