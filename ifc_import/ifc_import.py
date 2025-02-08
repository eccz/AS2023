import xml.etree.ElementTree as ET
import csv
import os
from utils import indent
import openpyxl
from d_parser.d_parser import parse


def ifc_input_xml_output(el, filepath):
    indent(el)
    mydata = ET.tostring(el, encoding="utf-8", method="xml")

    with open(filepath, 'w', encoding="utf-8") as f:
        f.write('<?xml version="1.0" encoding="utf-8"?>\n')
        f.write(mydata.decode(encoding="utf-8"))


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
    # value.text = f'[{ifc_property_set}.{param}])'
    # value.text = f'if([{ifc_property_set}.{param}]="[{param}]", "", [{ifc_property_set}.{param}])'
    value.text = f'if([{ifc_property_set}.{param}]="[{param}]" or [{ifc_property_set}.{param}]="0", "", [{ifc_property_set}.{param}])'

    comment = ET.Element('Comment')

    param_category = ET.Element('ParamCategory')
    param_category.text = ifc_property_set

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


def ifc_import_profile_build(inp: dict, ifc_property_set: str):
    root = ifc_root_build()
    profile = ET.Element('Profile')
    item = ET.Element('item')
    key = ET.Element('key')

    string = ET.Element('string')
    string.text = 'AnyIfcClass'

    value = ET.Element('value')

    array_of_cadlib_par_records = ET.Element('ArrayOfCADLibParRecord')

    # print(inp)
    for k, v in inp.items():
        array_of_cadlib_par_records.append(ifc_par_record_maker(v, f'{ifc_property_set}'))

    key.append(string)
    item.append(key)
    value.append(array_of_cadlib_par_records)
    item.append(value)
    profile.append(item)
    root.append(profile)

    return root


def ifc_import_maker_console():
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

        inp = csv_input()
        res = ifc_import_profile_build(inp, ifc_property_set=property_set)

        if prettify and prettify.lower() != 'false':
            indent(res)

        mydata = ET.tostring(res, encoding="utf-8", method="xml")

        with open(r'ifc_import_profile.xml', 'w', encoding="utf-8") as f:
            f.write('<?xml version="1.0" ?>\n')
            f.write(mydata.decode(encoding="utf-8"))

        print(r'Файл ifc_import_profile.xml создан по адресу "/ifc_import"')
        print('Окно нужно закрыть')

        while True:
            input()


def ifc_import_maker_no_interface_d(source, property_set='EngeneeringDesign'):
    # без интерфейса работает на основе парсинга элементов приложения Д, см d_parser.py
    params = dict()
    inp = [j.get('el_attr_list') for j in source.values() if j.get('el_attr_list')]

    for m in inp:
        for n in m:
            params.update({n[1]: n[0]})

    return ifc_import_profile_build(params, ifc_property_set=property_set)


def ifc_import_maker_no_interface_full(source, property_set='EngeneeringDesign'):
    # без интерфейса работает на основе парсинга вкладки ATTRIBUTES приложения Д, см d_parser.py
    params = source['full_attr_list']
    return ifc_import_profile_build(params, ifc_property_set=property_set)


if __name__ == '__main__':
    # ifc_import_maker_console()

    workbook = openpyxl.load_workbook("../src/add_D.xlsx")
    src = parse(workbook, to_term=True, to_json=False)
    ifc_input_xml_output(ifc_import_maker_no_interface_d(src), 'ifc_import_2.xml')
    # ifc_input_xml_output(ifc_import_maker_no_interface_full(src), 'ifc_import_3.xml')
