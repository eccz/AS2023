import xml.etree.ElementTree as ET
import csv
import os


def csv_input():
    result = set()
    for f_name in os.listdir('work/src'):
        if f_name.endswith('.csv'):
            filepath = fr'work/src/{f_name}'
            with open(filepath, 'r', encoding='cp1251', newline='') as file:
                reader = csv.reader(file, delimiter=';')
                [result.add(row[1].strip()) for row in reader if row[1]]
    return result


def ifc_param_filter_record_build():
    param_filter_record = ET.Element('ParamFilterRecord')

    cadlib_par_name = ET.Element('CADLibParName')
    cadlib_par_name.text = 'AllObjects'

    cadlib_par_caption = ET.Element('CADLibParCaption')
    cadlib_par_caption.text = 'AllObjects'

    cadlib_par_value = ET.Element('CADLibParValue')
    cadlib_par_value.text = 'AllObjects'

    param_filter_record.append(cadlib_par_name)
    param_filter_record.append(cadlib_par_caption)
    param_filter_record.append(cadlib_par_value)

    return param_filter_record


def ifc_property_record_build(param, ifc_property_set='AS2023'):
    ifc_property_record = ET.Element('IfcPropertyRecord')
    name = ET.Element('Name')
    name.text = f'{ifc_property_set}.{param}'

    caption = ET.Element('Caption')

    value = ET.Element('Value')
    value.text = f'[{param}]'

    comment = ET.Element('Comment')

    property_set = ET.Element('PropertySet')
    property_set.text = 'CADLib_pset'

    value_is_function = ET.Element('ValueIsFunction')
    value_is_function.text = 'true'

    value_is_param = ET.Element('ValueIsParam')
    value_is_param.text = 'false'

    ifc_property_record.append(name)
    ifc_property_record.append(caption)
    ifc_property_record.append(value)
    ifc_property_record.append(comment)
    ifc_property_record.append(property_set)
    ifc_property_record.append(value_is_function)
    ifc_property_record.append(value_is_param)

    return ifc_property_record


def ifc_root_build():
    xsd = "http://www.w3.org/2001/XMLSchema"
    xsi = "http://www.w3.org/2001/XMLSchema-instance"
    ns = {"xmlns:xsd": xsd, "xmlns:xsi": xsi}

    return ET.Element('IfcExportProfileXML', **ns)


def ifc_xml_build(property_set):
    root = ifc_root_build()
    filter_records = ET.Element('FilterRecords')
    param_filter_record = ifc_param_filter_record_build()
    filter_records_ = ET.Element('FilterRecords')
    ifc_params = ET.Element('IfcParams')

    publish_non_prof_params = ET.Element('PublishNonProfParams')
    publish_non_prof_params.text = 'false'

    publish_empty_params = ET.Element('PublishEmptyParams')
    publish_empty_params.text = 'false'

    for par in csv_input():
        ifc_params.append(ifc_property_record_build(par, property_set))

    param_filter_record.append(filter_records_)
    param_filter_record.append(ifc_params)
    filter_records.append(param_filter_record)
    root.append(filter_records)
    root.append(publish_non_prof_params)
    root.append(publish_empty_params)

    return root


def main():
    while True:
        print(f'{"#" * 30}\nCкрипт для создания xml-профиля для экспорта IFC\n{"#" * 30}\n')

        print(f'Убедись, что в папке присутствуют все 3 csv-файла\n'
              f'Список файлов в папке "/specificator/work/src":')
        [print(f'{i + 1}. {e}') for i, e in enumerate(os.listdir('work/src'))]

        print(f'\nВведи наименование property_set в соответствии с заданием (например AS2023):')
        property_set = input()
        print(f'Выбрано наименование "{property_set.strip()}"\n')

        print('Нажми enter, чтобы создать xml')
        input()

        mydata = ET.tostring(ifc_xml_build(property_set=property_set), encoding="utf-8", method="xml")

        with open(r'work/results/ifc_profile.xml', 'w', encoding="utf-8") as f:
            f.write('<?xml version="1.0" ?>\n')
            f.write(mydata.decode(encoding="utf-8"))

        print(r'Файл ifc_profile.xml создан по адресу "/specificator/work/results/"')
        print('Окно нужно закрыть')

        while True:
            input()


if __name__ == '__main__':
    main()
