# from library_manager.libr import indent
import xml.etree.ElementTree as ET
import csv
import os


def csv_input(f_name):
    filepath = fr'work/src/{f_name}'
    result = dict()

    with open(filepath, 'r', encoding='cp1251', newline='') as file:
        reader = csv.reader(file, delimiter=';')
        for row in reader:
            if row[2].strip() in result:
                result[row[2].strip()].update({row[1].strip(): row[0].strip()})
            elif not row[0].strip():
                continue
            else:
                result[row[2].strip()] = {row[1].strip(): row[0].strip()}
    return result


def xml_table_build(elem_type, filter_param):
    table_args = {"caption": elem_type, "filter": f'[{filter_param}] = "{elem_type}"',
                  "result.filter": "", "aggregated": "0"}

    return ET.Element('Table', **table_args)


def xml_types_build():
    types = ET.Element('Types')
    for word in open('work/base_types', 'r', encoding='utf-8'):
        type_ = ET.Element('Type', name=word.strip())
        types.append(type_)

    return types


def xml_fields_build():
    return ET.Element('Fields')


def xml_field_build(elem_type, elem_param):
    field_args = dict(caption=f"{elem_type}", data=f"{elem_param}", type="0", aggregate="0", visible="1", format="")

    return ET.Element('Field', **field_args)


def xml_view_build():
    view = ET.Element('View')
    view.append(ET.Element('GroupFields'))
    view.append(ET.Element('SortFields'))

    return view


def xml_dataset_build():
    dataset_args = dict(assemblyGrouping="0", assemblyFilter="2", binding="Fields", relationType="", join="outer",
                        hierarchy="0")

    return ET.Element('Dataset', **dataset_args)


def xml_report_format_build():
    report_format_args = {"application": "6", "title": "0", "parser": "", "outlining": "0", "headers.style": "2",
                          "headers.bold": "1", "table.separate": "0", "table.offset": "50", "table.offset.dir": "0",
                          "groups.style": "2", "groups.bold": "1", "groups.column": "1", "groups.slant": "0",
                          "groups.underline": "0", "macros": "", "template": "", "usefullpathtemplate": "0",
                          "encoding": "0", "worksheet": "", "wrap": "0", "xml.application": "", "xml.arguments": "",
                          "xml.script": "", "identifiers.out": "0", "xml.wait.results": "1", "totals.bold": "0",
                          "totals.italic": "0", "totals.underline": "0", "totals.comment": "",
                          "totals.comment.column": "1", "csv.divider": ";"}

    return ET.Element('ReportFormat', **report_format_args)


def xml_extended_build():
    extended = ET.Element('Extended')
    extended.append(
        ET.Element('Parameter', name="PX_PARSE_ASSEMBLIES", value="0", caption="Учитывать объекты внутри сборок",
                   comment=""))
    extended.append(
        ET.Element('Parameter', name="PX_PARSE_BLOCKS", value="0", caption="Учитывать объекты внутри блоков",
                   comment=""))
    extended.append(
        ET.Element('Parameter', name="PX_PARSE_XREFS", value="0", caption="Учитывать объекты внутри внешних ссылок",
                   comment=""))
    extended.append(ET.Element('Parameter', name="PX_RESTORE_TYPES", value="0",
                               caption="Использовать исходный тип для объектов проекта", comment=""))
    extended.append(ET.Element('Parameter', name="PX_WHOLE_PROJECT", value="0",
                               caption="Учитывать объекты всех файлов текущего каталога", comment=""))

    return extended


def xml_build(src, filer_param='EHP_TYPE'):
    source = csv_input(src)

    root = ET.Element('Report')
    dataset_profile = ET.Element('DatasetProfile')

    for elem_type, param_dict in source.items():
        dataset = xml_dataset_build()
        table = xml_table_build(elem_type=elem_type, filter_param=filer_param)
        types = xml_types_build()
        view = xml_view_build()
        fields = xml_fields_build()

        for param, name in param_dict.items():
            field = xml_field_build(elem_type=name, elem_param=param)
            fields.append(field)

        table.append(types)
        table.append(fields)
        dataset.append(table)
        dataset.append(view)
        dataset_profile.append(dataset)

    root.append(dataset_profile)
    root.append(xml_report_format_build())
    root.append(xml_extended_build())

    # indent(root)
    return root


def main():
    while True:
        print(f'{"#" * 30}\nCкрипт для создания xml-профиля для спецификатора\n{"#" * 30}\n')

        print(f'Выбери csv-файл для работы\n'
              f'Список файлов в папке "/specificator/work/src":')
        [print(f'{i + 1}. {e}') for i, e in enumerate(os.listdir('work/src'))]

        print(f'\nВведи порядковый номер нужного файла:')
        src = os.listdir('work/src')[int(input().strip()) - 1]
        print(f'Выбран файл "{src}"')

        print(f'\nВведи наименование параметра фильтрации (например EHP_TYPE):')
        filter_param = input().strip()
        print(f'Выбран параметр "{filter_param}"\n')

        print('Нажми enter, чтобы создать xml')
        input()

        mydata = ET.tostring(xml_build(src=src, filer_param=filter_param), encoding="utf-8", method="xml")

        with open(f'work/results/{src.split(".")[0].strip()}.xml', 'w', encoding="utf-8") as f:
            f.write('<?xml version="1.0" ?>\n')
            f.write(mydata.decode(encoding="utf-8"))

        print(fr'Файл {src.split(".")[0]}.xml создан по адресу "/specificator/work/results/"')
        print('Окно нужно закрыть')

        while True:
            input()


if __name__ == '__main__':
    main()
