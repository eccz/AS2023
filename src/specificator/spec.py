import xml.etree.ElementTree as ET
import csv
import os
from utils.utils import indent
from src.d_parser import parse
from config.config import TYPE_ATTR_NAME, SPECIALITY_ATTR_NAME

# Базовый список всех допустимых типов элементов
BASE_TYPES = ["aec_elements", "aec_plate", "aec_roof", "aec_surf", "aec_wall", "aec_wall_lay", "aecbore",
              "aecsectionreinf", "aecsite", "background", "bearings", "blockxreferences", "cable_journal",
              "cable_net_links", "cable_routing", "cables", "collisions", "crossarm", "crosses", "equipment_items",
              "fire_alarm_zones", "garlands", "grounding", "groundprofiles", "lines", "links", "linksketchs",
              "locations", "metal", "metalasm", "metalnode", "modifier", "ms_weld_seam", "mscartogram",
              "msgeoarrowslope", "msgeocircuitbuilding", "msgeocircuitbuildingaxis", "msgeocircuitbuildingopening",
              "msgeositeroad", "msparcel", "mssite", "mssite_contour", "mssite_isoline", "mssite_mainslope",
              "mssite_roads", "mssite_structlines", "mssite_surfline", "mssite_zeroworkline", "mssitebasehorizontal",
              "mssitebreakline", "mssitehorizontal", "mssiteslopehatch", "nodes", "overpass_axis", "picketage", "pipe",
              "pipe_insulation", "pipe_part", "pline_entity", "pline_unity", "profiles", "projectdummy", "projectframe",
              "reinfbar", "reinfplane", "reinfsection", "routeconstruction", "routeprototype", "segments", "syscalc",
              "tbl_mssitecarto_grid", "template", "templateelement", "trench", "unit", "units"]


def csv_input(f_name):
    """
    Читает CSV-файл и преобразует данные в словарь.

    :param f_name: Имя CSV-файла в папке ../../data/
    :return: Словарь вида {тип объекта: {параметр: значение}}
    """

    filepath = fr'../../data/{f_name}'
    result = dict()

    with open(filepath, 'r', encoding='cp1251', newline='') as file:
        reader = csv.reader(file, delimiter=';')
        for row in reader:
            key, subkey, subval = row[2].strip(), row[1].strip(), row[0].strip()
            if subval:
                result.setdefault(key, {})[subkey] = subval
    return result


def xml_element(tag, **attrs):
    """
    Упрощённое создание XML-элемента.

    :param tag: Название тега
    :param attrs: Атрибуты элемента
    :return: Элемент XML
    """
    return ET.Element(tag, **attrs)


def xml_table_build(elem_type, filter_param_1, filter_param_2, param_2_value):
    """
    Создаёт элемент Table с фильтрацией по типу и значению параметра.

    :param elem_type: Тип элемента
    :param filter_param_1: Название первого фильтра
    :param filter_param_2: Название второго фильтра
    :param param_2_value: Значение второго фильтра
    :return: Элемент Table
    """
    return xml_element('Table',
                       caption=elem_type,
                       filter=f'[{filter_param_1}] = "{elem_type}" AND [{filter_param_2}] = "{param_2_value}"',
                       result_filter="", aggregated="0")


def xml_types_build():
    """
    Генерирует элемент Types со всеми базовыми типами.

    :return: Элемент Types
    """
    types = xml_element('Types')
    types.extend(xml_element('Type', name=word.strip()) for word in BASE_TYPES)

    return types


def xml_fields_build(fields_data):
    """
    Создаёт элемент Fields с вложенными полями.

    :param fields_data: Словарь параметров {параметр: заголовок}
    :return: Элемент Fields
    """
    fields = xml_element('Fields')
    fields.extend(xml_element('Field', caption=name, data=param, type="0", aggregate="0", visible="1", format="")
                  for param, name in fields_data.items())
    return fields


def xml_view_build():
    """
    Генерирует элемент View с группировкой и сортировкой.

    :return: Элемент View
    """
    view = xml_element('View')
    view.extend([xml_element('GroupFields'), xml_element('SortFields')])
    return view


def xml_dataset_build(table, view):
    """
    Генерирует Dataset и добавляет туда таблицу и представление.

    :param table: Элемент Table
    :param view: Элемент View
    :return: Элемент Dataset
    """
    dataset = xml_element('Dataset',
                           assemblyGrouping="0", assemblyFilter="2", binding="Fields", relationType="",
                           join="outer", hierarchy="0")
    dataset.extend([table, view])
    return dataset


def xml_report_format_build():
    """
    Генерирует элемент ReportFormat с настройками вывода отчёта.

    :return: Элемент ReportFormat
    """
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
    """
    Генерирует элемент Extended с дополнительными параметрами парсинга.
    :return: Элемент Extended
    """
    params = [
        ("PX_PARSE_ASSEMBLIES", "Учитывать объекты внутри сборок"),
        ("PX_PARSE_BLOCKS", "Учитывать объекты внутри блоков"),
        ("PX_PARSE_XREFS", "Учитывать объекты внутри внешних ссылок"),
        ("PX_RESTORE_TYPES", "Использовать исходный тип для объектов проекта"),
        ("PX_WHOLE_PROJECT", "Учитывать объекты всех файлов текущего каталога")
    ]
    extended = xml_element('Extended')
    extended.extend(xml_element('Parameter', name=name, value="0", caption=caption, comment="") for name, caption in params)
    return extended


def specificator_xml_build(src, filter_param_1, filter_param_2, param_2_value, from_csv=False):
    """
    Собирает полный XML-отчёт по исходным данным.

    :param src: Данные в формате dict или имя CSV-файла
    :param filter_param_1: Название первого фильтра
    :param filter_param_2: Название второго фильтра
    :param param_2_value: Значение второго фильтра
    :param from_csv: Флаг загрузки из CSV
    :return: Элемент Report
    """
    data = csv_input(src) if from_csv else src
    root = xml_element('Report')
    dataset_profile = xml_element('DatasetProfile')


    for elem_type, param_dict in data.items():
        table = xml_table_build(elem_type, filter_param_1, filter_param_2, param_2_value)
        table.extend([xml_types_build(), xml_fields_build(param_dict)])
        dataset = xml_dataset_build(table, xml_view_build())
        dataset_profile.append(dataset)

    root.extend([dataset_profile, xml_report_format_build(), xml_extended_build()])
    return root


def specificator_no_interface(src, first_filter_param, second_filter_param, ms_mapping):
    """
    Построение XML-отчёта без пользовательского интерфейса, с поддержкой сопоставления атрибутов.
    Используется с интерфейсом в Tkinter

    :param src: Исходные данные
    :param first_filter_param: Название первого фильтра
    :param second_filter_param: Название второго фильтра
    :param ms_mapping: Флаг сопоставления атрибутов ms_attr_mapping
    :return: Словарь {специальность: XML-элемент Report}
    """
    src_data = dict()

    for table, data in src.items():
        if table.startswith('Таблица') and data.get(TYPE_ATTR_NAME):
            speciality = data[SPECIALITY_ATTR_NAME]
            elem = data[TYPE_ATTR_NAME]

            if ms_mapping:
                elem_attrs = {src['ms_attr_mapping'].get(a): b for a, b in data.get('el_attr_list')}
            else:
                elem_attrs = {a: b for a, b in data.get('el_attr_list')}

            if src_data.get(speciality):
                src_data[speciality].update({elem: elem_attrs})
            else:
                src_data[speciality] = {elem: elem_attrs}

    return {spec: specificator_xml_build(data, first_filter_param, second_filter_param, spec, from_csv=False)
            for spec, data in src_data.items()}


def specificator_xml_output(el, filepath):
    """
    Сохраняет XML-элемент в файл с форматированием.

    :param el: XML-элемент
    :param filepath: Путь для сохранения
    """
    indent(el)
    mydata = ET.tostring(el, encoding="utf-8", method="xml")

    with open(filepath, 'w', encoding="utf-8") as f:
        f.write('<?xml version="1.0" encoding="utf-8"?>\n')
        f.write(mydata.decode(encoding="utf-8"))


def main():
    while True:
        print(f'{"#" * 30}\nCкрипт для создания xml-профиля для спецификатора\n{"#" * 30}\n')

        print(f'Выбери csv-файл для работы\n'
              f'Список файлов в папке "/AS2023/data":')
        [print(f'{i + 1}. {e}') for i, e in enumerate(os.listdir('../../data'))]

        print(f'\nВведи порядковый номер нужного файла:')
        src = os.listdir('../../data')[int(input().strip()) - 1]
        print(f'Выбран файл "{src}"')

        print(f'\nВведи наименование параметра фильтрации по типу (например EHP_TYPE):')
        filter_param_1 = input().strip()
        print(f'Выбран параметр "{filter_param_1}"')

        print(f'\nВведи наименование дополнительного параметра фильтрации (например EHP_SPECIALITY):\n'
              f'В случае, если применение этого параметра не требуется, наименование и значение нужно оставить пустыми')
        filter_param_2 = input().strip()
        print(f'Выбран параметр "{filter_param_2}"')

        print(f'\nВведи значение для дополнительного параметра {filter_param_2}, например Строительная_часть:')
        param_2_value = input().strip()
        print(f'Выбрано значение "{param_2_value}"')

        print('Нажми enter, чтобы создать xml')
        input()

        mydata = specificator_xml_build(src, filter_param_1, filter_param_2, param_2_value, from_csv=True)
        indent(mydata)
        mydata = ET.tostring(mydata, encoding="utf-8", method="xml")

        with open(f'../../data/specificator/{src.split(".")[0].strip()}_spec.xml', 'w', encoding="utf-8") as f:
            f.write('<?xml version="1.0" ?>\n')
            f.write(mydata.decode(encoding="utf-8"))

        print(fr'Файл {src.split(".")[0]}_spec.xml создан по адресу "../../data/specificator"')
        print('Окно нужно закрыть')

        while True:
            input()


if __name__ == '__main__':
    import openpyxl
    # main()

    workbook = openpyxl.load_workbook("../../data/ADD_D_AS_2025_cleaned.xlsx")
    source = parse(workbook, to_term=False, to_json=False)

    type_filter = 'TYPE'
    spec_filter = 'SPECIALITY'
    spec_filter_value = 'Строительная_часть'
    res = specificator_no_interface(source, type_filter, spec_filter, ms_mapping=False)

    for m, n in res.items():
        f_path = f'../../data/specificator/{m}_SPEC.xml'
        specificator_xml_output(n, f_path)
