import xml.etree.ElementTree as ET

from utils.utils import indent
from src.d_parser import parse
from config.config import TASK_TYPE_ATTR_NAME, TYPE_ATTR_NAME, ASSEMBLY_TYPE_ATTR_NAME


def ifc_export_xml_output(el, filepath):
    """
    Сохраняет XML-дерево `el` в файл `filepath`.
    Добавляет декларацию XML и форматирует содержимое с помощью indent().
    """
    indent(el)
    data = ET.tostring(el, encoding="utf-8", method="xml")
    with open(filepath, 'w', encoding="utf-8") as f:
        f.write('<?xml version="1.0" encoding="utf-8"?>\n')
        f.write(data.decode(encoding="utf-8"))


def create_element(tag: str, text: str = None) -> ET.Element:
    el = ET.Element(tag)
    if text is not None:
        el.text = text
    return el


def ifc_property_record_build(p_set: str, param: str, value: str, ifc: bool = False) -> ET.Element:
    """
    Строит XML-элемент `IfcPropertyRecord`, описывающий свойство IFC.
    - ifc: если True — создаётся элемент для IFC-класса
    """
    record = ET.Element('IfcPropertyRecord')
    ifc_name = 'IFC_ENTITY_CLASS' if ifc else f'{p_set}.{param}'
    value_text = value if ifc else f'[{value}]'

    elements = [
        create_element('Name', ifc_name),
        create_element('Caption'),
        create_element('Value', value_text),
        create_element('Comment'),
        create_element('PropertySet', 'CADLib_pset'),
        create_element('ValueIsFunction', 'true'),
        create_element('ValueIsParam', 'false')
    ]
    record.extend(elements)
    return record


def ifc_param_filter_record_build(par_name: str, par_caption: str, par_value: str,
                                  f_records: ET.Element = None, ifc_params: ET.Element = None,
                                  ms_mapping_dict: dict = None) -> ET.Element:
    """
    Строит XML-узел `ParamFilterRecord`, описывающий параметр фильтрации IFC.
    Аргументы:
    - par_name, par_caption, par_value — данные параметра
    - f_records — вложенные записи фильтрации (опционально)
    - ifc_params — параметры IFC (опционально)
    - ms_mapping — словарь для сопоставления параметров с принятыми в Model Studio и библиотеках элементов
    """
    record = ET.Element('ParamFilterRecord')

    name = ms_mapping_dict.get(par_name, par_name) if ms_mapping_dict else par_name
    caption = ms_mapping_dict.get(par_caption, par_caption) if ms_mapping_dict else par_caption

    record.extend([
        create_element('CADLibParName', name),
        create_element('CADLibParCaption', caption),
        create_element('CADLibParValue', par_value),
        f_records if f_records is not None else ET.Element('FilterRecords'),
        ifc_params if ifc_params is not None else ET.Element('IfcParams')
    ])
    return record


def ifc_params_build(el: dict, p_set: str, ms_mapping_dict: dict, pset_mapping: bool) -> ET.Element:
    """
    Строит XML-элемент `IfcParams` на основе списка атрибутов и IFC-класса.
    """
    ifc_params = ET.Element('IfcParams')
    if len(el['IFC']) == 1:
        ifc_class = el.get('IFC')[0]
        ifc_params.append(ifc_property_record_build(p_set, ifc_class, ifc_class, ifc=True))
    if not pset_mapping:
        for prop in el['el_attr_list']:
            if ms_mapping_dict:
                ifc_params.append(ifc_property_record_build(p_set, prop[0], ms_mapping_dict[prop[0]]))
            else:
                ifc_params.append(ifc_property_record_build(p_set, prop[0], prop[0]))
    return ifc_params


def ifc_filter_records_build(el: dict, p_set: str, ms_mapping_dict: dict = None) -> ET.Element | None:
    """
    Создаёт вложенные записи фильтрации (`FilterRecords`) для случая,
    когда объект имеет несколько IFC-классов.
    """
    if len(el['IFC']) > 1:
        f_records = ET.Element('FilterRecords')
        for _, ifc_class, par_value in el.get('IFC'):
            ifc_params = ET.Element('IfcParams')
            ifc_params.append(ifc_property_record_build(p_set, ifc_class, ifc_class, ifc=True))
            par_name = ms_mapping_dict.get(ASSEMBLY_TYPE_ATTR_NAME) if ms_mapping_dict else ASSEMBLY_TYPE_ATTR_NAME
            par_caption = par_name
            p_record = ifc_param_filter_record_build(par_name=par_name, par_caption=par_caption, par_value=par_value,
                                                     ifc_params=ifc_params)
            f_records.append(p_record)

        nonfiltered = ifc_param_filter_record_build('NonFiltered', 'NonFiltered', 'NonFiltered')
        f_records.append(nonfiltered)
        return f_records
    else:
        return None


def ifc_root_build() -> ET.Element:
    """
    Создаёт корневой элемент XML-документа с пространствами имён.
    """
    xsd = "http://www.w3.org/2001/XMLSchema"
    xsi = "http://www.w3.org/2001/XMLSchema-instance"
    ns = {"xmlns:xsd": xsd, "xmlns:xsi": xsi}

    return ET.Element('IfcExportProfileXML', **ns)


def build_general_params(ms_mapping_dict: dict, pset_mapping_dict: dict, full_attr_list: dict) -> ET.Element:
    """Создаёт общий блок IfcParams."""
    params = ET.Element('IfcParams')
    for prop in full_attr_list.values():
        prop_set = pset_mapping_dict.get(prop, 'CADLib_pset') if pset_mapping_dict else 'CADLib_pset'
        value = ms_mapping_dict.get(prop, prop) if ms_mapping_dict else prop
        params.append(ifc_property_record_build(prop_set, prop, value))
    return params


def ifc_export_xml_build(property_set: str, src: dict,
                         ms_mapping: bool = True, pset_mapping: bool = True,
                         class2025: bool = False) -> ET.Element:
    """
    Строит полный XML-документ экспорта IFC.
    Аргументы:
    - property_set: имя набора свойств IFC
    - data: источник данных (спарсенный Excel)
    - ms_mapping: включить переименование параметров (по умолчанию True)
    - class2025: использовать классификатор AS2025 (по умолчанию True)
    """
    root = ifc_root_build()
    filter_records = ET.Element('FilterRecords')

    # Добавление параметров классификации, если включен классификатор AS2025
    general_params = ET.Element('IfcParams')
    if class2025:
        classification = {
            'ED_Classification': ['SYSTEM', 'SUBSYS', 'ELEMENT', 'SYSTEM_TAG', 'SUBSYS_TAG', 'ELEMENT_TAG'],
            'ED_Encoding': ['SYSTEM_ID', 'SUBSYS_ID', 'ELEMENT_ID', 'SYSTEM_N', 'SUBSYS_N', 'ELEMENT_N',
                            'ELEMENT_F_ID']}
        for p_set, par_list in classification.items():
            for par in par_list:
                general_params.append(ifc_property_record_build(p_set, par, par, ifc=False))

    # Определение соответствий имён параметров, если есть
    ms_mapping_dict = src.get('ms_attr_mapping') if ms_mapping else None

    pset_mapping_dict = src.get('pset_mapping') if pset_mapping else None
    if pset_mapping_dict and src.get('full_attr_list'):
        general_params.extend(build_general_params(ms_mapping_dict, pset_mapping_dict, src['full_attr_list']))

    # Добавление записи фильтрации "AllObjects" для всех объектов
    empty_param_filter_record = ifc_param_filter_record_build('AllObjects', 'AllObjects', 'AllObjects',
                                                              ifc_params=general_params)
    filter_records.append(empty_param_filter_record)

    publish_non_prof_params = ET.Element('PublishNonProfParams')
    publish_non_prof_params.text = 'false'

    publish_empty_params = ET.Element('PublishEmptyParams')
    publish_empty_params.text = 'false'

    # Построение записей фильтрации для каждой таблицы
    for table, data in src.items():
        if table.startswith('Таблица'):
            el_type_attr = TYPE_ATTR_NAME if data.get(TYPE_ATTR_NAME) else TASK_TYPE_ATTR_NAME
            el_type = data.get(el_type_attr)
            ifc_params = ifc_params_build(data, property_set, ms_mapping_dict=ms_mapping_dict,
                                          pset_mapping=pset_mapping)
            ifc_f_records = ifc_filter_records_build(data, property_set, ms_mapping_dict=ms_mapping_dict)

            param_filter_record = ifc_param_filter_record_build(el_type_attr, el_type_attr, el_type,
                                                                f_records=ifc_f_records, ifc_params=ifc_params,
                                                                ms_mapping_dict=ms_mapping_dict)
            filter_records.append(param_filter_record)

    # Финальное добавление элементов в корень
    nonfiltered = ifc_param_filter_record_build('NonFiltered', 'NonFiltered', 'NonFiltered')
    filter_records.append(nonfiltered)
    root.append(filter_records)
    root.append(publish_non_prof_params)
    root.append(publish_empty_params)
    return root


if __name__ == '__main__':
    import openpyxl

    workbook = openpyxl.load_workbook("../../data/ADD_D_AS_2025_cleaned.xlsx")
    source = parse(workbook, to_term=False, to_json=False)

    pset = 'EngeneeringDesign'
    res = ifc_export_xml_build(pset, source, ms_mapping=True, pset_mapping=True, class2025=True)

    ifc_export_xml_output(res, '../../data/export/ifc_profile_TTT.xml')
