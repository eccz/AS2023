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


def ifc_param_filter_record_build(par_name, par_caption, par_value, f_records=None, ifc_params=None, ms_mapping=None):
    """
    Строит XML-узел `ParamFilterRecord`, описывающий параметр фильтрации IFC.
    Аргументы:
    - par_name, par_caption, par_value — данные параметра
    - f_records — вложенные записи фильтрации (опционально)
    - ifc_params — параметры IFC (опционально)
    - ms_mapping — словарь для сопоставления параметров с принятыми в Model Studio и библиотеках элементов
    """
    param_filter_record = ET.Element('ParamFilterRecord')

    cadlib_par_name = ET.Element('CADLibParName')
    cadlib_par_name.text = ms_mapping.get(par_name) if ms_mapping else par_name

    cadlib_par_caption = ET.Element('CADLibParCaption')
    cadlib_par_caption.text = ms_mapping.get(par_caption) if ms_mapping else par_caption

    cadlib_par_value = ET.Element('CADLibParValue')
    cadlib_par_value.text = par_value

    el_filter_records = ET.Element('FilterRecords') if not f_records else f_records
    el_ifc_params = ET.Element('IfcParams') if not ifc_params else ifc_params

    param_filter_record.append(cadlib_par_name)
    param_filter_record.append(cadlib_par_caption)
    param_filter_record.append(cadlib_par_value)
    param_filter_record.append(el_filter_records)
    param_filter_record.append(el_ifc_params)

    return param_filter_record


def ifc_property_record_build(ifc_property_set, param, value, ifc=False):
    """
    Строит XML-элемент `IfcPropertyRecord`, описывающий свойство IFC.
    - ifc: если True — создаётся элемент для IFC-класса
    """
    ifc_property_record = ET.Element('IfcPropertyRecord')
    name = ET.Element('Name')
    if ifc:
        name.text = f'IFC_ENTITY_CLASS'
        caption = ET.Element('Caption')
        value_ = ET.Element('Value')
        value_.text = f'{value}'
    else:
        name.text = f'{ifc_property_set}.{param}'
        caption = ET.Element('Caption')
        value_ = ET.Element('Value')
        value_.text = f'[{value}]'

    comment = ET.Element('Comment')

    property_set = ET.Element('PropertySet')
    property_set.text = 'CADLib_pset'

    value_is_function = ET.Element('ValueIsFunction')
    value_is_function.text = 'true'

    value_is_param = ET.Element('ValueIsParam')
    value_is_param.text = 'false'

    ifc_property_record.append(name)
    ifc_property_record.append(caption)

    ifc_property_record.append(value_)
    ifc_property_record.append(comment)
    ifc_property_record.append(property_set)
    ifc_property_record.append(value_is_function)
    ifc_property_record.append(value_is_param)

    return ifc_property_record


def ifc_params_build(el, ifc_property_set, ms_mapping):
    """
    Строит XML-элемент `IfcParams` на основе списка атрибутов и IFC-класса.
    """
    ifc_params = ET.Element('IfcParams')
    if len(el['IFC']) == 1:
        ifc_class = el.get('IFC')[0]
        ifc_params.append(ifc_property_record_build(ifc_property_set, ifc_class, ifc_class, ifc=True))
    for prop in el['el_attr_list']:
        if ms_mapping:
            ifc_params.append(ifc_property_record_build(ifc_property_set, prop[0], ms_mapping[prop[0]]))
        else:
            ifc_params.append(ifc_property_record_build(ifc_property_set, prop[0], prop[0]))
    return ifc_params


def ifc_filter_records_build(el, ifc_property_set, ms_mapping=None):
    """
    Создаёт вложенные записи фильтрации (`FilterRecords`) для случая,
    когда объект имеет несколько IFC-классов.
    """
    if len(el['IFC']) > 1:
        f_records = ET.Element('FilterRecords')
        for n in el.get('IFC'):
            ifc_class = n[1]
            ifc_params = ET.Element('IfcParams')
            ifc_params.append(ifc_property_record_build(ifc_property_set, ifc_class, ifc_class, ifc=True))
            par_name = ms_mapping.get(ASSEMBLY_TYPE_ATTR_NAME) if ms_mapping else ASSEMBLY_TYPE_ATTR_NAME
            par_caption = par_name
            p_record = ifc_param_filter_record_build(par_name=par_name, par_caption=par_caption, par_value=n[2], ifc_params=ifc_params)
            f_records.append(p_record)
        return f_records
    else:
        return None


def ifc_root_build():
    """
    Создаёт корневой элемент XML-документа с пространствами имён.
    """
    xsd = "http://www.w3.org/2001/XMLSchema"
    xsi = "http://www.w3.org/2001/XMLSchema-instance"
    ns = {"xmlns:xsd": xsd, "xmlns:xsi": xsi}

    return ET.Element('IfcExportProfileXML', **ns)


def ifc_export_xml_build(property_set, src, ms_mapping=True, class2025=True):
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
    if class2025:
        class_ifc_params = ET.Element('IfcParams')
        classification = {'ED_Classification':['SYSTEM', 'SUBSYS', 'ELEMENT', 'SYSTEM_TAG', 'SUBSYS_TAG', 'ELEMENT_TAG'],
                          'ED_Encoding': ['SYSTEM_ID', 'SUBSYS_ID', 'ELEMENT_ID', 'SYSTEM_N', 'SUBSYS_N', 'ELEMENT_N', 'ELEMENT_F_N']}
        for p_set, par_list in classification.items():
            for par in par_list:
                class_ifc_params.append(ifc_property_record_build(p_set, par, par, ifc=False))
    else:
        class_ifc_params = None

    # Добавление записи фильтрации "AllObjects" для всех объектов
    empty_param_filter_record = ifc_param_filter_record_build('AllObjects', 'AllObjects', 'AllObjects', ifc_params=class_ifc_params)
    filter_records.append(empty_param_filter_record)

    publish_non_prof_params = ET.Element('PublishNonProfParams')
    publish_non_prof_params.text = 'false'

    publish_empty_params = ET.Element('PublishEmptyParams')
    publish_empty_params.text = 'false'

    # Определение соответствий имён параметров, если есть
    if ms_mapping and src.get('ms_attr_mapping'):
        ms_mapping_list = src['ms_attr_mapping']
    else:
        ms_mapping_list = None

    # Построение записей фильтрации для каждой таблицы
    for table, data in src.items():
        if table.startswith('Таблица'):
            el_type_attr = TYPE_ATTR_NAME if data.get(TYPE_ATTR_NAME) else TASK_TYPE_ATTR_NAME
            el_type = data.get(TYPE_ATTR_NAME) if data.get(TYPE_ATTR_NAME) else data.get(TASK_TYPE_ATTR_NAME)
            ifc_params = ifc_params_build(data, property_set, ms_mapping=ms_mapping_list)
            ifc_f_records = ifc_filter_records_build(data, property_set, ms_mapping=ms_mapping_list)

            param_filter_record = ifc_param_filter_record_build(el_type_attr, el_type_attr, el_type,
                                                                f_records=ifc_f_records,ifc_params=ifc_params,
                                                                ms_mapping=ms_mapping_list)
            filter_records.append(param_filter_record)

    # Финальное добавление элементов в корень
    root.append(filter_records)
    root.append(publish_non_prof_params)
    root.append(publish_empty_params)
    return root


if __name__ == '__main__':
    import openpyxl

    workbook = openpyxl.load_workbook("../../data/ADD_D_AS_2025_cleaned.xlsx")
    source = parse(workbook, to_term=False, to_json=False)

    pset = 'EngeneeringDesign'
    res = ifc_export_xml_build(pset, source, ms_mapping=True, class2025=True)

    ifc_export_xml_output(res, '../../data/export/ifc_profile_NEW.xml')
