import config.config as cfg
from src.d_parser import parse
import xml.etree.ElementTree as ET
from src.ifc_import.ifc_import import indent

REPORT_TYPES = ["AEC_SURF", "COLLISIONS", "Cable_ETS", "HVAC", "HVACAxis", "HVACPipeItems", "IFCElement", "Insulation",
                "Materials", "OptionsList", "REINFORCEMENT", "TRAYS", "TowerSettings", "WireDefault", "WorkItem",
                "WorkList",
                "WorkResource", "aec_elements", "aec_layer_groups", "aec_layer_material", "assemblies", "assembly",
                "bolts",
                "cabinet", "cables", "cat160", "cat161", "cat47", "catDiameterList", "catLineSystems", "catMaterial",
                "catMaterials", "catPipeSupportStepData", "cat_clash", "climate", "commonLinks", "concreteware",
                "decoration",
                "device", "foundation", "garlandMisc", "garlands", "gesnCategoryList", "gesnResourceItem", "ground",
                "hierarchical_list", "iron_assemblies", "label", "metalware", "metalwarenode", "modifiers", "node",
                "nodes",
                "panels", "pipeAxis", "pipeCategoryList", "pipeItems", "pipeLineNumberList", "pipeMaterialList",
                "report",
                "routeConstructions", "routeConstructionsAssemblies", "structure_data", "suppressor", "symbol",
                "tank_eq_schm",
                "tank_equip", "towerequipment", "units", "valve_prototype", "wire", "wires", "zoneList"]

def report_profile_xml_output(el, filepath):
    """
    Сохраняет XML-элемент в файл с форматированием.

    :param el: Корневой XML-элемент.
    :param filepath: Путь к файлу для сохранения.
    """
    indent(el)
    mydata = ET.tostring(el, encoding="utf-8", method="xml")

    with open(filepath, 'w', encoding="utf-8") as f:
        f.write('<?xml version="1.0"?>\n')
        f.write(mydata.decode(encoding="utf-8"))


def cond_generator(loi):
    """
    Функция генерирует if-условие для CadLib на базе списка наименований атрибутов для определенного LOI
    :param loi: Список атрибутов LOI.
    :return: Строка условия.
    """
    res = []
    for i in loi:
        res.append(f'[{i}]<>""')
        res.append(f'[{i}]<>"0"')
    return ' and '.join(res)


def loi_if_generator(type_attr_name, type_attr, loi200=None, loi300=None, loi400=None):
    """
    Функция генерирует полное вложенное if-условие для CadLib по проверке LOI, с учетом наличия требований на LOI300 и LOI400.

    :param type_attr_name: Имя атрибута CadLib для фильтра, например TYPE.
    :param type_attr: Значение атрибута "TYPE" для фильтра в CadLib.
    :param loi200: Список атрибутов для LOI200.
    :param loi300: Список атрибутов для LOI300.
    :param loi400: Список атрибутов для LOI400.
    :return: Строка с выражением if.
    """
    if loi200 and loi300 and loi400:
        return f'if([{type_attr_name}]="{type_attr}", if({cond_generator(loi200)}, if({cond_generator(loi300)}, if({cond_generator(loi400)}, "LOI400", "LOI300"), "LOI200"), "LOI000"), "")'
    elif loi200 and loi300:
        return f'if([{type_attr_name}]="{type_attr}", if({cond_generator(loi200)}, if({cond_generator(loi300)}, "LOI300", "LOI200"),"LOI000"), "")'
    elif loi200:
        return f'if([{type_attr_name}]="{type_attr}", if({cond_generator(loi200)}, "LOI200","LOI000"), "")'
    return None


def ifc_if_generator(ifc_attr_name=cfg.IFC_ATTR_NAME, asm_attr_name=cfg.ASM_TYPE_ATTR_NAME, ifc_list=None):
    """
    Функция генерирует полное if-условие для CadLib по проверке правильности IFC-класса.
    Учитывает результаты парсинга приложения Д - учитывает требования по сборкам.
    type_attr - значение атрибута "TYPE" для фильтра в CadLib

    :param ifc_list: список списков для разных требований по сборкам либо список с одним элементом.
    :param ifc_attr_name: наименование атрибута CadLib, содержащего класс IFC
    :param asm_attr_name: наименование атрибута CadLib, содержащего тип сборки.
    :return: Строка с выражением if.
    """
    res = []
    if isinstance(ifc_list[0], list):
        for req in ifc_list:
            res.append(f'instr("{req[1]}", [{ifc_attr_name}])>=0 and [{asm_attr_name}]="{req[2]}"')
        res = ' or '.join(res)
        return f'if({res}, "CORRECT", "INCORRECT")'

    elif isinstance(ifc_list[0], str):
        return f'if(instr("{ifc_list[0]}", [IFC_TYPE])>=0, "CORRECT", "INCORRECT")'
    return None


def speciality_data_generator():
    """
    Генерирует проверку SPECIALITY через выражение if.
    :return: Строка условия.
    """
    specialities = cfg.SPECIALITIES
    speciality_attr_name = cfg.SPECIALITY_ATTR_NAME

    res = ' or '.join([f'[{speciality_attr_name}]="{i}"' for i in specialities])
    return f'if({res}, "CORRECT", "INCORRECT")'


def report_root_build():
    return ET.Element('Report', **{'db.version': "1"})


def report_dataset_profile_build():
    return ET.Element('DatasetProfile')


def report_format_build():
    ns = {'application': "2",
          'title': "0",
          'parser': "",
          'outlining': "0",
          'headers.style': "1",
          'headers.bold': "1",
          'table.separate': "0",
          'table.offset': "50",
          'table.offset.dir': "0",
          'groups.style': "2",
          'groups.bold': "1",
          'groups.column': "1",
          'groups.slant': "0",
          'groups.underline': "0",
          'macros': "",
          'template': "",
          'usefullpathtemplate': "0",
          'encoding': "0",
          'worksheet': "",
          'wrap': "0",
          'xml.application': "",
          'xml.arguments': "",
          'xml.script': "",
          'identifiers.out': "0",
          'xml.wait.results': "1",
          'totals.bold': "0",
          'totals.italic': "0",
          'totals.underline': "0",
          'totals.comment': "",
          'totals.comment.column': "1",
          'csv.divider': ";"}
    return ET.Element('ReportFormat', **ns)


def report_extended_build():
    extended = ET.Element('Extended')
    extended.append(ET.Element('Parameter',
                               name="PROP_EX_FILTER_STUCTURED",
                               value="1",
                               caption="Не применять к структурным объектам",
                               comment=""))
    extended.append(ET.Element('Parameter',
                               name="PROP_EX_FILTER_SUBSETS",
                               value="0",
                               caption="Не применять к подчиненным наборам данных",
                               comment=""))
    return extended


def report_dataset_build():
    ns = {'assemblyGrouping': "0", 'assemblyFilter': "2", 'binding': "Fields", 'relationType': "", 'join': "outer",
          'hierarchy': "0"}
    return ET.Element('Dataset', **ns)


def report_table_build(table_num, type_attr, _type):
    ns = {'caption': f"{table_num} [{_type}]", 'filter': f'[{type_attr}]="{_type}"', 'result.filter': "",
          'aggregated': "0"}
    return ET.Element('Table', **ns)


def report_types_build():
    types = ET.Element('Types')
    for i in REPORT_TYPES:
        types.append(ET.Element('Type', name=i))

    return types


def report_field_build(caption, data, type_='0'):
    return ET.Element('Field', caption=caption, data=data, type=type_, aggregate="0", visible="1", format='')


def report_fields_build(table_data, loi_data, ifc_flag_data, speciality_flag_data):
    fields = ET.Element('Fields')
    fields.append(report_field_build(caption="SYS_OBJECT_CATEGORY", data="SYS_OBJECT_CATEGORY"))
    fields.append(report_field_build(caption="SYS_OBJECT_NAME", data="@NAME"))
    fields.append(report_field_build(caption="IFC_TYPE", data="IFC_TYPE"))
    fields.append(report_field_build(caption="IfcGlobalId", data="IfcGlobalId"))
    fields.append(report_field_build(caption="TABLE", data=f'"{table_data}"', type_='1'))
    fields.append(report_field_build(caption="LOI", data=loi_data, type_='1'))
    fields.append(report_field_build(caption="IFC_TYPE_FLAG", data=ifc_flag_data, type_='1'))
    fields.append(report_field_build(caption="SPECIALITY_FLAG", data=speciality_flag_data, type_='1'))
    fields.append(report_field_build(caption=cfg.SPECIALITY_ATTR_NAME, data=cfg.SPECIALITY_ATTR_NAME))
    fields.append(report_field_build(caption=cfg.TYPE_ATTR_NAME, data=cfg.TYPE_ATTR_NAME))
    fields.append(report_field_build(caption=cfg.TASK_TYPE_ATTR_NAME, data=cfg.TASK_TYPE_ATTR_NAME))
    fields.append(report_field_build(caption=cfg.ASSEMBLY_TYPE_ATTR_NAME, data=cfg.ASSEMBLY_TYPE_ATTR_NAME))
    fields.append(report_field_build(caption=cfg.TAG_ATTR_NAME, data=cfg.TAG_ATTR_NAME))
    fields.append(report_field_build(caption=cfg.NAME_ATTR_NAME, data=cfg.NAME_ATTR_NAME))
    fields.append(report_field_build(caption=cfg.ASSEMBLY_MARK_ATTR_NAME, data=cfg.ASSEMBLY_MARK_ATTR_NAME))
    fields.append(report_field_build(caption=cfg.ASSEMBLY_N_ATTR_NAME, data=cfg.ASSEMBLY_N_ATTR_NAME))
    fields.append(report_field_build(caption=cfg.PROJECT_MARK_ATTR_NAME, data=cfg.PROJECT_MARK_ATTR_NAME))
    fields.append(report_field_build(caption=cfg.SYSTEM_TAG_ATTR_NAME, data=cfg.SYSTEM_TAG_ATTR_NAME))
    fields.append(report_field_build(caption=cfg.EL_LINE_TAG_ATTR_NAME, data=cfg.EL_LINE_TAG_ATTR_NAME))
    fields.append(report_field_build(caption=cfg.TASK_AUTHOR_ATTR_NAME, data=cfg.TASK_AUTHOR_ATTR_NAME))
    fields.append(report_field_build(caption=cfg.TASK_USER_ATTR_NAME, data=cfg.TASK_USER_ATTR_NAME))
    return fields


def report_view_build():
    view = ET.Element('View')
    view.append(ET.Element('GroupFields'))
    view.append(ET.Element('SortFields'))
    return view


def unknown_type_dataset_build(types_list):
    unknown_dataset = report_dataset_build()
    ftr = ' and '.join([f'[{i}]<>"{j}"' for i, j in types_list])
    ns = {'caption': "[НЕИЗВЕСТНЫЙ_ТИП]", 'filter': ftr, 'result.filter': "", 'aggregated': "0"}
    table = ET.Element('Table', **ns)
    table.append(report_types_build())
    table.append(
        report_fields_build(table_data='UNKNOWN_TYPE',
                            loi_data='"UNKNOWN_TYPE"',
                            ifc_flag_data='"UNKNOWN_TYPE"',
                            speciality_flag_data=speciality_data_generator()))
    unknown_dataset.append(table)

    return unknown_dataset


def xml_report_profile_build(source):
    """
    Формирует XML-структуру отчёта из словаря данных.

    :param source: Словарь с результатами парсинга приложения Д из Excel.
    :return: Корневой XML-элемент.
    """
    types_list = []
    types = report_types_build()

    root = report_root_build()
    report_dataset_profile = report_dataset_profile_build()

    for k, v in source.items():
        dataset = report_dataset_build()

        table_num = k
        type_attr_name = cfg.TYPE_ATTR_NAME if v.get(cfg.TYPE_ATTR_NAME) else cfg.TASK_TYPE_ATTR_NAME

        if v.get(type_attr_name) and v.get('element_name'):
            type_attr = v.get(type_attr_name)
        else:
            continue

        table = report_table_build(table_num, type_attr_name, type_attr)
        types_list.append([type_attr_name, type_attr])
        table.append(types)

        loi200 = v.get('LOI200')
        loi300 = v.get('LOI300')
        loi400 = v.get('LOI400')

        # print('NO LOI 400', '----', type_) if not loi400 else None

        fields = report_fields_build(table_data=f'{table_num} {type_attr}',
                                     loi_data=loi_if_generator(type_attr_name=type_attr_name,
                                                               type_attr=type_attr,
                                                               loi200=loi200,
                                                               loi300=loi300,
                                                               loi400=loi400),
                                     ifc_flag_data=ifc_if_generator(ifc_list=v.get('IFC')),
                                     speciality_flag_data=speciality_data_generator())

        table.append(fields)
        dataset.append(table)
        dataset.append(report_view_build())
        report_dataset_profile.append(dataset)

    report_dataset_profile.append(unknown_type_dataset_build(types_list))

    root.append(report_dataset_profile)

    root.append(report_format_build())
    root.append(report_extended_build())

    return root


if __name__ == '__main__':
    import openpyxl

    workbook = openpyxl.load_workbook("../../data/ADD_D_AS_2025_cleaned.xlsx")
    src = parse(workbook, to_term=True, to_json=False)
    output_xml_path = '../../data/report/'
    report_profile_xml_output(xml_report_profile_build(src), output_xml_path + 'report_profile.xml')

    # a = parse("../../data/add_D.xlsx", to_term=True)
    # name = 'Объемный элемент'
    # # loi200 = a[name].get('LOI200')
    # # loi300 = a[name].get('LOI300')
    # # loi400 = a[name].get('LOI400')
    # # print(loi_if_generator(name, loi200=loi200, loi300=loi300, loi400=loi400))
    # print(ifc_if_generator(name, ifc_list=['IfcPlate']))
