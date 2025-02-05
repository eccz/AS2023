import config
from config import TASK_TYPE_ATTR_NAME, TYPE_ATTR_NAME, SPECIALITY_ATTR_NAME
from d_parser.d_parser import parse
import xml.etree.ElementTree as ET
from ifc_import.ifc_import import indent


def cond_generator(loi):
    """
    Функция генерирует if-условие для CadLib на базе списка наименований атрибутов для определенного LOI
    """
    res = []
    for i in loi:
        res.append(f'[{i}]<>""')
        res.append(f'[{i}]<>"0"')
    return ' and '.join(res)


def loi_if_generator(type_attr_name, type_attr, loi200=None, loi300=None, loi400=None):
    """
    Функция генерирует полное if-условие для CadLib по проверке LOI, с учетом наличия требований на LOI300 и LOI400.
    type_attr_name - наименование атрибута CadLib для фильтра, например TYPE
    type_attr - значение атрибута "TYPE" для фильтра в CadLib
    """
    if loi200 and loi300 and loi400:
        return f'if([{type_attr_name}]="{type_attr}", if({cond_generator(loi200)}, if({cond_generator(loi300)}, if({cond_generator(loi400)}, "LOI400", "LOI300"), "LOI200"), "LOI000"), "")'
    elif loi200 and loi300:
        return f'if([{type_attr_name}]="{type_attr}", if({cond_generator(loi200)}, if({cond_generator(loi300)}, "LOI300", "LOI200"),"LOI000"), "")'
    elif loi200:
        return f'if([{type_attr_name}]="{type_attr}", if({cond_generator(loi200)}, "LOI200","LOI000"), "")'


def ifc_if_generator(ifc_attr_name=config.IFC_ATTR_NAME, asm_attr_name=config.ASM_TYPE_ATTR_NAME, ifc_list=None):
    """
    Функция генерирует полное if-условие для CadLib по проверке правильности IFC-класса.
    Учитывает результаты парсинга приложения Д - учитывает требования по сборкам.
    ifc_list - список списков для разных требований по сборкам либо список с одним элементом.
    ifc_attr_name - наименование атрибута CadLib, содержащего класс IFC
    type_attr - значение атрибута "TYPE" для фильтра в CadLib
    asm_attr_name - наименование атрибута CadLib, содержащего тип сборки.
    """
    res = []
    if isinstance(ifc_list[0], list):
        for req in ifc_list:
            res.append(f'instr("{req[1]}", [{ifc_attr_name}])>=0 and [{asm_attr_name}]="{req[2]}"')
        res = ' or '.join(res)
        return f'if({res}, "CORRECT", "INCORRECT")'

    elif isinstance(ifc_list[0], str):
        return f'if(instr("{ifc_list[0]}", [IFC_TYPE])>=0, "CORRECT", "INCORRECT")'


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
    ns = {'caption': f"{table_num} [{_type}]".strip(), 'filter': f'"[{type_attr}]="{_type}"', 'result.filter': "",
          'aggregated': "0"}
    return ET.Element('Table', **ns)


def report_types_build():
    types = ET.Element('Types')
    for i in config.REPORT_TYPES:
        types.append(ET.Element('Type', name=i))

    return types


def report_field_build(caption, data, unknown=False):
    if unknown:
        return ET.Element('Field', caption=caption, data=data, type="1", aggregate="0", visible="1", format='')
    return ET.Element('Field', caption=caption, data=data, type="0", aggregate="0", visible="1", format='')


def report_fields_build(table_data, loi_data, ifc_flag_data, unknown=False):
    fields = ET.Element('Fields')
    fields.append(report_field_build(caption="SYS_OBJECT_CATEGORY", data="SYS_OBJECT_CATEGORY"))
    fields.append(report_field_build(caption="SYS_OBJECT_NAME", data="@NAME"))
    fields.append(report_field_build(caption="IFC_TYPE", data="IFC_TYPE"))
    fields.append(report_field_build(caption="IfcGlobalId", data="IfcGlobalId"))
    fields.append(report_field_build(caption="TABLE", data=table_data, unknown=unknown))
    fields.append(report_field_build(caption="LOI", data=loi_data, unknown=unknown))
    fields.append(report_field_build(caption="IFC_TYPE_FLAG", data=ifc_flag_data, unknown=unknown))
    fields.append(report_field_build(caption=SPECIALITY_ATTR_NAME, data=SPECIALITY_ATTR_NAME))
    fields.append(report_field_build(caption=TYPE_ATTR_NAME, data=TYPE_ATTR_NAME))
    fields.append(report_field_build(caption=TASK_TYPE_ATTR_NAME, data=TASK_TYPE_ATTR_NAME))
    fields.append(report_field_build(caption=config.ASSEMBLY_TYPE_ATTR_NAME, data=config.ASSEMBLY_TYPE_ATTR_NAME))
    fields.append(report_field_build(caption=config.TAG_ATTR_NAME, data=config.TAG_ATTR_NAME))
    fields.append(report_field_build(caption=config.NAME_ATTR_NAME, data=config.NAME_ATTR_NAME))
    fields.append(report_field_build(caption=config.ASSEMBLY_MARK_ATTR_NAME, data=config.ASSEMBLY_MARK_ATTR_NAME))
    fields.append(report_field_build(caption=config.ASSEMBLY_N_ATTR_NAME, data=config.ASSEMBLY_N_ATTR_NAME))
    fields.append(report_field_build(caption=config.PROJECT_MARK_ATTR_NAME, data=config.PROJECT_MARK_ATTR_NAME))
    fields.append(report_field_build(caption=config.SYSTEM_TAG_ATTR_NAME, data=config.SYSTEM_TAG_ATTR_NAME))
    fields.append(report_field_build(caption=config.EL_LINE_TAG_ATTR_NAME, data=config.EL_LINE_TAG_ATTR_NAME))
    fields.append(report_field_build(caption=config.TASK_AUTHOR_ATTR_NAME, data=config.TASK_AUTHOR_ATTR_NAME))
    fields.append(report_field_build(caption=config.TASK_USER_ATTR_NAME, data=config.TASK_USER_ATTR_NAME))
    return fields


def report_view_build():
    view = ET.Element('View')
    view.append(ET.Element('GroupFields'))
    view.append(ET.Element('SortFields'))
    return view


def unknown_dataset_build(types_list):
    unknown_dataset = report_dataset_build()
    ftr = ' and '.join([f'[{i}]<>"{j}"' for i, j in types_list])
    ns = {'caption': "[НЕИЗВЕСТНЫЙ_ТИП]", 'filter': ftr, 'result.filter': "", 'aggregated': "0"}
    table = ET.Element('Table', **ns)
    table.append(report_types_build())
    table.append(
        report_fields_build(table_data='"UNKNOWN"', loi_data='"UNKNOWN"', ifc_flag_data='"UNKNOWN"', unknown=True))
    unknown_dataset.append(table)

    return unknown_dataset


def xml_report_profile_build(source):
    types_list = []
    types = report_types_build()

    root = report_root_build()
    report_dataset_profile = report_dataset_profile_build()

    for k, v in source.items():
        dataset = report_dataset_build()

        table_num = v['table_num']
        type_attr_name = config.TYPE_ATTR_NAME if v.get(config.TYPE_ATTR_NAME) else config.TASK_TYPE_ATTR_NAME
        type_attr = v[type_attr_name]
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
                                     unknown=False)

        table.append(fields)
        dataset.append(table)
        dataset.append(report_view_build())
        report_dataset_profile.append(dataset)

    report_dataset_profile.append(unknown_dataset_build(types_list))
    root.append(report_dataset_profile)

    root.append(report_format_build())
    root.append(report_extended_build())

    indent(root)
    mydata = ET.tostring(root, encoding="utf-8", method="xml")

    with open(r'report_profile.xml', 'w', encoding="utf-8") as f:
        f.write('<?xml version="1.0"?>\n')
        f.write(mydata.decode(encoding="utf-8"))


if __name__ == '__main__':
    src = parse("../src/add_D.xlsx", to_term=True, to_json=False)
    xml_report_profile_build(src)

    # a = parse("../src/add_D.xlsx", to_term=True)
    # name = 'Объемный элемент'
    # # loi200 = a[name].get('LOI200')
    # # loi300 = a[name].get('LOI300')
    # # loi400 = a[name].get('LOI400')
    # # print(loi_if_generator(name, loi200=loi200, loi300=loi300, loi400=loi400))
    # print(ifc_if_generator(name, ifc_list=['IfcPlate']))
