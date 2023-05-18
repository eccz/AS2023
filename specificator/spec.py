from library_manager.libr import indent
import xml.etree.ElementTree as ET
import csv


def csv_input(name='ELEC.csv'):  # указать в дальнейшем переменную
    filepath = fr'work/src/{name}'
    result = dict()

    with open(filepath, 'r', encoding='cp1251', newline='') as file:
        reader = csv.reader(file, delimiter=';')
        for row in reader:
            if row[2] in result:
                result[row[2]].update({row[1]: row[0]})
            elif not row[0]:
                continue
            else:
                result[row[2]] = {row[1]: row[0]}
    return result


def xml_dataset_build():
    dataset_args = dict(assemblyGrouping="0", assemblyFilter="2", binding="Fields", relationType="", join="outer",
                        hierarchy="0")

    return ET.Element('Dataset', **dataset_args)


def xml_table_build(elem_type='Фундаментная_подушка', filter_param='[EHP_TYPE]'):
    table_args = {"caption": f"{elem_type}", "filter": f'{filter_param} = f"{elem_type}"',
                  "result.filter": "", "aggregated": "0"}

    return ET.Element('Table', **table_args)


def types_build():
    types = ET.Element('Types')
    for word in open('work/base_types', 'r', encoding='utf-8'):
        type_ = ET.Element('Type', name=word.strip())
        types.append(type_)

    return types


def xml_field_build(elem_type='Фундаментная_подушка', elem_param='[EHP_TYPE]'):
    field_args = dict(caption=f"{elem_type}", data=f"{elem_param}", type="0", aggregate="0", visible="1", format="")

    return ET.Element('Dataset', **field_args)


def xml_build(base=True):
    root = ET.Element('Report')
    dataset_profile = ET.Element('DatasetProfile')
    dataset = xml_dataset_build()

    indent(root)
    return root

    # rank = ET.Element('rank', updated="yes")
    # rank.text = '2'
    # country.append(rank)
    # gdppc = ET.Element('gdppc')
    # gdppc.text = '141100'
    # country.append(gdppc)
    #
    # country.append(ET.Element('neighbor', name="Austria", direction="E"))
    # country.append(ET.Element('neighbor', name="Switzerland", direction="W"))
    #
    # root.append(country)


def main():
    source = csv_input()
    for element in source:


if __name__ == '__main__':
    main()
    # mydata = ET.tostring(xml_build(), encoding="utf-8", method="xml", xml_declaration=True)
    # print(mydata.decode(encoding="utf-8"))
