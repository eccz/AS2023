import csv
import xml.etree.ElementTree as ET


def csv_input(filepath):
    result = []

    with open(filepath, 'r', encoding='cp1251', newline='') as file:
        reader = csv.reader(file, delimiter=';')
        [result.append(row[2].strip()) for row in reader if row[2].strip() not in result]
    return result


def struct_element_build(name):
    element_args = dict(name=f"{name}", id="0", uid="0")

    return ET.Element('Element', **element_args)


def struct_parameters_build():
    parameters = ET.Element('Parameters')
    parameter_1_args = dict(name="PROJECT_STRUCT_LEVEL", value="1.  Стадия", caption="Уровень иерархии", comment="")
    parameter_2_args = dict(name="SYS_CATEGORY_GROUP", value="Разделы проекта", caption="Группа данных",
                            comment="DisciplinesHierarchy")
    parameter_3_args = dict(name="USER_ACCESS_GROUPS", value="", caption="Доступ", comment="")

    parameters.append(ET.Element('Parameter', **parameter_1_args))
    parameters.append(ET.Element('Parameter', **parameter_2_args))
    parameters.append(ET.Element('Parameter', **parameter_3_args))

    return parameters


def struct_xml_build():
    pass


def main():
    a = struct_element_build('123')
    a.append(struct_parameters_build())
    mydata = ET.tostring(a, encoding="utf-8", method="xml")
    print(mydata.decode(encoding="utf-8"))


if __name__ == '__main__':
    main()
