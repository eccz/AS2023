import xml.etree.ElementTree as ET


def indent(elem, level=0):
    i = "\n" + level * "  "
    if len(elem):
        if not elem.text or not elem.text.strip():
            elem.text = i + "  "
        if not elem.tail or not elem.tail.strip():
            elem.tail = i
        for elem in elem:
            indent(elem, level + 1)
        if not elem.tail or not elem.tail.strip():
            elem.tail = i
    else:
        if level and (not elem.tail or not elem.tail.strip()):
            elem.tail = i


file = r'C:\Users\user\Desktop\Мои xml\LIBR_TEST_AS2023.xml'

parameter = ET.Element('PARAMETER', type="1", prescision="-1", readonly="0", accuracy="-1", valueType="1",
                       sysNameMeasureUnitBase="", sysNameMeasureUnit="")

name = ET.Element('NAME')
name.text = 'CIRC_SERVICE_SETTING_1337'
parameter.append(name)

default = ET.Element('DEFAULT')
parameter.append(default)

categories = ET.Element('CATEGORIES')
category = ET.Element('CATEGORY', order="2147483647", categoryOrder="81")
category.text = 'Конструкция ограждения_1337'
categories.append(category)
parameter.append(categories)

indent(parameter)

xml_str = ET.tostring(parameter, encoding="utf-8", method="xml")
print(xml_str.decode(encoding="utf-8"))
