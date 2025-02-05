import openpyxl
import json
from string import ascii_letters
from datetime import datetime
from config import SHEET_NAMES, TASK_TYPE_ATTR_NAME, TYPE_ATTR_NAME


def el_sheet_processor(sheet):
    elements = {}
    element = {}
    element_name = ''

    cnt = 1
    for row in sheet.iter_rows(min_row=0, max_row=sheet.max_row, min_col=1, max_col=sheet.max_column, values_only=True):
        # print(row[0])
        if isinstance(row[0], str):
            if row[0].startswith('Таблица'):
                table_num = row[0].strip()
                element_name = row[1].strip()
                element = {element_name: {'table_num': table_num, 'IFC': [], 'LOI200': [], 'LOI300': [], 'LOI400': [], 'el_attr_list': []}}
                # print(element)

            if row[0].startswith('Класс IFC'):
                ifc = [i.strip() for i in row[2].replace(',', '').strip().split(' ') if
                       all(map(lambda c: c in ascii_letters, i))]
                element[element_name]['IFC'].extend(ifc)

            if 'атрибут' in row[0].lower():
                for row_2 in sheet.iter_rows(min_row=cnt + 1, max_row=sheet.max_row, min_col=1,
                                             max_col=sheet.max_column,
                                             values_only=True):

                    if not isinstance(row_2[0], str): break
                    if row_2[0].startswith('Таблица') or row_2[0].startswith('Примечание'): break
                    if row_2[1].strip() == TYPE_ATTR_NAME or row_2[1].strip() == TASK_TYPE_ATTR_NAME:
                        element[element_name][row_2[1].strip()] = row_2[2].strip()

                    if isinstance(row_2[2], str):
                        element[element_name]['LOI200'].append(row_2[1])
                        # element[element_name]['LOI300'].append(row_2[1])
                        # element[element_name]['LOI400'].append(row_2[1])

                    if isinstance(row_2[3], str):
                        element[element_name]['LOI300'].append(row_2[1])

                    if isinstance(row_2[4], str) and not row_2[4].startswith('-'):
                        element[element_name]['LOI400'].append(row_2[1])

                    # используется для составления списка параметров
                    element[element_name]['el_attr_list'].append([row_2[1], row_2[0]])

                if len(element[element_name]['LOI400']) < len(element[element_name]['LOI300']):
                    element[element_name].pop('LOI400')
                elements.update(element)

                # print(element)

        cnt += 1

    return elements


def gr_sheet_processor(sheet, el_name):
    res = []
    for row in sheet.iter_rows(min_row=0, max_row=sheet.max_row, min_col=1, max_col=sheet.max_column, values_only=True):
        if isinstance(row[2], str) and row[2].strip() == el_name:
            res.append([el_name, row[5].strip(), row[12].strip()])
    return res
    # print(res)


def parse(input_file, to_json=False, to_term=False):
    res = {}
    wb = openpyxl.load_workbook(input_file)
    for sheet in wb.sheetnames:
        if sheet in SHEET_NAMES:
            res.update(el_sheet_processor(wb[sheet]))

    for el_name, el_props in res.items():
        # print(el_name, el_props)
        # print(len(el_props['IFC']))
        if len(el_props['IFC']) > 1:
            el_props['IFC'] = gr_sheet_processor(wb['AEC_GR'], el_name)

    if to_term:
        print(res)

    if to_json:
        dt = datetime.now().strftime("file_%Y-%m-%d_%H-%M-%S")
        with open(f'{dt}_output.json', 'w', encoding='utf-8') as outfile:
            json.dump(res, outfile, ensure_ascii=False, indent=4)

    return res


if __name__ == '__main__':
    parse("../src/add_D.xlsx", to_term=True)
