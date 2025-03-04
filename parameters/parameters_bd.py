import pyodbc
from d_parser.d_parser import parse
from config import SPECIALITIES


def fix_parameters_counter(crsr):
    query = "SELECT TOP 1 * FROM dbo.ParamDefs ORDER BY idParamDef DESC"
    crsr.execute(query)

    max_param_id = (crsr.fetchone()[0])

    update_query = f"UPDATE dbo.ParamDefs_IDENTITY SET id = {max_param_id}"
    crsr.execute(update_query)


def parameter_add(crsr, name_, caption_):
    fix_parameters_counter(crsr)
    crsr.execute("{CALL dbo.spParamDefUpdate (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)}",
                 ('1', name_, caption_, '-1', '0', '', '', '1', 'NULL', 'NULL', '-1', '1'))


def parameters_maker_db_full(crsr, source, param_group_name):
    # без интерфейса работает на основе парсинга приложения Д, см d_parser.py
    params = source['full_attr_list']
    for caption_, name_ in params.items():
        parameter_add(crsr, f'{name_}', f'{caption_}')
        param_id = crsr.fetchone()[0]
        cursor.execute("{CALL dbo.spParamDefAddCategory (?, ?, ?, ?)}",
                       (param_id, f'{param_group_name}', '2147483647', '2147483647'))


if __name__ == '__main__':
    import openpyxl

    conn = pyodbc.connect(f'DRIVER={{SQL Server}};SERVER=(local)\\SQLEXPRESS;DATABASE=clear;UID=sa;PWD=123')
    cursor = conn.cursor()

    workbook = openpyxl.load_workbook("../src/YS_2025_add_D_v20.xlsx")
    src = parse(workbook, to_term=True, to_json=False)

    parameters_maker_db_full(cursor, src, 'EngineeringDesign')
    conn.commit()
    cursor.close()
    conn.close()
