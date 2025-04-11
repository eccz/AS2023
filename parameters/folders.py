import pyodbc
from d_parser.d_parser import parse
from config import SPECIALITIES, SPECIALITY_ATTR_NAME, TYPE_ATTR_NAME


def fix_folders_counter(crsr):
    max_param_id = get_counter(crsr, 'dbo.Folders', 'idFolder')
    crsr.execute(f"UPDATE dbo.Folders_IDENTITY SET id = {max_param_id}")


def delete_folder(crsr, id_folder):
    crsr.execute("{CALL dbo.spDeleteFolder (?)}", id_folder)
    fix_folders_counter(crsr)


def get_param_def_id_by_name(crsr, param_name):
    try:
        param_id_query = f"SELECT idParamDef FROM dbo.ParamDefs WHERE Name = '{param_name}'"
        crsr.execute(param_id_query)
        return crsr.fetchone()[0]
    except Exception as ex:
        print(ex)


def get_counter(crsr, table_name, col_name):
    try:
        query = f"SELECT TOP 1 * FROM {table_name} ORDER BY {col_name} DESC"
        crsr.execute(query)
        return crsr.fetchone()[0]
    except Exception as ex:
        print(ex)


def filter_condition_maker(crsr, element, loi_level):
    type_attr = TYPE_ATTR_NAME
    type_value = element[type_attr]
    type_param_id = get_param_def_id_by_name(crsr, type_attr)

    spec_attr = SPECIALITY_ATTR_NAME
    spec_value = element[spec_attr]
    spec_param_id = get_param_def_id_by_name(crsr, spec_attr)

    counter = 3
    res = []
    param_list = set()

    if loi_level == 'LOI200':
        param_list.update(element['LOI200'])
    elif loi_level == 'LOI300':
        param_list.update(element['LOI200'])
        param_list.update(element['LOI300'])
    elif loi_level == 'LOI400':
        param_list.update(element['LOI200'])
        param_list.update(element['LOI300'])
        param_list.update(element['LOI400'])

    param_list = [i for i in param_list]

    counter += len(param_list)
    print(len(param_list))
    parameters = ' , '.join([f'Parameters_STR PT{i}' for i in range(counter)])

    for n in range(counter - 3):
        id_param = get_param_def_id_by_name(crsr, param_list[n])
        res.append(f"(PT{n}.idObject = O.idObject) AND (PT{n}.idParamDef = {id_param}) AND (PT{n}.Value<>'')")

    type_cond = f"""(PT{range(counter)[-3]}.idObject = O.idObject) AND (PT{range(counter)[-3]}.idParamDef = {type_param_id}) AND (PT{range(counter)[-3]}.Value='{type_value}')"""
    spec_cond = f"""(PT{range(counter)[-2]}.idObject = O.idObject) AND (PT{range(counter)[-2]}.idParamDef = {spec_param_id}) AND (PT{range(counter)[-2]}.Value='{spec_value}')"""
    specs_cond = f"""(PT{range(counter)[-1]}.idObject = O.idObject) AND (PT{range(counter)[-1]}.idParamDef = {spec_param_id}) AND (PT{range(counter)[-1]}.Value IN ({','.join([f"'{i}'" for i in SPECIALITIES])}) ) """
    ender = f'(O.idParentObject is null)'
    res.extend([type_cond, spec_cond, specs_cond, ender])

    return f", {parameters} where {' AND '.join(res)}"


def get_parameters_db(crsr):
    crsr.execute("{CALL dbo.spGetParamDefs}")

    return {i[2]: i[0] for i in crsr.fetchall()}


def create_filter_source(crsr, id_folder, id_parameter, condition, value):
    counter = get_counter(crsr, table_name='dbo.FilterSource', col_name='idFilterSource')

    query = f"""
        INSERT INTO dbo.FilterSource (idFilterSource, idFolder, idParameter, Condition, Value, IsLinkedObjectParam, ObjectLinkDirection, idObjectRelationType, idLinkedObjectCategory, ConditionTarget) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """
    update_query = f"UPDATE dbo.FilterSource_IDENTITY SET id = {counter + 1}"

    crsr.execute(query, (counter + 1, id_folder, id_parameter, condition, value, 0, None, None, None, 0))
    crsr.execute(update_query)


def create_main_pset_folder(crsr, folder_name):
    param_id = get_param_def_id_by_name(crsr, SPECIALITY_ATTR_NAME)
    fix_folders_counter(crsr)
    counter = get_counter(crsr, 'dbo.Folders_IDENTITY', 'id')

    filter_value = f""", Parameters_STR PT0 where (PT0.idObject = O.idObject) AND (PT0.idParamDef = {param_id}) AND (PT0.Value IN ({','.join([f"'{i}'" for i in SPECIALITIES])}) ) AND (O.idParentObject is null)"""

    insert_query = f"""
        INSERT INTO dbo.Folders (idFolder, idParentFolder, Name, Filter, idDirectory, FolderFlags, ResultSetType, ParentFilter) VALUES (?, ?, ?, ?, ?, ?, ?, ?)
        """

    update_query = f"UPDATE dbo.Folders_IDENTITY SET id = {counter + 1}"
    crsr.execute(insert_query, (counter + 1, None, folder_name, filter_value, None, '0', '0', None))
    create_filter_source(crsr, counter + 1, param_id, 'один из', f"{';'.join(SPECIALITIES)}")

    crsr.execute(update_query)


if __name__ == '__main__':
    import openpyxl

    conn = pyodbc.connect(f'DRIVER={{SQL Server}};SERVER=(local)\\SQLEXPRESS;DATABASE=clear;UID=sa;PWD=123')
    cursor = conn.cursor()
    workbook = openpyxl.load_workbook("../src/YS_2025_add_D_v20.xlsx")
    src = parse(workbook, to_term=True, to_json=False)

    print(filter_condition_maker(cursor, src['Таблица 1.0.4'], 'LOI400'))
    create_main_pset_folder(cursor, '!EngineeringDesign')
    # delete_folder(cursor, 29)
    cursor.commit()
    conn.close()
