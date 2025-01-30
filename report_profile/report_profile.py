from d_parser.d_parser import parse


def cond_generator(loi):
    res = []
    for i in loi:
        res.append(f'[{i}]<>""')
        res.append(f'[{i}]<>"0"')
    return ' and '.join(res)


def loi_if_generator(table_el_name, name_attr='TYPE', loi200=None, loi300=None, loi400=None):
    if loi200 and loi300 and loi400:
        return f'if([{name_attr}]="{table_el_name}", if({cond_generator(loi200)}, if({cond_generator(loi300)}, if({cond_generator(loi400)}, "LOI400", "LOI300"), "LOI200"), "LOI000"), "")'
    elif loi200 and loi300:
        return f'if([{name_attr}]="{table_el_name}", if({cond_generator(loi200)}, if({cond_generator(loi300)}, "LOI300", "LOI200"),"LOI000"), "")'
    elif loi200:
        return f'if([{name_attr}]="{table_el_name}", if({cond_generator(loi200)}, "LOI200","LOI000"), "")'


a = parse("../src/add_D.xlsx", to_term=True)
name = 'Объемный элемент'
loi200 = a[name].get('LOI200')
loi300 = a[name].get('LOI300')
loi400 = a[name].get('LOI400')
print(loi_if_generator(name, loi200=loi200, loi300=loi300, loi400=loi400))
# print(cond_generator(loi_ex))
