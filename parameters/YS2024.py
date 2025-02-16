import csv

def csv_input():
    result = dict()
    filepath = 'zadanie_YS2025.csv'
    with open(filepath, 'r', encoding='cp1251', newline='') as file:
        reader = csv.reader(file, delimiter=';')
        for row in reader:
            result[row[1]] = row[0]
    return result

# Генерирует запрос для таблицы ParamDefs для перебивки параметров из AS2024 на новые.
# В файле zadanie_2025.csv параметры из AS2024. Код ниже добавляет к ним префикс ED. Параметры с RD перебиваются потом вручную.
# Написано на коленке в день проверки.

if __name__ == '__main__':
    with open('res.txt', 'w', encoding='utf-8') as f:

        for i, j in csv_input().items():
            f.write('UPDATE dbo.ParamDefs\n')
            f.write('SET\n')
            f.write(f"    Name = 'ED_{j}',\n")
            f.write(f"    Caption = 'ED_{j}'\n")
            f.write(f"WHERE Name = '{j}';\n")
            print(f'{i} ------ {j}')