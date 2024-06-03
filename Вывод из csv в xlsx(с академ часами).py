import pandas as pd
from openpyxl import Workbook

# Загрузите входной файл "Преподы все.csv" с указанием кодировки
input_file = "Преподаватели_общая.csv"

# Указываем кодировку файла
encoding = 'cp1251'

# Создайте выходной файл "Преподы выход.xlsx"
output_file = "Преподаватели_выход.xlsx"

# Чтение входного файла CSV с указанием кодировки
data = pd.read_csv(input_file, delimiter=";", encoding=encoding)

# Создайте книгу Excel
workbook = Workbook()

# Получите уникальные имена преподавателей
teachers = data['Преподаватель'].unique()

# Для каждого преподавателя создайте свой лист
for teacher in teachers:
    teacher_data = data[data['Преподаватель'] == teacher]

    # Группировка данных по уникальным строкам и суммирование академических часов
    grouped_data = teacher_data.groupby(["Группа", "Дисциплина"]).sum().reset_index()

    # Создайте лист с именем преподавателя
    sheet = workbook.create_sheet(title=teacher)

    # Запишите данные на лист и подсчитайте сумму академических часов для каждой уникальной строки
    sheet.append(["Группа", "Дисциплина", "Академические часы"])
    for index, row in grouped_data.iterrows():
        sheet.append([row["Группа"], row["Дисциплина"], row["Академические часы"]])

# Удалите лист "Sheet", созданный по умолчанию
workbook.remove(workbook['Sheet'])

# Сохраните выходной файл
workbook.save(output_file)