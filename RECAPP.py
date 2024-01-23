import xlrd
import pandas as pd
import json
import re
from datetime import datetime
import matplotlib.pyplot as plt
from matplotlib.backends.backend_pdf import PdfPages
import seaborn as sns

file_path = 'BRATINA.xls'

workbook = xlrd.open_workbook(file_path)

# Get len of columns думаю надо будет удалить эту строку
f = workbook.sheet_by_index(0).nrows

# Select the appropriate sheet (e.g., the first sheet)
sheet = workbook.sheet_by_index(0)

# Извлечение данных из колонки С начиная с 7 строки
column_data = [(row, 2, sheet.cell_value(row, 2)) for row in range(6, sheet.nrows)]

# Удаляем пустые значения в проверенных ячейках
non_empty_values_C = [(row + 1, col, value) for row, col, value in column_data if value != '']

# создаём DataFrame
df = pd.DataFrame({'Data': [value for row, col, value in non_empty_values_C],
                   'Row': [row for row, col, value in non_empty_values_C],
                   'Column': [col for row, col, value in non_empty_values_C]})

# Создаём словарик где ключи это название товаров, а значения это координаты где надо пройтись
data_dict = {}
for row, col, value in non_empty_values_C:
    if value not in data_dict:
        data_dict[value] = []
    data_dict[value].append((row, col))

# убирай из словаря значение о строке
data_dict = {key: [value[0] for value in values] for key, values in data_dict.items()}

# Новый словарь new_dict
map_dict = {}

# Проход по словарю data_dict
keys = list(data_dict.keys())
for i in range(len(keys)):
    key = keys[i]
    current_values = data_dict[key]

    # Если не последний элемент, добавляем следующее значение
    if i < len(keys) - 1:
        next_key = keys[i + 1]
        next_value = data_dict[next_key][0]
        map_dict[key] = current_values + [next_value]
    else:
        # Для последнего элемента добавляем текущее значение без следующего
        map_dict[key] = [data_dict[key][0]] + [f+1]

def extract_and_convert_date(input_string):
    # Используем регулярное выражение для извлечения даты после слова 'от'
    match = re.search(r'от (\d{1,2}\.\d{1,2}\.\d{4})', input_string)

    if match:
        date_string = match.group(1)
        try:
            date_object = datetime.strptime(date_string, '%d.%m.%Y')
            date_as_integer = int(date_object.strftime('%Y%m%d'))
            return date_as_integer
        except ValueError:
            print("Ошибка: Неверный формат даты")
            return None
    else:
        print("Ошибка: Дата не найдена")
        return None

def process_excel_data(file_path, ranges_dict):
    result_dict = {}

    workbook = xlrd.open_workbook(file_path)
    
    for key, value in ranges_dict.items():
        start_row, end_row = value
        data_list = []

        sheet = workbook.sheet_by_index(0)

        # Проверяем, что стартовая и конечная строки находятся в пределах допустимых значений
        if 0 <= start_row <= sheet.nrows and 0 <= end_row <= sheet.nrows:
            for row_number in range(start_row, end_row + 1):
                # Проверяем, что столбец D существует в данной строке
                if 0 <= 3 < sheet.ncols:
                    cell_value_D = str(sheet.cell_value(row_number - 1, 3))  # Assuming 0-indexed rows and columns
                    cell_date = extract_and_convert_date(cell_value_D)
                    
                    if 'Поступление товаров' in cell_value_D:
                        # Проверяем, что столбец F существует в данной строке
                        if 0 <= 5 < sheet.ncols:
                            cell_value_F = sheet.cell_value(row_number - 1, 5)
                            data_list.append((cell_date, cell_value_F))
                    elif 'Реализация товаров' in cell_value_D:
                        # Проверяем, что столбец G существует в данной строке
                        if 0 <= 6 < sheet.ncols:
                            cell_value_G = sheet.cell_value(row_number - 1, 6)
                            data_list.append((cell_date, -cell_value_G))  # Adding with a negative sign

            # Сортируем список по дате и значениям cell_value_G, cell_value_F
            sorted_data_list = sorted(data_list, key=lambda x: (x[0], x[1]))

            # Добавляем отсортированные значения в словарь
            result_dict[key] = [x[1] for x in sorted_data_list]

    return result_dict

# Получаем словарь c приходными и расходными накладными, продажи мы не считаем
sales_dict = process_excel_data(file_path, map_dict)

def split_list_by_negative(numbers):
    result = []
    current_sublist = []

    for num in numbers:
        if num < 0:
            # Отрицательное число, заканчиваем текущий подсписок и начинаем новый
            current_sublist.append(num)
            result.append(current_sublist)
            current_sublist = []
        else:
            # Вне зависимости от знака, добавляем число к текущему подсписку
            current_sublist.append(num)

    # Добавляем последний подсписок, если он не пустой
    if current_sublist:
        result.append(current_sublist)

    return result

def process_original_dict(original_dict):
    result_dict = {}

    for key, value in original_dict.items():
        result_dict[key] = split_list_by_negative(value)

    return result_dict

cycly_dict = process_original_dict(sales_dict)


def Average(lst): 
    return sum(lst) / len(lst)

def calculate_return_rate(sales_cycle):
    positive_lst = []
    negative_lst = []
    for i in sales_cycle:
        if i > 0:
            positive_lst.append(i)
        if i < 0:
            negative_lst.append(i)
    if negative_lst == []:
        return [0, Average(positive_lst)]
    if sum(positive_lst) != 0:
        return [abs(round(sum(negative_lst)/sum(positive_lst),2)), round(Average(positive_lst),2)]

# Мы получаем final_dict где в value первое число это вероятность возврата, а вторая это количество товара
final_dict = {}
for k,v in cycly_dict.items():
    value_lst = []
    for i in v:
        if calculate_return_rate(i) != None:
            value_lst.append([calculate_return_rate(i)[0],calculate_return_rate(i)[1]])
    if value_lst != []:
        final_dict[k] = value_lst
    value_lst = []

def start_sales_dict(sales_dict):
    result_dict = {}

    for k, v in sales_dict.items():
        avg_list = []  # Обнуляем список перед каждой итерацией
        if v is not None:
            for i in v:
                if i > 0:
                    avg_list.append(i)

        result_value = sum(avg_list) / len(avg_list) if len(avg_list) > 0 else 0
        result_dict[k] = round(result_value, 2) if result_value is not None else 0

    return result_dict


start_dict = start_sales_dict(sales_dict)


# Пример данных
data_dict = final_dict


# Создаем DataFrame
data = []
for pie_name, values_list in data_dict.items():
    min_quantity = min(values_list, key=lambda x: x[0])[1]
    data.append([pie_name, min_quantity, start_dict[pie_name]])

df = pd.DataFrame(data, columns=['Продукт', 'Рекомендуется', 'Заказывали'])

# Добавляем колонку "Рекомендация"
df['Рекомендация'] = df.apply(lambda row: 'Увеличить' if row['Рекомендуется'] > row['Заказывали'] + 0.1
                                            else 'Уменьшить' if row['Рекомендуется'] < row['Заказывали'] - 0.1
                                            else 'Оставить без изменений', axis=1)

# Записываем в Excel
df.to_excel('output.xlsx', index=False)

