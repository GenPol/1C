import xlrd
import pandas as pd
import json
import re
from datetime import datetime
import matplotlib.pyplot as plt
from matplotlib.backends.backend_pdf import PdfPages
import seaborn as sns
from collections import Counter
from sklearn.linear_model import LinearRegression
import numpy as np

# указываем путь к файлу
file_path = 'BRATINA.xls'
# открываем XLS файл
workbook = xlrd.open_workbook(file_path)
# открываем нужный лист
sheet = workbook.sheet_by_index(0)

# Извлечение данных из колонки С начиная с 7 строки
column_data = [(row, 2, sheet.cell_value(row, 2)) for row in range(6, sheet.nrows)]

# Удаляем пустые значения в проверенных ячейках
non_empty_values_C = [(row + 1, col, value) for row, col, value in column_data if value != '']

# Последняя координата где будет поиск
last_number = workbook.sheet_by_index(0).nrows

# Создаём словарь с продуктами и их координатами
data_dict = {}
for i, (row, col, value) in enumerate(non_empty_values_C):
    if value not in data_dict:
        data_dict[value] = []

    # Добавляем текущий row в список
    data_dict[value].append(row)

    # Если это не последний элемент списка и col совпадает с col следующего элемента,
    # добавляем row следующего элемента в тот же список
    if i < len(non_empty_values_C) - 1 and col == non_empty_values_C[i + 1][1]:
        data_dict[value].append(non_empty_values_C[i + 1][0])
    else:
        # Если это последний элемент, добавляем last_number в список
        data_dict[value].append(last_number)

# Найти дату и представить её в виде числа
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

# Поиск поступления и возвратов товара (В других версиях скрипта модифицировать!!!)
def process_excel_data(file_path, ranges_dict):
    result_dict = {}

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
sales_dict = process_excel_data(file_path, data_dict)

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



def predict_sales(data):
    X = np.array([i+1 for i in range(len(data))]).reshape(-1, 1)  # Добавляем столбец дней как признаки
    y = np.cumsum(data)  # Создаем кумулятивные продажи (сумму по дням)
    
    model = LinearRegression()
    model.fit(X, y)
    
    # Предсказываем будущие продажи
    future_days = np.array([i+len(data)+1 for i in range(3)]).reshape(-1, 1)  # Предсказание на 3 будущих дня
    future_sales = model.predict(future_days)
    
    # Преобразуем предсказанные значения в разницу между продажами на последующие дни
    future_diff = np.diff(future_sales)
    
    # Убираем отрицательные значения
    future_diff[future_diff < 0] = 0
    
    return future_diff

# Пример использования
sales_data = {'бургер': [1, 2, 4, -1, 2, 2, -1], 'пицца': [1, 2, 2, 2, -2, 1, 2, -1]}
predicted_sales = {product: predict_sales(data) for product, data in sales_data.items()}

print(predicted_sales)



