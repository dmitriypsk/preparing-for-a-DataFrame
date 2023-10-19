import pandas as pd
import re

def is_date(value):
    """Проверяет наличие даты в формате 'dd.mm.yyyy'"""
    if not isinstance(value, str):
        return False
    return bool(re.match(r"^\s*\d{2}\.\d{2}\.\d{4}\s*$", value))

def process_data(file_path):
    # Загрузка данных
    data = pd.read_excel(file_path, sheet_name='Исходные данные')
    
    # Инициализация переменных
    current_fio, current_polis, current_date = None, None, None
    
    for idx, row in data.iterrows():
        if "Полис №" in str(row['№']):
            current_fio, _, current_polis = row['№'].partition("Полис №")
            current_date = None
        elif is_date(row['Диагноз']):
            current_date = row['Диагноз'].strip()
        elif pd.notna(row['Код \nуслуги']):
            data.at[idx, 'Пациент ФИО'] = current_fio
            data.at[idx, 'Пациент полис'] = current_polis
            data.at[idx, 'Дата оказания услуги'] = current_date
    
    # Формирование результата
    result = data[data['Код \nуслуги'].notna()]
    columns = ['Пациент ФИО', 'Пациент полис', 'Дата оказания услуги', 
               'Код \nуслуги', 'Диагноз', 'Название услуги', 
               'Кол-\nво', 'Цена по прейскуранту', 'Сумма со скидкой']
    return result[columns]

# Путь к файлу 
file_path = "/Users/dmitriy/Downloads/задание_1_реестр.xlsx"
result_data = process_data(file_path)

# Сохранение результата
output_path = "/Users/dmitriy/Downloads/результат_реестр.xlsx"
result_data.to_excel(output_path, index=False)

# Отображение первых 10 строк результата
result_data.head(10)
