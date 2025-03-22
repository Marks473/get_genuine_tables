import os
import requests
import pandas as pd
from bs4 import BeautifulSoup

def is_relatable_table(table):
    """
    Проверка таблицы на реляционные данные:
    - Строки или столбцы должны быть однородными по содержанию тегов.
    - Возможны два варианта:
      1) С заголовками (проверяется первая строка или первый столбец)
      2) Без заголовков (проверяется как ранее)
    """
    # Получение всех строк таблицы
    rows = table.find_all('tr')

    if not rows:
        return False

    num_rows = len(rows)
    num_cols = len(rows[0].find_all(['th', 'td']))

    if num_cols == 0:
        return False

    # Проверка вертикальной таблицы
    vertical = True
    # Проверка первой строки на заголовки
    first_row_headers = True
    first_cell = rows[0].find_all(['th', 'td'])[0]
    first_cell_tag = first_cell.find()

    for col_index in range(1, num_cols):
        cell = rows[0].find_all(['th', 'td'])[col_index]
        if ((first_cell_tag and cell.find(first_cell_tag.name) is None) or 
            (not first_cell_tag and cell.find() is not None)):
            first_row_headers = False
            break

    # Проверка строк или столбцов в зависимости от наличия заголовков
    if first_row_headers:
        for col_index in range(1, num_cols):
            first_cell = rows[0].find_all(['th', 'td'])[col_index]
            tag_in_first_cell = first_cell.find()
            for row_index in range(1, num_rows):
                current_cell = rows[row_index].find_all(['th', 'td'])[col_index]
                if tag_in_first_cell and not current_cell.find(tag_in_first_cell.name):
                    vertical = False
                    break
            if not vertical:
                break
    else:
        for col_index in range(num_cols):
            if not vertical:
                break
            first_cell = rows[0].find_all(['th', 'td'])[col_index]
            tag_in_first_cell = first_cell.find()
            for row_index in range(1, num_rows):
                current_cell = rows[row_index].find_all(['th', 'td'])[col_index]
                if tag_in_first_cell and not current_cell.find(tag_in_first_cell.name):
                    vertical = False
                    break
            if not vertical:
                break

    if vertical:
        return True
    
    # Проверка горизонтальной таблицы
    horizontal = True
    # Проверка первого столбца на заголовки
    first_col_headers = True
    first_cell = rows[0].find_all(['th', 'td'])[0]
    first_cell_tag = first_cell.find()

    for row_index in range(1, num_rows):
        cell = rows[row_index].find_all(['th', 'td'])[0]
        if ((first_cell_tag and cell.find(first_cell_tag.name) is None) or 
            (not first_cell_tag and cell.find() is not None)):
            first_col_headers = False
            break

    if first_col_headers:
        first_row = rows[1].find_all(['th', 'td'])
        for row_index in range(2, num_rows):
            if not horizontal:
                break
            for col_index in range(num_cols):
                current_cell = rows[row_index].find_all(['th', 'td'])[col_index]
                if (first_row[col_index].find() is not None and 
                    current_cell.find(first_row[col_index].find().name) is None):
                    horizontal = False
                    break
            if not horizontal:
                break
    else:
        first_row = rows[0].find_all(['th', 'td'])
        for row_index in range(1, num_rows):
            if not horizontal:
                break
            for col_index in range(num_cols):
                current_cell = rows[row_index].find_all(['th', 'td'])[col_index]
                if (first_row[col_index].find() is not None and 
                    current_cell.find(first_row[col_index].find().name) is None):
                    horizontal = False
                    break
            if not horizontal:
                break

    return horizontal

def download_html(url):
    """Скачивание HTML содержимого по заданному URL."""
    response = requests.get(url)
    response.raise_for_status()  # Проверка успешности запроса
    return response.text

def convert_html_to_excel(html_content, excel_file):
    try:
        # Парсинг HTML с помощью BeautifulSoup
        soup = BeautifulSoup(html_content, 'html.parser')

        # Нахождение всех таблиц
        tables = soup.find_all('table')

        # Фильтрация реляционных таблиц
        relatable_tables = [table for table in tables if is_relatable_table(table)]

        # Проверка наличия реляционных таблиц в HTML
        if len(relatable_tables) == 0:
            raise ValueError("В HTML-файле не найдено соответствующих таблиц.")

        # Преобразование реляционных таблиц в DataFrame и экспорт их в Excel
        with pd.ExcelWriter(excel_file) as writer:
            for i, table in enumerate(relatable_tables):
                df = pd.read_html(str(table))[0]
                df.to_excel(writer, sheet_name=f'Table_{i+1}', index=False)

        print(f"Таблицы успешно экспортированы в {excel_file}")
    except Exception as e:
        print(f"Произошла ошибка: {e}")

# Пример использования
url = 'http://voencentre.isu.ru/ru/staff'  # Введите URL-адрес HTML-страницы
html_content = download_html(url)
excel_file = 'table.xlsx'

convert_html_to_excel(html_content, excel_file)

os.startfile(excel_file)
