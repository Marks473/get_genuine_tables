import os
import requests
import pandas as pd
from bs4 import BeautifulSoup


def get_leaf_tags(tag):
    """Рекурсивно извлекает все листья из тега."""
    if not tag.find_all():
        return [tag] if tag.text.strip() else []
    leaves = []
    for child in tag.find_all(recursive=False):
        leaves.extend(get_leaf_tags(child))
    return leaves


def split_cell(cell, vertical=True):
    """
    Разделяет содержимое ячейки на несколько листьев тегов и возвращает список строк или ячеек.
    """
    leaf_tags = get_leaf_tags(cell)
    texts = []
    images = []
    for tag in leaf_tags:
        if tag.name == 'img' and 'src' in tag.attrs:
            images.append(tag['src'])
        elif tag.text.strip():
            if vertical:
                texts.append(tag.text.strip())
            else:
                texts.append(tag)

    if vertical:
        return texts + images
    else:
        return texts + images


def expand_table_with_tags(table, vertical=True):
    """
    Расширяет таблицу, добавляя новые строки или столбцы, если в ячейках есть несколько тегов.
    """
    rows = table.find_all('tr')
    expanded_data = []

    for row in rows:
        cells = row.find_all(['th', 'td'])
        expanded_row = [[] for _ in range(len(cells))]

        for i, cell in enumerate(cells):
            split_cells = split_cell(cell, vertical)
            for j, split_cell_item in enumerate(split_cells):
                if len(expanded_row) <= j:
                    # Добавить новые строки или столбцы по мере необходимости
                    expanded_row.extend([[] for _ in range(j - len(expanded_row) + 1)])
                if vertical:
                    expanded_row[j].append(split_cell_item)
                else:
                    expanded_row[j].append(split_cell_item)

        if vertical:
            if len(expanded_data) < len(expanded_row):
                expanded_data.extend([[] for _ in range(len(expanded_row) - len(expanded_data))])
            for idx, item in enumerate(expanded_row):
                expanded_data[idx].extend(item)
        else:
            expanded_data.extend(expanded_row)

    return expanded_data


def is_relatable_table(table):
    """
    Проверка таблицы на реляционные данные:
    - Строки или столбцы должны быть однородными по содержанию тегов.
    - Возможны два варианта:
      1) С заголовками (проверяется первая строка или первый столбец)
      2) Без заголовков (проверяется как ранее)
    """
    rows = table.find_all('tr')

    if not rows:
        return False

    num_rows = len(rows)
    num_cols = len(rows[0].find_all(['th', 'td']))

    if num_cols == 0:
        return False

    vertical = True
    first_row_headers = True
    first_cell = rows[0].find_all(['th', 'td'])[0]
    first_cell_tag = first_cell.find()

    for col_index in range(1, num_cols):
        cell = rows[0].find_all(['th', 'td'])[col_index]
        if ((first_cell_tag and cell.find(first_cell_tag.name) is None) or
                (not first_cell_tag and cell.find() is not None)):
            first_row_headers = False
            break

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
        expanded_data = expand_table_with_tags(table, vertical=True)
        return expanded_data

    horizontal = True
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

    if horizontal:
        expanded_data = expand_table_with_tags(table, vertical=False)
        return expanded_data

    return False


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
                expanded_data = is_relatable_table(table)
                df = pd.DataFrame(expanded_data)
                df.to_excel(writer, sheet_name=f'Table_{i + 1}', index=False)

        print(f"Таблицы успешно экспортированы в {excel_file}")
    except Exception as e:
        print(f"Произошла ошибка: {e}")


# Пример использования
url = 'http://voencentre.isu.ru/ru/staff'  # Введите URL-адрес HTML-страницы
html_content = download_html(url)
excel_file = 'table.xlsx'

convert_html_to_excel(html_content, excel_file)

os.startfile(excel_file)
