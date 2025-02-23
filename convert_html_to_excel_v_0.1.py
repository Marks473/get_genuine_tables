import os
import pandas as pd
from bs4 import BeautifulSoup

def is_relatable_table(table):
    """
    Проверка таблицы на реляционные данные:
    - наличие тегов <th> для заголовков,
    - наличие строк данных <tr> с ячейками <td>,
    - строки данных должны содержать ровно столько же <td>, сколько <th> в заголовке.
    """
    headers = table.find_all('th')
    data_rows = table.find_all('tr')[1:]  # Пропускаем первую строку (заголовок)

    if len(headers) == 0 or len(data_rows) == 0:
        return False

    header_count = len(headers)
    
    for row in data_rows:
        data_cells = row.find_all('td')
        if len(data_cells) != header_count:
            return False

    return True

def convert_html_to_excel(html_file, excel_file):
    try:
        # Чтение HTML файла
        with open(html_file, 'r', encoding='utf-8') as f:
            html_content = f.read()
        
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
html_file = 'table2.html'
excel_file = 'table.xlsx'
convert_html_to_excel(html_file, excel_file)
os.startfile(excel_file)
