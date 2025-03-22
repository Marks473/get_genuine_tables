import requests
from bs4 import BeautifulSoup

def is_relatable_table(table):
    """
    Проверка таблицы на реляционные данные:
    - Все строки должны иметь одинаковое количество колонок, учитывая атрибут colspan.
    - Таблица должна иметь как минимум две колонки.
    """
    rows = table.find_all('tr')
    if not rows:
        return False

    def count_cols(row):
        """Подсчет количества колонок с учётом colspan."""
        cells = row.find_all(['th', 'td'])
        total_cols = 0
        for cell in cells:
            try:
                total_cols += int(cell.get('colspan', 1))
            except ValueError:
                total_cols += 1
        return total_cols

    # Считаем число колонок в первой строке
    expected_cols = count_cols(rows[0])
    # Если колонок меньше двух, считаем таблицу нереляционной
    if expected_cols < 2:
        return False

    # Проверяем, что каждая строка имеет то же число колонок
    for row in rows:
        cols = count_cols(row)
        if cols != expected_cols:
            return False

    return True

if __name__ == "__main__":
    # HTML-код с ТРЕМЯ таблицами (две хорошие, одна плохая)
    html_content = """
    <html>
      <body>
        <!-- Хорошая многоуровневая таблица -->
        <table border="1">
          <tr>
            <th colspan="4" style="text-align:center;">Университет</th>
          </tr>
          <tr>
            <th colspan="2">1 год</th>
            <th colspan="2">2 год</th>
          </tr>
          <tr>
            <td>1 группа</td>
            <td>2 группа</td>
            <td>1 группа</td>
            <td>2 группа</td>
          </tr>
        </table>

        <!-- Плохая таблица -->
        <table border="1">
          <tr>
            <td>Только 1 колонка</td>
          </tr>
          <!-- Следующая строка отличается числом ячеек: 2, что делает её неоднородной -->
          <tr>
            <td>Неправильная</td>
            <td>Структура</td>
          </tr>
        </table>

        <!-- Ещё одна хорошая таблица -->
        <table border="1">
          <tr>
            <th>Колонка 1</th>
            <th>Колонка 2</th>
          </tr>
          <tr>
            <td>Значение 1</td>
            <td>Значение 2</td>
          </tr>
          <tr>
            <td>Значение 3</td>
            <td>Значение 4</td>
          </tr>
        </table>
      </body>
    </html>
    """

    # Парсим HTML-код
    soup = BeautifulSoup(html_content, 'html.parser')
    all_tables = soup.find_all('table')

    # Фильтрация «хороших» (реляционных) таблиц
    good_tables = [table for table in all_tables if is_relatable_table(table)]

    # Вывод краткой информации
    print(f"Всего таблиц: {len(all_tables)}")
    print(f"Хороших таблиц: {len(good_tables)}")

    # Сохранение «хороших» таблиц в массив (список)
    # (ниже пример хранения в виде HTML-строк, при желании можно хранить объекты BeautifulSoup)
    good_tables_html = [str(table) for table in good_tables]

    # Пример использования: просто выводим HTML каждой «хорошей» таблицы (укороченный вывод)
    for i, table_html in enumerate(good_tables_html, start=1):
        print(f"\n=== Хорошая таблица №{i} ===")
        print(table_html)
