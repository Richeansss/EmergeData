import openpyxl

def count_yes_no_in_shipped_column(file_path, sheet_name, start_marker, end_marker, shipped_column):
    # Загрузить рабочую книгу и выбрать указанный лист
    wb = openpyxl.load_workbook(file_path)
    sheet = wb[sheet_name]

    # Найти все диапазоны между маркерами
    ranges = []
    start_row = None
    first_row_skipped = False
    for row in sheet.iter_rows():
        if not first_row_skipped:
            first_row_skipped = True
            continue  # Пропустить первую строку

        for cell in row:
            if cell.value and start_marker in str(cell.value):
                start_row = cell.row
                start_value = cell.value  # Запомнить значение маркера
            if cell.value and end_marker in str(cell.value):
                end_row = cell.row
                if start_row is not None:
                    ranges.append((start_row, end_row, start_value))
                start_row = None  # Сбросить начальный маркер для поиска следующего диапазона
                break

    # Проверить, что найдены хотя бы один диапазон
    if not ranges:
        print("Не удалось найти диапазоны между маркерами.")
        return None

    # Найти индекс столбца "Отгружено (from file1)"
    header_row = sheet[1]  # Взять первую строку, где есть заголовки
    shipped_col_index = None
    for cell in header_row:
        if cell.value and shipped_column in str(cell.value):
            shipped_col_index = cell.column - 1
            break

    if shipped_col_index is None:
        print(f"Не удалось найти столбец с названием '{shipped_column}' в первой строке.")
        return None

    # Подсчитать количество "Да" и "Нет" (без учета регистра) в каждом диапазоне
    results = []
    total_yes_count = 0
    total_no_count = 0

    for start_row, end_row, start_value in ranges:
        yes_count = 0
        no_count = 0
        for row in sheet.iter_rows(min_row=start_row + 1, max_row=end_row - 1):
            cell_value = row[shipped_col_index].value
            if isinstance(cell_value, str):
                cell_value_lower = cell_value.lower()
                if cell_value_lower == "да":
                    yes_count += 1
                elif cell_value_lower == "нет":
                    no_count += 1
        results.append((start_value, start_row, end_row, yes_count, no_count))
        total_yes_count += yes_count
        total_no_count += no_count

    return results, total_yes_count, total_no_count

# Пример использования функции
file_path = 'C:\\Users\\Alexey\\IdeaProjects\\DataMerge\\SortedData_Output.xlsx'
sheet_name = 'GroupedData'  # замените на имя вашего листа
start_marker = 'Код позиции (from file1)'
end_marker = 'Кол-во'
shipped_column = 'Отгружено (from file1)'

results, total_yes_count, total_no_count = count_yes_no_in_shipped_column(file_path, sheet_name, start_marker, end_marker, shipped_column)

for start_value, start_row, end_row, yes_count, no_count in results:
    print(f"{start_marker}: {start_value} (Диапазон с {start_row} по {end_row}): 'Да': {yes_count}, 'Нет': {no_count}")

print(f"Общее количество 'Да': {total_yes_count}")
print(f"Общее количество 'Нет': {total_no_count}")
