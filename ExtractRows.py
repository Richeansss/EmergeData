import openpyxl

from SortColumn import move_ppp_column_to_left


def count_yes_no_in_shipped_column(file_path, sheet_name, start_marker, end_marker, shipped_column):
    try:
        # Загрузить рабочую книгу и выбрать указанный лист
        wb = openpyxl.load_workbook(file_path)
        sheet = wb[sheet_name]
    except Exception as e:
        print(f"Ошибка при загрузке файла или листа: {e}")
        return None

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
            shipped_col_index = cell.col_idx - 1  # col_idx дает индекс на основе 1, а не 0
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

def create_summary_rows(file_path, sheet_name, start_marker, end_marker, shipped_column):
    try:
        # Загрузить рабочую книгу и выбрать указанный лист
        wb = openpyxl.load_workbook(file_path)
        sheet_names = wb.sheetnames
        new_sheet_name = f"{sheet_name}_Summary"

        # Проверяем, есть ли уже лист с таким именем
        if new_sheet_name in sheet_names:
            wb.remove(wb[new_sheet_name])  # Удаляем существующий лист, если есть

        # Создаем новый лист
        wb.create_sheet(new_sheet_name)
        new_sheet = wb[new_sheet_name]

        sheet = wb[sheet_name]
    except Exception as e:
        print(f"Ошибка при загрузке файла или листа: {e}")
        return None

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
                start_row = cell.row + 1  # Следующая строка после маркера
                start_value = cell.value  # Запомнить значение маркера
            if cell.value and end_marker in str(cell.value):
                end_row = cell.row - 1  # Предыдущая строка перед маркером
                if start_row is not None:
                    ranges.append((start_row, end_row, start_value))
                start_row = None  # Сбросить начальный маркер для поиска следующего диапазона
                break

    # Проверить, что найдены хотя бы один диапазон
    if not ranges:
        print("Не удалось найти диапазоны между маркерами.")
        return None

    # Найти количество столбцов на листе
    num_columns = sheet.max_column

    # Подготовить заголовки для общих строк
    header_row = [cell.value for cell in sheet[1]]
    summary_rows = [header_row + ['Да', 'Нет']]  # Добавляем новые столбцы в заголовок

    # Обработать каждый диапазон
    for start_row, end_row, start_value in ranges:
        # Создать список для текущей общей строки
        summary_row = [None] * num_columns

        # Пройти по каждому столбцу и проверить значения в диапазоне
        for col_idx in range(num_columns):
            values_in_range = set()

            # Собрать все значения из текущего столбца в диапазоне
            for row_idx in range(start_row, end_row + 1):
                cell_value = sheet.cell(row=row_idx, column=col_idx + 1).value
                values_in_range.add(cell_value)

            # Если все значения в столбце одинаковые, записать это значение
            if len(values_in_range) == 1:
                summary_row[col_idx] = next(iter(values_in_range))
            else:
                summary_row[col_idx] = "Non"

        # Добавить количество "Да" и "Нет" в конец общей строки
        for start_val, row_start, row_end, yes_count, no_count in results:
            if start_value == start_val:
                summary_row.append(yes_count)
                summary_row.append(no_count)

        # Добавить общую строку в список
        summary_rows.append(summary_row)

    # Записать заголовок на новый лист
    for col_idx, header in enumerate(summary_rows[0]):
        new_sheet.cell(row=1, column=col_idx + 1, value=header)

    # Записать данные на новый лист
    for row_idx, row_data in enumerate(summary_rows[1:], start=2):
        for col_idx, value in enumerate(row_data):
            new_sheet.cell(row=row_idx, column=col_idx + 1, value=value)

    # Сохранить изменения в файл
    wb.save(file_path)

    return summary_rows


# Пример использования функций
file_path = 'C:\\DataMerge\\SortedData_Output.xlsx'
sheet_name = 'GroupedData'  # замените на имя вашего листа
start_marker = 'Код позиции (from file1)'
end_marker = 'Кол-во'
shipped_column = 'Отгружено (from file1)'

# Получаем результаты подсчета "Да" и "Нет"
results, total_yes_count, total_no_count = count_yes_no_in_shipped_column(file_path, sheet_name, start_marker, end_marker, shipped_column)

# Создаем общие строки с добавлением столбцов "Да" и "Нет" и записываем на новый лист
summary_rows = create_summary_rows(file_path, sheet_name, start_marker, end_marker, shipped_column)

#move_ppp_column_to_left(file_path)


if summary_rows is not None:
    print(f"Общие строки записаны на лист '{sheet_name}_Summary'.")
