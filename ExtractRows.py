import openpyxl
from datetime import datetime


# Функция для подсчета "Да" и "Нет" в столбце "Отгружено"
def count_yes_no_in_shipped_column(sheet, start_marker, end_marker, shipped_column):
    # Найти все диапазоны между маркерами
    ranges = find_ranges(sheet, start_marker, end_marker)
    if not ranges:
        print("Не удалось найти диапазоны между маркерами.")
        return None, 0, 0

    # Найти индекс столбца "Отгружено (from file1)"
    shipped_col_index = find_column_index(sheet, shipped_column)
    if shipped_col_index is None:
        print(f"Не удалось найти столбец с названием '{shipped_column}' в первой строке.")
        return None, 0, 0

    # Подсчитать количество "Да" и "Нет" (без учета регистра) в каждом диапазоне
    results = []
    total_yes_count = 0
    total_no_count = 0

    for start_row, end_row, start_value in ranges:
        yes_count, no_count = count_yes_no(sheet, start_row, end_row, shipped_col_index)
        results.append((start_value, start_row, end_row, yes_count, no_count))
        total_yes_count += yes_count
        total_no_count += no_count

    return results, total_yes_count, total_no_count


# Функция для поиска всех диапазонов между маркерами
def find_ranges(sheet, start_marker, end_marker):
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
    return ranges


# Функция для поиска индекса столбца по названию
def find_column_index(sheet, column_name):
    header_row = sheet[1]  # Взять первую строку, где есть заголовки
    for cell in header_row:
        if cell.value and column_name in str(cell.value):
            return cell.col_idx - 1  # col_idx дает индекс на основе 1, а не 0
    return None


# Функция для подсчета "Да" и "Нет" в указанном диапазоне строк
def count_yes_no(sheet, start_row, end_row, col_index):
    yes_count = 0
    no_count = 0
    for row in sheet.iter_rows(min_row=start_row + 1, max_row=end_row - 1):
        cell_value = row[col_index].value
        if isinstance(cell_value, str):
            cell_value_lower = cell_value.lower()
            if cell_value_lower == "да":
                yes_count += 1
            elif cell_value_lower == "нет":
                no_count += 1
    return yes_count, no_count


# Функция для создания и записи общих строк на новый лист
def create_summary_rows(file_path, sheet_name, start_marker, end_marker, shipped_column, schedule_date_column,
                        start_work_column, end_work_column):
    try:
        # Загрузить рабочую книгу и выбрать указанный лист
        wb = openpyxl.load_workbook(file_path)
        new_sheet_name = f"{sheet_name}_Summary"

        # Проверяем, есть ли уже лист с таким именем
        if new_sheet_name in wb.sheetnames:
            wb.remove(wb[new_sheet_name])  # Удаляем существующий лист, если есть

        # Создаем новый лист
        new_sheet = wb.create_sheet(new_sheet_name)
        sheet = wb[sheet_name]
    except Exception as e:
        print(f"Ошибка при загрузке файла или листа: {e}")
        return None

    # Найти все диапазоны между маркерами
    ranges = find_ranges(sheet, start_marker, end_marker)
    if not ranges:
        print("Не удалось найти диапазоны между маркерами.")
        return None

    # Найти индексы столбцов
    schedule_date_col_index = find_column_index(sheet, schedule_date_column)
    start_work_col_index = find_column_index(sheet, start_work_column)
    end_work_col_index = find_column_index(sheet, end_work_column)

    if None in [schedule_date_col_index, start_work_col_index, end_work_col_index]:
        print(
            f"Не удалось найти один или несколько столбцов с названиями '{schedule_date_column}', '{start_work_column}' или '{end_work_column}' в первой строке.")
        return None

    # Подготовить заголовки для общих строк
    header_row = [cell.value for cell in sheet[1]]
    summary_headers = header_row + ['Да', 'Нет', 'Последняя дата отгрузки', 'Последняя дата начала работ',
                                    'Последняя дата окончания работ']
    summary_rows = [summary_headers]

    # Обработать каждый диапазон
    for start_row, end_row, start_value in ranges:
        summary_row = prepare_summary_row(sheet, start_row, end_row, schedule_date_col_index, start_work_col_index,
                                          end_work_col_index)

        # Добавить количество "Да" и "Нет" в конец общей строки
        for start_val, row_start, row_end, yes_count, no_count in results:
            if start_value == start_val:
                summary_row.append(yes_count)
                summary_row.append(no_count)

        # Добавить общую строку в список
        summary_rows.append(summary_row)

    # Записать заголовок и данные на новый лист
    write_summary_to_sheet(new_sheet, summary_rows)

    # Сохранить изменения в файл
    wb.save(file_path)

    return summary_rows


# Функция для подготовки общей строки
def prepare_summary_row(sheet, start_row, end_row, schedule_date_col_index, start_work_col_index, end_work_col_index):
    num_columns = sheet.max_column
    summary_row = [None] * num_columns

    for col_idx in range(num_columns):
        values_in_range = set()
        for row_idx in range(start_row, end_row + 1):
            cell_value = sheet.cell(row=row_idx, column=col_idx + 1).value
            values_in_range.add(cell_value)
        if len(values_in_range) == 1:
            summary_row[col_idx] = next(iter(values_in_range))
        else:
            summary_row[col_idx] = "Non"

    summary_row.append(find_latest_date(sheet, start_row, end_row, schedule_date_col_index))
    summary_row.append(find_latest_date(sheet, start_row, end_row, start_work_col_index))
    summary_row.append(find_latest_date(sheet, start_row, end_row, end_work_col_index))
    return summary_row


# Функция для поиска последней даты в диапазоне строк
def find_latest_date(sheet, start_row, end_row, col_index):
    latest_date = None
    for row_idx in range(start_row, end_row + 1):
        cell_value = sheet.cell(row=row_idx, column=col_index + 1).value
        if isinstance(cell_value, str):
            try:
                cell_value = datetime.strptime(cell_value, "%d.%m.%Y")
            except ValueError:
                cell_value = None
        if latest_date is None or (cell_value and cell_value > latest_date):
            latest_date = cell_value
    return latest_date


# Функция для записи общих строк на новый лист
def write_summary_to_sheet(sheet, summary_rows):
    for row_idx, row_data in enumerate(summary_rows, start=1):
        for col_idx, value in enumerate(row_data):
            sheet.cell(row=row_idx, column=col_idx + 1, value=value)


# Пример использования функций
file_path = 'C:\\Users\\Alexey\\IdeaProjects\\DataMerge\\SortedData_Output.xlsx'
sheet_name = 'GroupedData'  # замените на имя вашего листа
start_marker = 'Код позиции (from file1)'
end_marker = 'Кол-во'
shipped_column = 'Отгружено (from file1)'
schedule_date_column = 'Дата отгрузки (from file1)'
start_work_column = 'Дата начала работ (from file1)'
end_work_column = 'Дата окончания работ (from file1)'

# Получаем результаты подсчета "Да" и "Нет"
results, total_yes_count, total_no_count = count_yes_no_in_shipped_column(file_path, sheet_name, start_marker,
                                                                          end_marker, shipped_column)

# Создаем общие строки с добавлением столбцов "Да", "Нет" и "Последняя дата отгрузки" и записываем на новый лист
summary_rows = create_summary_rows(file_path, sheet_name, start_marker, end_marker, shipped_column,
                                   schedule_date_column, start_work_column, end_work_column)

if summary_rows is not None:
    print(f"Общие строки записаны на лист '{sheet_name}_Summary'.")
else:
    print("Не удалось создать общие строки.")
