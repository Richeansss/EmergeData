import openpyxl
import os
from datetime import datetime

def load_workbook_and_sheet(file_path, sheet_name):
    try:
        wb = openpyxl.load_workbook(file_path)
        sheet = wb[sheet_name]
        return wb, sheet
    except Exception as e:
        print(f"Ошибка при загрузке файла или листа: {e}")
        return None, None

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

def find_column_index(header_row, column_name):
    for cell in header_row:
        if cell.value and column_name in str(cell.value):
            return cell.col_idx - 1  # col_idx дает индекс на основе 1, а не 0
    return None

def count_yes_no(sheet, ranges, shipped_col_index):
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

def count_yes_no_in_shipped_column(file_path, sheet_name, start_marker, end_marker, shipped_column):
    wb, sheet = load_workbook_and_sheet(file_path, sheet_name)
    if not wb or not sheet:
        return None

    ranges = find_ranges(sheet, start_marker, end_marker)
    if not ranges:
        print("Не удалось найти диапазоны между маркерами.")
        return None

    header_row = sheet[1]  # Взять первую строку, где есть заголовки
    shipped_col_index = find_column_index(header_row, shipped_column)
    if shipped_col_index is None:
        print(f"Не удалось найти столбец с названием '{shipped_column}' в первой строке.")
        return None

    return count_yes_no(sheet, ranges, shipped_col_index)

def find_last_date_in_range(sheet, start_row, end_row, col_index, date_format="%d.%m.%Y"):
    latest_date = None
    for row_idx in range(start_row, end_row + 1):
        cell_value = sheet.cell(row=row_idx, column=col_index + 1).value
        if isinstance(cell_value, str):
            try:
                cell_value = datetime.strptime(cell_value, date_format)
            except ValueError:
                cell_value = None
        if latest_date is None or (cell_value and cell_value > latest_date):
            latest_date = cell_value
    return latest_date

def create_summary_rows(file_path, sheet_name, start_marker, end_marker, shipped_column, schedule_date_column, start_work_column, end_work_column):
    wb, sheet = load_workbook_and_sheet(file_path, sheet_name)
    if not wb or not sheet:
        return None

    new_sheet_name = f"{sheet_name}_Summary"
    if new_sheet_name in wb.sheetnames:
        wb.remove(wb[new_sheet_name])  # Удаляем существующий лист, если есть

    new_sheet = wb.create_sheet(new_sheet_name)
    ranges = find_ranges(sheet, start_marker, end_marker)
    if not ranges:
        print("Не удалось найти диапазоны между маркерами.")
        return None

    num_columns = sheet.max_column
    header_row = sheet[1]
    schedule_date_col_index = find_column_index(header_row, schedule_date_column)
    start_work_col_index = find_column_index(header_row, start_work_column)
    end_work_col_index = find_column_index(header_row, end_work_column)

    if None in [schedule_date_col_index, start_work_col_index, end_work_col_index]:
        print(f"Не удалось найти один из столбцов: '{schedule_date_column}', '{start_work_column}', '{end_work_column}'.")
        return None

    summary_header = [cell.value for cell in header_row] + ['Да', 'Нет', 'Последняя дата отгрузки', 'Последняя дата начала работ', 'Последняя дата окончания работ']
    summary_rows = [summary_header]

    results, total_yes_count, total_no_count = count_yes_no(sheet, ranges, find_column_index(header_row, shipped_column))

    for start_row, end_row, start_value in ranges:
        summary_row = [None] * num_columns
        for col_idx in range(num_columns):
            values_in_range = {sheet.cell(row=row_idx, column=col_idx + 1).value for row_idx in range(start_row, end_row + 1)}
            summary_row[col_idx] = next(iter(values_in_range)) if len(values_in_range) == 1 else "Non"

        summary_row.append(find_last_date_in_range(sheet, start_row, end_row, schedule_date_col_index))
        summary_row.append(find_last_date_in_range(sheet, start_row, end_row, start_work_col_index))
        summary_row.append(find_last_date_in_range(sheet, start_row, end_row, end_work_col_index))

        for start_val, row_start, row_end, yes_count, no_count in results:
            if start_value == start_val:
                summary_row.append(yes_count)
                summary_row.append(no_count)

        summary_rows.append(summary_row)

    for row_idx, row_data in enumerate(summary_rows, start=1):
        for col_idx, value in enumerate(row_data):
            new_sheet.cell(row=row_idx, column=col_idx + 1, value=value)

    wb.save(file_path)
    return summary_rows

# Пример использования функций
file_path = 'C:\\Users\\Alexey\\IdeaProjects\\DataMerge\\SortedData_Output.xlsx'
sheet_name = 'GroupedData'
start_marker = 'Код позиции (from file1)'
end_marker = 'Кол-во'
shipped_column = 'Отгружено (from file1)'
schedule_date_column = 'Дата отгрузки по графику (from file1)'
start_work_column = 'Дата начала работ (from file2)'
end_work_column = 'Дата окончания работ (from file2)'

results, total_yes_count, total_no_count = count_yes_no_in_shipped_column(file_path, sheet_name, start_marker, end_marker, shipped_column)
summary_rows = create_summary_rows(file_path, sheet_name, start_marker, end_marker, shipped_column, schedule_date_column, start_work_column, end_work_column)

if summary_rows is not None:
    print(f"Общие строки записаны на лист '{sheet_name}_Summary'.")
