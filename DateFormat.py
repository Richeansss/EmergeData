import openpyxl
from datetime import datetime

def convert_dates_in_place(file_path, sheet_name, columns):
    try:
        # Загрузить рабочую книгу и выбрать указанный лист
        wb = openpyxl.load_workbook(file_path)
        sheet = wb[sheet_name]
    except Exception as e:
        print(f"Ошибка при загрузке файла или листа: {e}")
        return

    # Найти индексы столбцов по их названиям
    header_row = sheet[1]
    column_indices = {}
    for col in columns:
        for cell in header_row:
            if cell.value and col in str(cell.value):
                column_indices[col] = cell.col_idx
                break

    # Пройти по всем строкам и указанным столбцам
    for col_name, col_idx in column_indices.items():
        for row in sheet.iter_rows(min_row=2, min_col=col_idx, max_col=col_idx):
            for cell in row:
                if cell.value:
                    try:
                        # Преобразовать дату в нужный формат, если она строка или datetime
                        if isinstance(cell.value, str):
                            try:
                                cell.value = datetime.strptime(cell.value, "%d.%m.%Y %H:%M:%S").strftime("%d.%m.%Y")
                            except ValueError:
                                try:
                                    cell.value = datetime.strptime(cell.value, "%d.%m.%Y").strftime("%d.%м.%Y")
                                except ValueError:
                                    continue
                        elif isinstance(cell.value, datetime):
                            cell.value = cell.value.strftime("%d.%м.%Y")

                        # Установить формат ячейки как текст, чтобы сохранить единство формата
                        cell.number_format = 'dd.mm.yyyy'
                    except ValueError:
                        continue

    # Сохранить изменения в исходный файл
    wb.save(file_path)
    print(f"Даты успешно преобразованы и сохранены в файл '{file_path}'")

# Пример использования функции
file_path = 'C:\\Users\\Alexey\\IdeaProjects\\DataMerge\\SortedData_Output.xlsx'
sheet_name = 'GroupedData'  # замените на имя вашего листа
columns = ['Дата отгрузки по графику (from file1)', 'Дата начала работ (from file2)', 'Дата окончания работ (from file2)']

convert_dates_in_place(file_path, sheet_name, columns)
