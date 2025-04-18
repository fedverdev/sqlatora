import openpyxl
from datetime import datetime
import json
import re
from typing import Any, Dict, List, Union
from pathlib import Path
import os

COLUMNS = "columns"
INPUT_FILE_NAME = "input_file_name"
SQL_FILE = "output_sql"
SHEET_NAME = "sheet_name"
CREATE_TABLE_SCRIPT = "create_table_script"
TABLE_NAME = "table_name"

errors_map = {}

def safe_open(file_path, mode="r", encoding=None, create_dirs=True):
    """
    Безопасное открытие файла с созданием директорий при необходимости
    
    :param file_path: путь к файлу
    :param mode: режим открытия (по умолчанию 'r')
    :param encoding: кодировка файла
    :param create_dirs: создавать ли директории
    :return: файловый объект
    """
    path = Path(file_path)
    
    if create_dirs and mode not in ("r", "rb"):
        path.parent.mkdir(parents=True, exist_ok=True)
    
    if "b" in mode:
        encoding = None
    
    return open(file_path, mode=mode, encoding=encoding)

def xlsx_to_postgresql_sql(xlsx_file_path, table_name, sql_file_path, column_types, sheet_name=None, create_table_script=True):
    """
    Конвертирует данные из XLSX-файла в SQL-скрипт для PostgreSQL с поддержкой массивов
    
    :param xlsx_file_path: Путь к XLSX-файлу
    :param table_name: Имя таблицы в PostgreSQL
    :param sql_file_path: Путь для сохранения SQL-файла
    :param column_types: Словарь с типами столбцов {'column_name': 'postgres_type'}
    :param sheet_name: Название листа (если None, берется первый лист)
    """
    # Загружаем книгу Excel
    workbook = openpyxl.load_workbook(xlsx_file_path)
    
    # Выбираем лист
    sheet = workbook[sheet_name] if sheet_name else workbook.active
    
    # Открываем файл для записи SQL-скрипта
    with safe_open(sql_file_path, 'w+', encoding='utf-8') as sql_file:
        # Получаем заголовки столбцов
        headers = [str(cell.value).strip() for cell in sheet[1]]
        
        # Проверяем, что все заголовки есть в column_types
        missing_columns = [col for col in headers if col not in column_types]
        if missing_columns:
            raise ValueError(f"Типы не указаны для столбцов: {', '.join(missing_columns)}")
        
        # Генерируем CREATE TABLE
        if create_table_script:
            sql_file.write(f"CREATE TABLE {table_name} (\n")
            columns_def = []
            for header in headers:
                columns_def.append(f"    {header} {column_types[header]}")
            sql_file.write(",\n".join(columns_def))
            sql_file.write("\n);\n\n")
        
        # Генерируем INSERT-запросы
        for row in sheet.iter_rows(min_row=2):
            values = []
            for idx, cell in enumerate(row):
                header = headers[idx]
                pg_type = column_types[header].upper()
                cell_value = cell.value
                
                if cell_value is None:
                    values.append("NULL")
                elif '[]' in pg_type:  # Обработка массивов
                    if isinstance(cell_value, str):
                        # Пытаемся разобрать как JSON массив
                        try:
                            array_data = json.loads(cell_value)
                            if not isinstance(array_data, list):
                                array_data = [array_data]
                        except json.JSONDecodeError:
                            # Разделяем строку по запятым или другим разделителям
                            array_data = re.split(r'[,;]\s*', cell_value.strip())
                        
                        # Экранируем элементы массива
                        escaped_array = []
                        for item in array_data:
                            if item is None:
                                escaped_array.append("NULL")
                            else:
                                escaped_item = str(item).replace("'", "''")
                                escaped_array.append(f"'{escaped_item}'")
                        
                        values.append(f"ARRAY[{', '.join(escaped_array)}]")
                    elif isinstance(cell_value, (list, tuple)):
                        escaped_array = []
                        for item in cell_value:
                            if item is None:
                                escaped_array.append("NULL")
                            else:
                                escaped_item = str(item).replace("'", "''")
                                escaped_array.append(f"'{escaped_item}'")
                        values.append(f"ARRAY[{', '.join(escaped_array)}]")
                    else:
                        values.append(f"ARRAY['{str(cell_value).replace("'", "''")}']")
                elif isinstance(cell_value, (int, float)):
                    if 'INT' in pg_type or 'NUMERIC' in pg_type or 'DECIMAL' in pg_type:
                        values.append(str(cell_value))
                    else:
                        values.append(str(cell_value))
                elif isinstance(cell_value, datetime):
                    if 'DATE' in pg_type:
                        values.append(f"'{cell_value.strftime('%Y-%m-%d')}'")
                    elif 'TIME' in pg_type:
                        values.append(f"'{cell_value.strftime('%H:%M:%S')}'")
                    else:
                        values.append(f"'{cell_value.strftime('%Y-%m-%d %H:%M:%S')}'")
                elif isinstance(cell_value, str):
                    if pg_type == 'JSON' or pg_type == 'JSONB':
                        try:
                            # Проверяем, валидный ли это JSON
                            json.loads(cell_value)
                            values.append(f"'{cell_value}'::{pg_type}")
                        except json.JSONDecodeError:
                            # Если не JSON, заключаем в кавычки как обычную строку
                            escaped_value = cell_value.replace("'", "''")
                            values.append(f"'{escaped_value}'")
                    else:
                        escaped_value = cell_value.replace("'", "''")
                        values.append(f"'{escaped_value}'")
                elif isinstance(cell_value, bool):
                    values.append("TRUE" if cell_value else "FALSE")
                else:
                    values.append(f"'{str(cell_value).replace("'", "''")}'")
            
            # Формируем SQL-запрос
            columns = ', '.join(headers)
            values_str = ', '.join(values)
            insert_sql = f"INSERT INTO {table_name} ({columns}) VALUES ({values_str});\n"
            
            # Записываем в файл
            sql_file.write(insert_sql)
    print(f"SQL-скрипт успешно создан: {sql_file_path}")

def json_to_map(json_data: Union[str, bytes, bytearray, dict]) -> Dict[str, Any]:
    """
    Преобразует JSON в словарь (map) Python.
    
    Аргументы:
        json_data: JSON-строка, bytes, bytearray или уже готовый dict
    
    Возвращает:
        Словарь Python с данными из JSON
    
    Исключения:
        json.JSONDecodeError: если переданная строка не является валидным JSON
    """
    if isinstance(json_data, (str, bytes, bytearray)):
        return json.loads(json_data)
    elif isinstance(json_data, dict):
        return json_data
    else:
        raise TypeError("Неподдерживаемый тип данных. Ожидается JSON-строка, bytes или dict")

def json_file_to_map(file_path: str, encoding: str = 'utf-8') -> Dict[str, Any]:
    """
    Читает JSON из файла и преобразует в словарь Python.
    
    Аргументы:
        file_path: путь к JSON-файлу
        encoding: кодировка файла (по умолчанию utf-8)
    
    Возвращает:
        Словарь Python с данными из JSON-файла
    """
    with open(file_path, 'r', encoding=encoding) as f:
        return json.load(f)
    
def deep_json_to_map(data: Any) -> Any:
    """
    Рекурсивно преобразует все JSON-строки в структуре данных в словари.
    
    Аргументы:
        data: данные для обработки (может быть строкой, списком, словарем)
    
    Возвращает:
        Обработанные данные с преобразованными JSON-строками
    """
    if isinstance(data, str):
        try:
            return json_to_map(data)
        except json.JSONDecodeError:
            return data
    elif isinstance(data, dict):
        return {k: deep_json_to_map(v) for k, v in data.items()}
    elif isinstance(data, (list, tuple)):
        return [deep_json_to_map(item) for item in data]
    else:
        return data

# Пример использования
if __name__ == "__main__":
    
    # Типы столбцов для PostgreSQL (включая массивы)
    config = deep_json_to_map(json_file_to_map("config.json"))

    missed_keywords_list = []

    if COLUMNS in config:
        column_types = config[COLUMNS]
    else:
        missed_keywords_list.append(COLUMNS)
    
    if INPUT_FILE_NAME in config:
        xlsx_file = config[INPUT_FILE_NAME] + ".xlsx"
    else:
        missed_keywords_list.append(INPUT_FILE_NAME)
    
    if TABLE_NAME in config:
        table_name = config[TABLE_NAME]
    else:
        missed_keywords_list.append(TABLE_NAME)
    
    if SQL_FILE in config:
        sql_file = config[SQL_FILE] + ".sql"
    else:
        missed_keywords_list.append(SQL_FILE)

    if SHEET_NAME in config:
        sheet_name = config[SHEET_NAME]
    else:
        sheet_name = None

    if CREATE_TABLE_SCRIPT in config:
        create_table_script = config[CREATE_TABLE_SCRIPT]
    else:
        create_table_script = True
    
    if len(missed_keywords_list) > 0:
        errors_map["в конфиге нет необходимых значений"] = str(missed_keywords_list)
    
    if len(errors_map) > 0:
        print("SQL-скрипт не был создан из-за следующих ошибок:")
        for err_name, err in errors_map.items():
            print(f"- {err_name}: {err}")
    else:
        xlsx_to_postgresql_sql(xlsx_file, table_name, sql_file, column_types, sheet_name, create_table_script)
    
    

    
    