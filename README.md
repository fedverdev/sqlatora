# Конвертатор данных из XLSX-файла в SQL-скрипт для PostgreSQL с поддержкой массивов.

### 1. Установка
Установите python 3 с оффициального сайта:
https://www.python.org/downloads/

Скачайте репозиторий используя команду:
```
git clone https://github.com/fedverdev/sqlatora.git
```

Установите зависимости с помощью команды:

```
pip install -r req.txt
```

### 2. Использование

Создайте ***config.json*** в той же директории где и скрипт ***sqlatora.py***. Укажите в поле ***columns*** название и тип полей в таблице, в поле ***input_file_name*** укажите имя xlsx файла, без расширения (.xlsx), в ***table_name*** имя таблицы и в ***output_sql*** имя выходного файла.

Запустите скрипт ***sqlatora.py***

Пример того, как может выглядеть ***config.json***:
```
{
    "columns": {
        "id": "SERIAL PRIMARY KEY",
        "name": "VARCHAR(100)",
        "age": "INTEGER",
        "salary": "DECIMAL(10,2)",
        "hire_date": "DATE",
        "is_active": "BOOLEAN",
        "description": "TEXT",
        "tags": TEXT[]
    },
    "input_file_name": "input",
    "table_name": "employee",
    "output_sql": "sql/output"
}
```
На выходе получиться файл output.py в папке sql

Список всех ключей:
| Ключ        | Описание   | Обязательность    |
| ----------- | ----------- | ------------------|
| columns     | поля и типы таблицы      | +             |
| table_name   | имя таблицы        | +              |
| input_file_name | имя входящего файла (без расширения) | +
| output_sql | имя выходящего sql файла (без расширения) | +
| sheet_name | имя листа из которого брать данные (по дефолту первый лист) | -
| create_table_script | будет ли в выходном файле скрипт для созданияе таблицы с указанными полями и типами данных (по дефолту true) | -
