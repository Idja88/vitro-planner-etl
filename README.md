# Vitro Planner ETL

## Описание

Проект представляет собой ETL-процесс на языке Python, который загружает и трансформирует данные из системы 1C в виде Excel-файлов в базу данных MS SQL Server, для дальнейшего построения отчетов для ERP системы Vitro Planner на платформе SSRS.

## Зависимости

Для запуска проекта убедитесь, что у вас установлены следующие библиотеки:

- pandas
- openpyxl
- pyodbc
- sqlalchemy
- cryptography

## Настройка

1. Создайте файл `config.json` в корневой папке проекта с следующей структурой:

```json
{
    "file_paths": [
        "path//to//your//first//excel//file.xlsx",
        "path//to//your//second//excel//file.xlsx"
    ],
    "table_names": [
        "your_first_table_name",
        "your_second_table_name"
    ],
    "table_map": {
        "ExcelValue1": 1,
        "ExcelValue2": 2
    },
    "secure_connection_string": "your_encrypted_connection_string_here"
}
```
2. Замените значения в file_paths, table_names, table_map и secure_connection_string на соответствующие значения.

3. Установите переменную окружения ENCRYPTION_KEY с ключом шифрования для строки подключения к базе данных.

4. Создайте в БД соответствующие таблицы со следующей структурой:

```sql
CREATE TABLE [dbo].[your_first_table_name](
	[ID] [int] IDENTITY(1,1) NOT NULL,
	[Num] [int] NULL,
	[Date] [datetime] NULL,
	[Value] [decimal](10, 4) NULL
) ON [PRIMARY]

CREATE TABLE [dbo].[your_second_table_name](
	[ID] [int] IDENTITY(1,1) NOT NULL,
	[Num] [int] NULL,
	[Date] [datetime] NULL,
	[Value] [decimal](10, 4) NULL
) ON [PRIMARY]
```
## Использование
### Вариант 1: Использование с установленным интерпретатором Python

Запустите `etl.py` для выполнения ETL-процесса:

```
python etl.py
```

### Вариант 2: Использование с компиляцией в .exe-файл
Если на вашем сервере не установлен интерпретатор Python, вы можете скомпилировать код в .exe-файл с помощью библиотеки pyinstaller. 
Установите pyinstaller с помощью команды:

```
pip install pyinstaller
```
Далее выполните следующую команду для компиляции кода в .exe-файл:

```
pyinstaller --onedir --add-data "config.json;." etl.py
```
В результате компиляции в папке dist будет создана папка с названием вашего Python-файла. 
Внутри этой папки будет расположен исполняемый файл etl.exe. Запустите этот файл для выполнения ETL-процесса:

```
./dist/etl/etl.exe
```

## Автоматический запуск
Можно настроить автоматический запуск скрипта через Windows Task Scheduler (для Windows) или Cron (для Linux).

### Windows Task Scheduler
1. Откройте "Планировщик заданий" (нажмите Win и введите "Планировщик заданий" в поисковой строке).
2. В правой части окна выберите "Создать задачу".
3. Задайте имя и описание для вашей задачи на вкладке "Общие".
4. Перейдите на вкладку "Триггеры" и нажмите "Новый", чтобы создать новый триггер. Установите параметры запуска по расписанию.
5. Перейдите на вкладку "Действия" и нажмите "Новый", чтобы создать новое действие. В поле "Программа или скрипт" укажите путь к интерпретатору Python (например, C:\Python39\python.exe) или к скомпилированному исполняемому файлу (например, C:\path\to\dist\etl\etl.exe). В поле "Аргументы" укажите имя Python-файла (только для интерпретатора Python, например, C:\path\to\etl.py).
6. Нажмите "ОК" для сохранения задачи. Скрипт будет автоматически запускаться по заданному расписанию.