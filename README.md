# Vitro Planner ETL

## Описание

Проект представляет собой ETL-процесс, который загружает и трансформирует данные из системы 1C:ЗиУП в виде Excel-файлов в базу данных MS SQL Server или PostgreSQL, для дальнейшего построения отчетов.

## Зависимости

Все зависимости указаны в файле `pyproject.toml`. Для установки зависимостей выполните следующую команду:

```
uv sync
```

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
    "connection_string": "your_connection_string_here",
    "mail_message": {
		"from_email": "",
        "to_emails": [""],
        "smtp_server": "",
        "smtp_port": ,
		"smtp_login": "",
		"smtp_password": ""
    }
}
```
2. Замените значения в file_paths, table_names, table_map, connection_string, mail_message на соответствующие значения.

3. Создайте в БД соответствующие таблицы со следующей структурой:

Для MS SQL Server:
```sql
CREATE TABLE [dbo].[your_first_table_name](
	[ID] [int] IDENTITY(1,1) NOT NULL,
	[Num] [int] NULL,
	[Date] [datetime] NULL,
	[Value] [decimal](10, 4) NULL
) ON [PRIMARY]
```
```sql
CREATE TABLE [dbo].[your_second_table_name](
	[ID] [int] IDENTITY(1,1) NOT NULL,
	[Num] [int] NULL,
	[Date] [datetime] NULL,
	[Value] [decimal](10, 4) NULL
) ON [PRIMARY]
```

Для PostgreSQL:
```sql
CREATE TABLE "public"."your_first_table_name" (
    id serial4 NOT NULL,
    num int4 NULL, 
    "date" date NULL, 
    value numeric(10, 4) NULL, 
    CONSTRAINT "your_first_table_name_pkey" PRIMARY KEY (id)
    );
```
```sql
CREATE TABLE "public"."your_second_table_name" (
    id serial4 NOT NULL,
    num int4 NULL, 
    "date" date NULL, 
    value numeric(10, 4) NULL, 
    CONSTRAINT "your_second_table_name_pkey" PRIMARY KEY (id)
    );
```

## Использование
### Вариант 1: Использование с установленным интерпретатором Python

Запустите `main.py` для выполнения ETL-процесса:

```
uv run main.py
```

### Вариант 2: Использование с компиляцией в .exe-файл
Если на вашем сервере не установлен интерпретатор Python, вы можете скомпилировать код в .exe-файл с помощью библиотеки pyinstaller. 
Установите pyinstaller с помощью команды:

```
pip install pyinstaller
```
Далее выполните следующую команду для компиляции кода в .exe-файл:

```
pyinstaller --onedir --add-data "config.json;." main.py
```
В результате компиляции в папке dist будет создана папка с названием вашего Python-файла. 
Внутри этой папки будет расположен исполняемый файл main.exe. Запустите этот файл для выполнения ETL-процесса:

```
./dist/main/main.exe
```

## Автоматический запуск
Можно настроить автоматический запуск скрипта через Windows Task Scheduler (для Windows) или Cron (для Linux).