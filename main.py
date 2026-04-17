import os
import sys
import json
from urllib.parse import quote
import pyodbc
import psycopg2
from datetime import date
import pandas as pd
import sqlalchemy as sa
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

#Functions
def delete_from_table(cursor, table):
    date = date.today()
    query = f"DELETE FROM {table} WHERE Date >= CAST('{date}' AS DATE)"
    cursor.execute(query)
    cursor.commit()

def check_if_num_exists(cursor, Num, Date, Table):
    query = f"SELECT Num FROM {Table} WHERE Num = {Num} and Date = CAST('{Date}' AS DATE)"
    cursor.execute(query)
    return cursor.fetchone() is not None

def update_row(cursor, Num, Date, Value, Table):
    query = f"UPDATE {Table} SET Value = {Value} WHERE Num = {Num} and Date = CAST('{Date}' AS DATE)"
    cursor.execute(query)

def update_db(cursor, df, table):
    tmp_df = pd.DataFrame(columns=['Num', 'Date', 'Value'])
    for i, row in df.iterrows():
        if check_if_num_exists(cursor, row['Num'], row['Date'], table):
            update_row(cursor, row['Num'], row['Date'], row['Value'], table)
        else:
            tmp_df = pd.concat([tmp_df, row.to_frame().T], axis=0, ignore_index=True)      
    return tmp_df

def insert_into_table(cursor, Num, Date, Value, Table):
    query = f"INSERT INTO {Table} (Num, Date, Value) VALUES('{Num}', '{Date}', '{Value}')"
    cursor.execute(query)

def append_from_df_to_db(cursor, df, Table):
    for i, row in df.iterrows():
        insert_into_table(cursor, row['Num'], row['Date'], row['Value'], Table)

def resolve_path(path):
    if getattr(sys, "frozen", False):
        resolved_path = os.path.abspath(os.path.join(sys._MEIPASS, path))
    else:
        resolved_path = os.path.abspath(os.path.join(os.getcwd(), path))

    return resolved_path

def connect_to_db(connection_string: str):
    parts = {}
    for param in connection_string.split(';'):
        if '=' in param:
            key, value = param.split('=', 1)
            parts[key.strip()] = value.strip().strip('{}')
    
    driver = parts.get('Driver') or parts.get('DRIVER')
    host = parts.get('Server') or parts.get('SERVER')
    port = parts.get('Port') or parts.get('PORT')
    database = parts.get('Database') or parts.get('DATABASE')
    user = parts.get('Uid') or parts.get('UID')
    password = parts.get('Pwd') or parts.get('PWD')
    
    password_encoded = quote(password, safe='')
    
    # Определяем тип БД на основе driver в строке подключения
    if "PostgreSQL" in driver:
        connection_uri = f"postgresql+psycopg2://{user}:{password_encoded}@{host}:{port if port else 5432}/{database}"
        engine = sa.create_engine(connection_uri, echo=True)
        
    elif "SQL Server" in driver:
        connection_uri = f"mssql+pyodbc://{user}:{password_encoded}@{host}:{port if port else 1433}/{database}?driver={quote(driver)}"
        engine = sa.create_engine(connection_uri, echo=True, fast_executemany=True)
    
    else:
        raise ValueError(f"Unknown database driver: {driver}")
    
    connection = engine.raw_connection()
    return connection

def send_email(subject, message, from_email, to_emails, smtp_server, smtp_port, smtp_login, smtp_password):
    try:
        server = smtplib.SMTP(smtp_server, smtp_port)
        server.starttls()
        server.login(smtp_login, smtp_password)
        for to_email in to_emails:
            msg = MIMEMultipart()
            msg['From'] = from_email
            msg['To'] = to_email
            msg['Subject'] = subject
            msg.attach(MIMEText(message, 'plain'))
            server.send_message(msg)
    except Exception as e:
        print("Error sending email:", str(e))
    finally:
        server.quit()

#MAIN
def main(cursor, file_paths, database_names, table_map):
    for (file_path, database_name) in zip(file_paths, database_names):

        delete_from_table(cursor, database_name)

        sheet_to_df = pd.read_excel(file_path, sheet_name=None)

        for sheet in sheet_to_df.keys():
            df = pd.read_excel(file_path, sheet_name=sheet)
            #определяем название первого столбца и переименовываем в удобный
            first_column_name = df.iloc[:,0].name
            df.rename(columns={first_column_name:'Num'}, inplace=True)
            #убираем лишние пробелы в значениях столбцов
            df.rename(columns=lambda x: x.replace(' ', ''), inplace=True)
            #Транспонируем значения датафрейма
            df = df.melt(id_vars=["Num"],var_name="Date",value_name="Value")
            #Мэппинг значений
            df["Value"] = df["Value"].map(table_map).fillna(0, downcast='infer') 
            df_to_app = update_db(cursor, df, database_name)
            append_from_df_to_db(cursor, df_to_app, database_name)

        cursor.commit()

if __name__ == "__main__":
    config_path = os.path.join(os.getcwd(), 'config.json')
    #считываем настройки
    with open(config_path, "r", encoding="utf-8") as config_file:
        config = json.load(config_file)      
        file_paths = config['file_paths']
        table_names = config['table_names']
        table_map = config['table_map']
        connection_string = config['connection_string']
        from_email = config['mail_message']['from_email']
        to_emails = config['mail_message']['to_emails']
        smtp_server = config['mail_message']['smtp_server']
        smtp_port = config['mail_message']['smtp_port']
        smtp_login = config['mail_message']['smtp_login']
        smtp_password = config['mail_message']['smtp_password']

    connection = connect_to_db(connection_string)
    cursor = connection.cursor()

    try:
        main(cursor, file_paths, table_names, table_map)
    except Exception as e:
        error_message = f"An error occurred: {str(e)}"
        send_email("Error in ETL process", error_message, from_email, to_emails, smtp_server, smtp_port)
        raise SystemExit(1)
    finally:
        connection.close()