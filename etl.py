import os
import sys
import json
import urllib
import pyodbc
import pandas as pd
import sqlalchemy as sa
from cryptography.fernet import Fernet

#Functions
def check_if_num_exists(cursor, Num, Date, Table):
    query = f"SELECT Num FROM {Table} WHERE Num={Num} and Date=CONVERT(date,'{Date}')"
    cursor.execute(query)
    return cursor.fetchone() is not None

def update_row(cursor, Num, Date, Value, Table):
    query = f"UPDATE {Table} SET Value={Value} WHERE Num={Num} and Date=CONVERT(date,'{Date}')"
    cursor.execute(query)

def update_db(cursor, df, table):
    tmp_df = pd.DataFrame(columns=['Num', 'Date', 'Value'])

    for i, row in df.iterrows():
        if check_if_num_exists(cursor, row['Num'], row['Date'], table):
            update_row(cursor, row['Num'], row['Date'], row['Value'], table)
        else:
            tmp_df = tmp_df.append(row)
            
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

def connect_to_db(secure_connection_string):
    key = os.environ.get("ENCRYPTION_KEY").encode()
    cipher_suite = Fernet(key)
    connection_string = cipher_suite.decrypt(secure_connection_string)
    connection_uri = f"mssql+pyodbc:///?odbc_connect={urllib.parse.quote_plus(connection_string)}"
    engine = sa.create_engine(connection_uri, fast_executemany=True, echo=True)
    connection = engine.raw_connection()
    return connection

#MAIN
def main(cursor, file_paths, database_names, table_map):
    for (file_path, database_name) in zip(file_paths, database_names):

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
        secure_connection_string = config['secure_connection_string']

    connection = connect_to_db(secure_connection_string)
    cursor = connection.cursor()

    try:
        main(cursor, file_paths, table_names, table_map)
    finally:
        connection.close()