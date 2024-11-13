import pyodbc
import pymysql
import pandas as pd
import configparser
from cv1_migration_sql_query import source_query_1, destination_table_1, source_query_2, destination_table_2,source_query_3, destination_table_3, source_query_4, destination_table_4, source_query_5, destination_table_5, source_query_6, destination_table_6

config = configparser.ConfigParser()
config.read('configloc.ini')

configpath = config.get('CONFIG','path')
config.read(configpath)

fmserver = config.get('multiDB','fmserver')
fmport = config.get('multiDB','fmport')
fmdatabase = config.get('multiDB','fmdatabase')
fmusername = config.get('multiDB','fmusername')
fmpassword = config.get('multiDB','fmpassword')

toserver = config.get('archDB','toserver')
toport = config.get('archDB','toport')
todatabase = config.get('archDB','todatabase')
tousername = config.get('archDB','tousername')
topassword = config.get('archDB','topassword')

import pyodbc

def copy_data(source_conn_str, destination_conn_str, source_query, destination_table):
    src_cursor = source_conn_str.cursor()
    dst_cursor = destination_conn_str.cursor()

    src_cursor.execute(source_query)
    column_count = len(src_cursor.description)
    placeholders = ", ".join(["?"] * column_count)
    dst_sql = f"INSERT INTO {destination_table} VALUES ({placeholders})"

    row = src_cursor.fetchone()
    while row:
        dst_cursor.execute(dst_sql, row)
        row = src_cursor.fetchone()

    src_cursor.commit()
    dst_cursor.commit()

    src_cursor.close()
    dst_cursor.close()

    print(f"Data from query inserted into {destination_table} successfully.")

def copy_paste_data():
    driver = '{ODBC Driver 17 for SQL Server}'
    fmconn = pyodbc.connect('DRIVER='+driver+';SERVER='+fmserver+';PORT='+fmport+';DATABASE='+fmdatabase+';UID='+fmusername+';PWD='+ fmpassword)
    toconn = pyodbc.connect('DRIVER='+driver+';SERVER='+toserver+';PORT='+toport+';DATABASE='+todatabase+';UID='+tousername+';PWD='+ topassword)

    copy_data(fmconn, toconn, source_query_1, destination_table_1)
    copy_data(fmconn, toconn, source_query_2, destination_table_2)
    copy_data(fmconn, toconn, source_query_3, destination_table_3)
    copy_data(fmconn, toconn, source_query_4, destination_table_4)
    copy_data(fmconn, toconn, source_query_5, destination_table_5)
    copy_data(fmconn, toconn, source_query_6, destination_table_6)



