import pyodbc
import pymysql
import pandas as pd
import configparser
import math
from sqlalchemy import create_engine
import numpy as np
import urllib.request, urllib.error, urllib.parse
import urllib
from datetime import datetime

config = configparser.ConfigParser()
config.read('configloc.ini')

configpath = config.get('CONFIG','path')
config.read(configpath)

fmserver = config.get('sosDB','fmserver')
fmport = config.get('sosDB','fmport')
fmdatabase = config.get('sosDB','fmdatabase')
fmusername = config.get('sosDB','fmusername')
fmpassword = config.get('sosDB','fmpassword')
fmschema = config.get('sosDB','fmschema')

toserver = config.get('archDB','toserver')
toport = config.get('archDB','toport')
todatabase = config.get('archDB','todatabase')
tousername = config.get('archDB','tousername')
topassword = config.get('archDB','topassword')
toschema = config.get('archDB','toschema')

# toserver = config.get('hkgDB','toserver')
# toport = config.get('hkgDB','toport')
# todatabase = config.get('hkgDB','todatabase')
# tousername = config.get('hkgDB','tousername')
# topassword = config.get('hkgDB','topassword')
# toschema = config.get('hkgDB','toschema')


def get_Tbl():
    # Example hardcoded data, you can query this from a SQL Server table
    tables = [
         {'s_table': 'tbl_account_closure_request', 't_table': 'Dwh_Sos_tbl_account_closure_request'}
        ,{'s_table': 'tbl_assignment_log', 't_table': 'Dwh_Sos_tbl_assignment_log'}
        ,{'s_table': 'tbl_coca_log', 't_table': 'Dwh_Sos_tbl_coca_log'}
        ,{'s_table': 'tbl_crrlog', 't_table': 'Dwh_Sos_tbl_crrlog'}
        ,{'s_table': 'tbl_cs', 't_table': 'Dwh_Sos_tbl_cs'}
        ,{'s_table': 'tbl_dpmscheckinglog', 't_table': 'Dwh_Sos_tbl_dpmscheckinglog'}
        ,{'s_table': 'tbl_mpfa_checking', 't_table': 'Dwh_Sos_tbl_mpfa_checking'}
        ,{'s_table': 'tbl_coca', 't_table': 'Dwh_Sos_tbl_coca'}
        ,{'s_table': 'tbl_transfer_externallog', 't_table': 'Dwh_Sos_tbl_transfer_externallog'}
        ,{'s_table': 'tbl_transfer_log', 't_table': 'Dwh_Sos_tbl_transfer_log'}
        ,{'s_table': 'tbl_pdudailylog', 't_table': 'Dwh_Sos_tbl_pdudailylog'}
    ]
    return tables

def convert_float_to_int(value):
    """Convert float values to int if applicable, otherwise return the value unchanged."""
    try:
        if isinstance(value, float) and math.isnan(value):
            print("Hello " + value)
            # print("Encountered NaN, returning None or default value")
            return None  
        elif isinstance(value, float):
            return int(value)
        return value  # Return the original value if it's not a float
    except (ValueError, TypeError) as e:
        print(f"Error converting value: {e}")
        return None 

# Main process
def copy_paste_data():
    table_list = get_Tbl()

    driver = '{ODBC Driver 17 for SQL Server}'
    fmconn = f"mysql+pymysql://{fmusername}:{fmpassword}@{fmserver}/{fmdatabase}"
    engine = create_engine(fmconn)

    # fmconn = pymysql.connect(host=fmserver, port=3306, database=fmdatabase, user=fmusername, password=fmpassword)
    toconn = pyodbc.connect('DRIVER='+driver+';SERVER='+toserver+';PORT='+toport+';DATABASE='+todatabase+';UID='+tousername+';PWD='+ topassword)
    tocsr = toconn.cursor()


    for table in table_list:
        s_table = table['s_table']
        t_table = table['t_table']
        # print("Hello " + s_table + " + " + t_table)

        # Fetch data from source
        query = f"SELECT * FROM {fmschema}.{s_table}"   
        df = pd.read_sql(query, engine)
        df = df.replace({np.nan: None})

        # df.to_csv(f'{s_table}.xlsx', index=False)

        # print(df)
        # #Truncate from destination table data
        sql= "Truncate table " + toschema + "." + t_table
        tocsr.execute(sql)
        toconn.commit()

        for index, row in df.iterrows():
            # print("hello sk 1" + str(index) + " + " + str(row))
            placeholders = ', '.join(['?'] * len(row))
            columns = ', '.join(df.columns)
            cleaned_row = [convert_float_to_int(col) for col in row]
            # print(f"Cleaned Row: {cleaned_row}")
            insert_query = f"INSERT INTO {toschema}.{t_table} ({columns}) VALUES ({placeholders})"
            tocsr.execute(insert_query, tuple(cleaned_row))
        
        toconn.commit()

        now = datetime.now()
        current_time = now.strftime("%H:%M:%S")
        print("Migrated data from " + s_table + " to " + t_table + " : " + current_time)

    # fmconn.close()
    tocsr.close()
    toconn.close()

# Execute the function
# copy_paste_data()
