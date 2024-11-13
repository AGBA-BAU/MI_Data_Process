import os
import configparser
import win32com.client
import pyodbc
import pymsgbox

#read config file
config = configparser.ConfigParser() 
config.read('config.ini')

server = config.get('DB','SERVER')
database = config.get('DB','DATABASE')
username = config.get('DB','USERNAME')
password = config.get('DB','PASSWORD')

driver= '{ODBC Driver 17 for SQL Server}'
cnxn = pyodbc.connect('DRIVER='+driver+';PORT=1433;SERVER='+server+';PORT=1443;DATABASE='+database+';UID='+username+';PWD='+ password+';Charset=UTF8')
cursor = cnxn.cursor()

# Specify the folder path
data_folder_path = str(os.path.dirname(__file__))+"\MI_Serv_Mgmt_Overview\Data"
template_folder_path = str(os.path.dirname(__file__))+"\MI_Serv_Mgmt_Overview\Template\MI Serv Mgmt Overview-template.xlsx"

yr = config.get('Reporting','YEAR')

# List files in the folder
files = os.listdir(data_folder_path)

# Check if there is exactly one file in the folder
if len(files) == 1:
    file_name = files[0]  # Get the file name
    print(f"The file in the folder is: {file_name}")
else:
    print("There are either no files or more than one file in the folder.")
    exit()


excel = win32com.client.Dispatch("Excel.Application")       # Create an instance of Excel
excel.Visible = True            # Make Excel visible (optional)
excel.DisplayAlerts = False     #To surpress warning dialog

Data_workbook = excel.Workbooks.Open(data_folder_path + "\\" + file_name)           # Open a data workbook
Template_workbook = excel.Workbooks.Open(template_folder_path)           # Open a template workbook

mth_row = 12
sql_queries_lists = []

for sheet_index in range(1, Data_workbook.Worksheets.Count + 1):
    
    d_worksheet = Data_workbook.Worksheets(sheet_index)
    sheet_name = d_worksheet.Name
    t_worksheet = Template_workbook.Worksheets(sheet_name)
    print(f"Sheet {sheet_index}: {sheet_name}")

    d_cnt_formula = f'=COUNTIF(C:C,"?*")'
    d_cnt = d_worksheet.Evaluate(d_cnt_formula)
    t_cnt_formula = f'=COUNTIF(C:C,"?*")'
    t_cnt = d_worksheet.Evaluate(t_cnt_formula)

    
    if d_cnt != t_cnt:
        pymsgbox.alert(f"The number of attributes in the sheet named '{sheet_name}' differs between the data file and the template file. Kindly verify this difference.", 'Alert Box')
        exit()

    #Data
    d_match_formula = f'=MATCH("Jan",{mth_row}:{mth_row},0)'
    d_start_col = d_worksheet.Evaluate(d_match_formula)
    d_last_col = d_start_col + 14
    d_last_row = d_worksheet.Cells(d_worksheet.Rows.Count, 3).End(-4162).Row 

    d_rng_formula = f'=ADDRESS({mth_row + 1},{d_start_col}) & ":" &ADDRESS({d_last_row},{d_last_col})'
    copy_rng = d_worksheet.Evaluate(d_rng_formula)
    print(copy_rng)

    #Template
    t_match_formula = f'=MATCH("Jan",{mth_row}:{mth_row},0)'
    t_start_col = t_worksheet.Evaluate(t_match_formula)

    t_rng_formula = f'=ADDRESS({mth_row+ 1} ,{t_start_col})'
    paste_rng = d_worksheet.Evaluate(t_rng_formula)
    print(paste_rng)

    excel.Application.Calculation = -4135  #set calculation option to manual to prevent values change when copying

    d_worksheet.Activate()
    d_worksheet.Range(copy_rng).Copy()
    t_worksheet.Range(paste_rng).PasteSpecial(Paste=-4163)

    excel.CutCopyMode = False
    excel.Application.Calculation = -4105 #reset calculation option to automatic

    
    #Get SQL queries
    sql_match_formula = f'=MATCH("SQL",{mth_row}:{mth_row},0)'
    sql_start_col = t_worksheet.Evaluate(sql_match_formula)
    sql_rng_formula = f'=ADDRESS({mth_row + 1},{sql_start_col}) & ":" &ADDRESS({d_last_row},{sql_start_col})'
    sql_rng = t_worksheet.Evaluate(sql_rng_formula)
    sql_source_rng = t_worksheet.Range(sql_rng)

    for row in range(1, sql_source_rng.Rows.Count + 1):
        cell_value = sql_source_rng.Cells(row, 1).Value
        sql_queries_lists.append(cell_value)
    

result = pymsgbox.confirm("Please review both tabs in the 'MI Serv Mgmt Overview-template' file to confirm that the data has been pasted accurately. Once confirmed, please click 'OK' to proceed.", 'Confirmation')
print(result)
if result == 'OK':
    print("No issue with both tabs in the 'MI Serv Mgmt Overview-template' file")
else:
    print('Clicked Cancel')
    exit()


# Execute SQL queries from each inner list
cursor.execute(f"delete from dbo.Dwh_Mi_ServMgmt where channel = 'CFS' and year = {yr};")
cursor.execute(f"delete from dbo.Dwh_Mi_ServMgmt where channel = 'Perform' and year = {yr};")
cursor.commit()
print(f"Deleted CFS & Perform's channel data for year {yr} in dbo.Dwh_Mi_ServMgmt table")

for sql_queries in sql_queries_lists:
    cursor.execute(sql_queries)
cursor.commit()  # Commit the transaction if necessary
print(f"Inserted CFS & Perform's channel data for year {yr} in dbo.Dwh_Mi_ServMgmt table")

cursor.close()


# Close the workbook without saving changes (optional)
Data_workbook.Close(False)
Template_workbook.Close(False)

excel.DisplayAlerts = True #To reopen warning dialog

# # Quit Excel
# excel.Quit()