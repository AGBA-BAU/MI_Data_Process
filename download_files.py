from office365_api import SharePoint
import re
import sys, os
from pathlib import PurePath
import configparser
import shutil

#read config file
config = configparser.ConfigParser() 
config.read('config.ini')


yrmonth = config.get('Reporting','YRMONTH')

# 1 args = locate or remote folder_dest
DATA_FOLDER_DEST =str(os.path.dirname(__file__))+"\MI_Serv_Mgmt_Overview\Data"
# print(DATA_FOLDER_DEST)
TEMPLATE_FOLDER_DEST =str(os.path.dirname(__file__))+"\MI_Serv_Mgmt_Overview\Template"
# print(TEMPLATE_FOLDER_DEST)
# 2 args = SharePoint folder name. May include subfolders YouTube/2022
DATA_FOLDER_NAME = config.get('download','DATA_FOLDER_NAME') + yrmonth
TEMPLATE_FOLDER_NAME = config.get('download','TEMPLATE_FOLDER_NAME')
# 3 args = SharePoint file name. This is used when only one file is being downloaded
# If all files will be downloaded, then set this value as "None"
TEMPLATE_FILE_NAME = config.get('download','TEMPLATE_FILE_NAME')
# 4 args = SharePoint file name pattern
# If no pattern match files are required to be downloaded, then set this value as "None"
DATA_FILE_NAME_PATTERN = config.get('download','DATA_FILE_NAME_PATTERN')

# print(DATA_FOLDER_NAME )
# print(TEMPLATE_FOLDER_NAME)
# print(TEMPLATE_FILE_NAME)
# print(DATA_FILE_NAME_PATTERN)

def save_file(file_n, file_obj):
    if re.search(rf"\b{re.escape('template')}\b", file_n):
        file_dir_path = PurePath(TEMPLATE_FOLDER_DEST, file_n)
    else:
        file_dir_path = PurePath(DATA_FOLDER_DEST, file_n)

    with open(file_dir_path, 'wb') as f:
        f.write(file_obj)

def get_file(file_n, folder):
    file_obj = SharePoint().download_file(file_n, folder)
    save_file(file_n, file_obj)

def get_files(folder):
    files_list = SharePoint()._get_files_list(folder)
    for file in files_list:
        get_file(file.name, folder)

def get_files_by_pattern(keyword, folder):
    files_list = SharePoint()._get_files_list(folder)
    for file in files_list:
        if re.search(keyword, file.name):
            get_file(file.name, folder)


shutil.rmtree(DATA_FOLDER_DEST)
os.mkdir(DATA_FOLDER_DEST)

shutil.rmtree(TEMPLATE_FOLDER_DEST)
os.mkdir(TEMPLATE_FOLDER_DEST)

# To get data from Sharepoint
get_files_by_pattern(TEMPLATE_FILE_NAME, TEMPLATE_FOLDER_NAME)

# To get template from Sharepoint
get_files_by_pattern(DATA_FILE_NAME_PATTERN, DATA_FOLDER_NAME)  
