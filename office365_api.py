# from distutils.command.config import config
# from distutils.command.upload import upload
from urllib import response
from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.user_credential import UserCredential
from office365.sharepoint.files.file import File
import datetime
import configparser
import re 

#read config file
config = configparser.ConfigParser() 
config.read('config.ini')

USERNAME = config.get('SharePoint','USERNAME')
PASSWORD = config.get('SharePoint','PASSWORD')
SHAREPOINT_SITE = config.get('SharePoint','SHAREPOINT_SITE')
SHAREPOINT_SITE_NAME = config.get('SharePoint','SHAREPOINT_SITE_NAME')
SHAREPOINT_DOC = config.get('SharePoint','SHAREPOINT_DOC')

class SharePoint:
    def _auth(self):
        conn = ClientContext(SHAREPOINT_SITE).with_credentials(
            UserCredential(
                USERNAME,
                PASSWORD
            )
        )
        return conn

    def _get_files_list(self, folder_name):
        conn = self._auth()
        target_folder_url = f'{SHAREPOINT_DOC}/{folder_name}'
        root_folder = conn.web.get_folder_by_server_relative_url(target_folder_url)
        root_folder.expand(["Files", "Folders"]).get().execute_query()
        return root_folder.files

    def check_file_version(self, file_name, file_list):
        file_ver=0
        for file in file_list:
            check_name = re.sub('v1.(.+?)\)', "v1.", str.replace(file_name,".xlsx","")) 
            if check_name  in file.name :
                file_ver+=1       
        return file_ver

    def rename_file_name(self, file_name, file_version, replace):
        if replace == 'true' :
            file_version-=1
        if file_version > 0 :
            if file_name.find("v1.")>0:
                file_name_new = str.replace(file_name, "v1.0", "v1." + str(file_version))
            else:
                file_name_new = str.replace(file_name, ".xlsx", "_v" + str(file_version+1) + ".xlsx")  
        else:
            file_name_new = file_name
        return file_name_new

    def download_file(self, file_name, folder_name):
        conn = self._auth()
        file_url = f'/sites/{SHAREPOINT_SITE_NAME}/{SHAREPOINT_DOC}/{folder_name}/{file_name}'
        file = File.open_binary(conn, file_url)
        return file.content

    def upload_file(self, file_name, folder_name, content, replace):
        conn = self._auth()
        target_folder_url = f'/sites/{SHAREPOINT_SITE_NAME}/{SHAREPOINT_DOC}/{folder_name}'
        target_folder = conn.web.get_folder_by_server_relative_path(target_folder_url)
        file_name_new = re.sub('{(.+?)}', "", file_name) 
        #create folder if not exists    
        try:
            #get the version number and rename accordingly
            uploaded_file_list = self._get_files_list(folder_name)
            file_ver = self.check_file_version(file_name_new, uploaded_file_list)
            file_name_new = self.rename_file_name(file_name_new, file_ver, replace)
            response = target_folder.upload_file(file_name_new, content).execute_query()
        except:
            target_folder = conn.web.folders.add(target_folder_url).execute_query()
            response = target_folder.upload_file(file_name_new, content).execute_query()
        return response

    def get_list(self, list_name):
        conn = self._auth()
        target_list = conn.web.lists.get_by_title(list_name)
        items = target_list.items.get().execute_query()
        return items

