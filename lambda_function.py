import re
import os, io
import boto3
from botocore.exceptions import ClientError
from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.user_credential import UserCredential
from office365.sharepoint.files.file import File

# 1 args = SharePoint folder name. May include subfolders YouTube/2022
SHAREPOINT_FOLDER_NAME = os.environ['sharepoint_folder_name']
# Office 365
USERNAME = os.environ['sharepoint_user']
PASSWORD = os.environ['sharepoint_password']
SHAREPOINT_URL = os.environ['sharepoint_url']
SHAREPOINT_SITE = os.environ['sharepoint_site']
SHAREPOINT_DOC = os.environ['sharepoint_doc_library']
# AWS
AWS_ACCESS_KEY_ID = os.environ['aws_access_key']
AWS_SECRET_ACCESS_KEY = os.environ['aws_secret_access_key']
BUCKET = os.environ['bucket_name']
BUCKET_SUBFOLDER = os.environ['bucket_subfolder']

def lambda_handler(event, context):   
    get_files(SHAREPOINT_FOLDER_NAME)


# functions used for aws
def upload_file_to_s3(file_obj, bucket, file_name):
    s3_client = boto3.client(
        's3',
        aws_access_key_id=AWS_ACCESS_KEY_ID,
        aws_secret_access_key=AWS_SECRET_ACCESS_KEY
    )
    try:
        response = s3_client.upload_fileobj(io.BytesIO(file_obj), bucket, file_name)
    except ClientError as e:
        return False
    return True
        

def bucket_subfolder_build(BUCKET_SUBFOLDER, file_name):
    if BUCKET_SUBFOLDER != '':
        file_path_name = '/'.join([BUCKET_SUBFOLDER, file_name])
        return file_path_name
    else:
        return file_name

def get_file(file_n, folder):
    file_obj = SharePoint().download_file(file_n, folder)
    file_name = bucket_subfolder_build(BUCKET_SUBFOLDER, file_n)
    upload_file_to_s3(file_obj.content, BUCKET, file_name)

def get_files(folder):
    files_list = SharePoint().download_files(folder)
    for file in files_list:
        get_file(file.name, folder)

def get_files_by_pattern(pattern, folder):
    files_list = SharePoint().download_files(folder)
    for file in files_list:
        if re.search(pattern, file['Name']):
            get_file(file['Name'], folder)


class SharePoint:
    def auth(self):
        conn = ClientContext(SHAREPOINT_SITE).with_credentials(UserCredential(
            USERNAME,
            PASSWORD
        ))

        return conn

    def download_file(self, file_name, folder_name):
        conn = self.auth()
        file_url = f'/sites/Development/{SHAREPOINT_DOC}{folder_name}/{file_name}'
        file = File.open_binary(conn, file_url)
        return file

    def _get_files_list(self, folder_name):
        conn = self.auth()
        target_folder_url = f'{SHAREPOINT_DOC}/{folder_name}'
        root_folder = conn.web.get_folder_by_server_relative_url(target_folder_url)

        root_folder.expand(["Files", "Folders"]).get().execute_query()
        return root_folder.files
            
    def download_files(self, folder_name):
        self._files_list = self._get_files_list(folder_name)
        return self._files_list