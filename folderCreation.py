from __future__ import print_function
import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import pandas
import numpy as np

# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/drive.appdata', 'https://www.googleapis.com/auth/drive.file', 'https://www.googleapis.com/auth/drive.install']

folderTreeExcel = r"./foldertree/folderTree.xlsx"

def main():
    """Shows basic usage of the Drive v3 API.
    Prints the names and ids of the first 10 files the user has access to.
    """
    creds = None
    # The file token.pickle stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists(r'./credential/token.pickle'):
        with open('./credential/token.pickle', 'rb') as token:
            creds = pickle.load(token)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                './credential/credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('./credential/token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    # service Instance
    service = build('drive', 'v3', credentials=creds)
    
    # excel file where the folder tree is
    folderTreeExcel = r"./foldertree/folderTree.xlsx"

    # open the original excel file we want to split
    workbook = pandas.read_excel(folderTreeExcel, index_col = False) 
    # save total number of rows
    numberRows = len(workbook)

    # define function to create a folder 
    def create_folder(values, folder_id):
        # if folder_id is none means that is the root Folder, it has no parent
        if folder_id is None:
            current = {
                'name': str(values),
                'mimeType': 'application/vnd.google-apps.folder',      
            }
        # If folder_id is not none means that is a child folder which folder_id is the parent   
        else:
            current = {
                'name': str(values),
                'mimeType': 'application/vnd.google-apps.folder',
                'parents': [folder_id]       
            }
        # create it on GoogleDrive
        file = service.files().create(body=current, fields='id').execute() 
        # get and return the file id          
        return file.get('id')

    # Iterate over he excel File
    for col in list(workbook.columns):         
        if not pandas.isna(col): 
            # Assign the value to a root folder name 
            folder_id = create_folder(col, folder_id=None)
            # Iterate over next lines to get the subfolder names
            for values in workbook[col]:
                if not pandas.isna(values):                     
                    folder_id = create_folder(values, folder_id)    

if __name__ == '__main__':
    main()