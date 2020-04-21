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

    service = build('drive', 'v3', credentials=creds)

    # Call the Drive v3 API
    # for folder in folders:
    #     current = {
    #         'name': folder,
    #         'mimeType': 'application/vnd.google-apps.folder'
    #     }
    #     file = service.files().create(body=current,
    #                                 fields='id').execute()

    #     root_folder_id = file.get('id')

    folderTreeExcel = r"./foldertree/folderTree.xlsx"

    # open the original excel file we want to split
    workbook = pandas.read_excel(folderTreeExcel, index_col = False) 
    # save total number of rows
    numberRows = len(workbook)

    	
    # iterate over excel
    # for i in range(0, len(workbook)):
    #     if workbook['Root1'][i] != None:
    #         current = {
    #             'name': workbook['Root1'][i],
    #             'mimeType': 'application/vnd.google-apps.folder'        
    #         }
    #         file = service.files().create(body=current,
    #                             fields='id').execute()
    #         folder_id = file.get('id')

    for col in list(workbook.columns): 
        if not pandas.isna(col):      
            print(col)
            current = {
                'name': str(col),
                'mimeType': 'application/vnd.google-apps.folder'        
            }
            file = service.files().create(body=current, fields='id').execute()
            folder_id = file.get('id')
            sub_folder = {
                'name': 'ciao', 
                'mimeType': 'application/vnd.google-apps.folder',
                'parents': [{'id': folder_id}]
            }            
            file = service.files().create(body=sub_folder, fields='id').execute()

            # for values in workbook[col]:
            #     if not pandas.isna(values):
            #         print(values)    
            #         current = {
            #             'name': str(values),
            #             'mimeType': 'application/vnd.google-apps.folder',
            #             'parents': folder_id       
            #         }
            #         file = service.files().create(body=current,
            #                     fields='id').execute() 
            #         folder_id = file.get('id')                   

if __name__ == '__main__':
    main()