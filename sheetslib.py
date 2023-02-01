#https://developers.google.com/sheets/api/quickstart/python

import os.path

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# If modifying these scopes, delete the file token.json.
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

# DONE
def authorize():
    creds = None
    # The file token.json stores the user's access and refresh tokens, and is created
    # automatically when the authorization flow completes for the first time.
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.json', 'w') as token:
            token.write(creds.to_json())
    try:
        service = build('sheets', 'v4', credentials=creds)
    except HttpError as error:
        print(f"An error occurred: {error}")
        return error
    return creds, service

def update_values(spreadsheet_id, range_name, value_input_option, values, service):
    body = {
        'values': values
    }
    try:
        result = service.spreadsheets().values().update(spreadsheetId=spreadsheet_id, range=range_name, valueInputOption=value_input_option, body=body).execute()
        print(result)
        print(f"{result.get('updatedCells')} cells updated.")
        return result
    except HttpError as error:
        print(f"An error occurred: {error}")
        return error
            
def read_values(spreadsheet_id, range_name, service):
    try:
        # Call the Sheets API
        sheet = service.spreadsheets()
        result = sheet.values().get(spreadsheetId=spreadsheet_id, range=range_name).execute()
        values = result.get('values', [])

        if not values:
            print('No data found.')
            return
        return values
    
    except HttpError as err:
        print(err)

def add_new_tab(spreadsheet_id, tab_name, service):
    add_sheet_body = {
        'requests': [
            {
                'addSheet': {
                    'properties': {
                        'title': tab_name
                    }
                }
            }
        ]
    }
    try:
        response = service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body=add_sheet_body).execute()
        return response['replies'][0]['addSheet']['properties']['sheetId']
    except HttpError as err:
        print(err)
    

def get_sheet_id(spreadsheet_id, tab_name, service):
    try:
        response = service.spreadsheets().get(spreadsheetId=spreadsheet_id, fields='sheets.properties').execute()
        for sheet in response['sheets']:
            if sheet['properties']['title'] == tab_name:
                return sheet['properties']['sheetId']
    except HttpError as err:
        print(err)

def create_sheet_or_get_sheet_id(spreadsheet_id, tab_name, service):
    try:
        sheet_id = add_new_tab(spreadsheet_id, tab_name, service)
    except HttpError as err:
        print(err)
        
    if sheet_id is None:
        sheet_id = get_sheet_id(spreadsheet_id, tab_name, service)
    
    print (sheet_id)
    return sheet_id
