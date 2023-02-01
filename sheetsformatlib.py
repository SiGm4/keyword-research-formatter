from googleapiclient.errors import HttpError

def align_vertical_middle_all(spreadsheet_id, sheet_id, service):    
    formatting_body = {
        'requests': [
            {
                'repeatCell': {
                    'range': {'sheetId': sheet_id},
                    'cell': {'userEnteredFormat': {'verticalAlignment': 'MIDDLE'}},
                    'fields': 'userEnteredFormat.verticalAlignment'
                }
            }
        ]
    }
    service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body=formatting_body).execute()

def bold_center_header_row(spreadsheet_id, sheet_id, service):
    formatting_body = {
        'requests': [
            {
                'repeatCell': {
                    'range': {'sheetId': sheet_id, 'endRowIndex': 1},
                    'cell': {'userEnteredFormat': {'horizontalAlignment': 'CENTER', "textFormat": {"bold": True}}},
                    'fields': 'userEnteredFormat(horizontalAlignment,textFormat)'
                }
            }
        ]
    }
    service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body=formatting_body).execute()

def color_header_row(spreadsheet_id, sheet_id, service, col_number):
    color = {'r': 246, 'g': 178, 'b': 107} #TODO add it to config file
    formatting_body = {
        'requests': [
            {
                'repeatCell': {
                    'range': {'sheetId': sheet_id, 'endRowIndex': 1, 'endColumnIndex': col_number},
                    'cell': {'userEnteredFormat': {'backgroundColor': {"red": color['r']/255, "green": color['g']/255, "blue": color['b']/255}}},
                    'fields': 'userEnteredFormat.backgroundColor'
                }
            }
        ]
    }
    service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body=formatting_body).execute()

def center_columns(spreadsheet_id, sheet_id, service, start_col, end_col):
    formatting_body = {
        'requests': [
            {
                'repeatCell': {
                    'range': {'sheetId': sheet_id, 'startColumnIndex': start_col, 'endColumnIndex': end_col},
                    'cell': {'userEnteredFormat': {'horizontalAlignment': 'CENTER'}},
                    'fields': 'userEnteredFormat.horizontalAlignment'
                }
            }
        ]
    }
    service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body=formatting_body).execute()

def auto_resize_column(spreadsheet_id, sheet_id, service, col_no):
    formatting_body = {
        "requests": [
            {
                "autoResizeDimensions": {
                    "dimensions": {
                        "sheetId": sheet_id,
                        "dimension": "COLUMNS",
                        "startIndex": col_no-1,
                        "endIndex": col_no
                    }
                }
            }
        ]
    }
    service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body=formatting_body).execute()

def resize_columns(spreadsheet_id, sheet_id, service, size, start_col, end_col):
    formatting_body = {
        "requests": [
            {
                "updateDimensionProperties": {
                    "range": {
                        "sheetId": sheet_id,
                        "dimension": "COLUMNS",
                        "startIndex": start_col,
                        "endIndex": end_col
                    },
                    "properties": {
                        "pixelSize": size
                    },
                    "fields": "pixelSize"
                }
            }
        ]
    }
    service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body=formatting_body).execute()    

def set_columns_wrap_strategy(spreadsheet_id, sheet_id, service, strategy, start_col, end_col):
    formatting_body = {
        'requests': [
            {
                'repeatCell': {
                    'range': {'sheetId': sheet_id, 'startColumnIndex': start_col, 'endColumnIndex': end_col},
                    'cell': {'userEnteredFormat': {'wrapStrategy': strategy}},
                    'fields': 'userEnteredFormat.wrapStrategy'
                }
            }
        ]
    }
    service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body=formatting_body).execute()

def color_row(spreadsheet_id, sheet_id, service, row_no):
    color = {'r': 204, 'g': 204, 'b': 204} #TODO add it to config file # gray
    formatting_body = {
        'requests': [
            {
                'repeatCell': {
                    'range': {'sheetId': sheet_id, 'startRowIndex': row_no-1,'endRowIndex': row_no},
                    'cell': {'userEnteredFormat': {'backgroundColor': {"red": color['r']/255, "green": color['g']/255, "blue": color['b']/255}}},
                    'fields': 'userEnteredFormat.backgroundColor'
                }
            }
        ]
    }
    service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body=formatting_body).execute()

 

    
def add_checkboxes_with_color(spreadsheet_id, sheet_id, service, range_object):
    color = {'r': 246, 'g': 178, 'b': 107} #TODO add it to config file
    formatting_body = {
        'requests': [
            {
                'repeatCell': {
                    'range': {
                        'sheetId': sheet_id,
                        'startRowIndex': range_object['startRowIndex'],
                        'endRowIndex': range_object['endRowIndex'],
                        'startColumnIndex': range_object['startColumnIndex'],
                        'endColumnIndex': range_object['endColumnIndex']
                    },
                    'cell': {
                        'dataValidation': {'condition': {'type': 'BOOLEAN'}},
                        'userEnteredFormat': {
                            'textFormat': {
                                'foregroundColorStyle': {"rgbColor": {"red": color['r']/255, "green": color['g']/255, "blue": color['b']/255}}
                            }
                        }
                    },
                    'fields': 'dataValidation,userEnteredFormat.TextFormat.foregroundColorStyle'
                }
            }
        ]
    }
    service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body=formatting_body).execute()
