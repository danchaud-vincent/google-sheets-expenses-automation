# google libraries
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

import os

# scopes
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]


# ----- functions ---------

def create_authorized_service():
    """
    Create an authorized service for google sheets using a personal credentials.json file

    Returns:
    - service: Google Sheets service
    """
    creds = None

    # The file token.json stores the user's access and refresh tokens
    # It is created automatically when the authorization flow completes for the first time
    if os.path.exists("token.json"):
        creds = Credentials.from_authorized_user_file("token.json", SCOPES)

    # Check if there are no credentials available
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file("credentials.json", SCOPES)
            creds = flow.run_local_server(port=0)

        # save the credentials for the next run
        with open("token.json", "w") as token:
            token.write(creds.to_json())


    try:
        # create sheets service
        service = build("sheets", "v4", credentials=creds)

        # console log for the user
        print("Sheets service created successfully")
        
        return service

    except HttpError as err:
        print(err)


def create_sheet(service):
    """
    Create a google spreadsheet using an authorized sheets service

    Arguments:
    - service: authorized google sheets service
    """
    
    # sheet body used for the google sheet
    sheet_body ={
        "properties": {
            "title": "expenses"
        },
        "sheets": [
            {
                "properties": {
                    "sheetId": 0,
                    "title": "dashboard"
                }
            },
            {
                "properties": {
                    "sheetId": 1,
                    "title": "pivot_tables"
                    }
            },
            {
                "properties": {
                    "sheetId": 2,
                    "title": "data"
                    }
            }
        ]
    }

    # google sheets file
    sheets_file = service.spreadsheets().create(body=sheet_body).execute()

    # console log
    message_user = f"""Created sheet:
    {sheets_file["spreadsheetUrl"]}
    {sheets_file["spreadsheetId"]}"""

    print(message_user)



def update_values(service, spreadsheet_id, data):
    """
    Update the values of the google spreadsheets

    Arguments:
    - service : service google sheets
    - spreadsheet_id (str): id of the google spreadsheet
    - data (list): headers and rows to store in the google sheets
    """

    range = "data!A1"

    service.spreadsheets().values().update(
        spreadsheetId=spreadsheet_id,
        valueInputOption="USER_ENTERED",
        range=range,
        body={"majorDimension": "ROWS", "values": data}
    ).execute()


def create_pivot_tables(service, spreadsheet_id, data, sheetId=1,sheetId_source=2):
    """
    Create the pivot tables in the sheet 1

    Arguments:
    - service : google sheets service
    - spreadsheet_id (string): id of the google spreadsheet
    - sheetId (int) : if of the sheet where store the pivot tables
    """

    requests = {
        "requests": [
            {
                "updateCells": {
                    "rows": {
                        "values": [
                            {
                                "pivotTable": {
                                    "source": {
                                        "sheetId": sheetId_source,
                                        "startRowIndex": 0,
                                        "endRowIndex": len(data),
                                        "startColumnIndex": 0,
                                        "endColumnIndex":5
                                    },
                                    "rows":[
                                        {
                                            "sourceColumnOffset": 2,
                                            "showTotals": False,
                                            "sortOrder": "ASCENDING",
                                            "valueBucket": {}
                                        }
                                    ],
                                    "values":[
                                            {
                                            "summarizeFunction": "SUM",
                                            "sourceColumnOffset": 3
                                        }
                                    ],
                                    "valueLayout" : "HORIZONTAL"
                                }
                            }
                        ]
                    },
                    "start": {
                        "sheetId": sheetId,
                        "rowIndex": 0,
                        "columnIndex": 0
                    },
                    "fields": "pivotTable"
                }
            },
            {
                "updateCells": {
                    "rows": {
                        "values": [
                            {
                                "pivotTable": {
                                    "source": {
                                        "sheetId": sheetId_source,
                                        "startRowIndex": 0,
                                        "endRowIndex": len(data),
                                        "startColumnIndex": 0,
                                        "endColumnIndex":5
                                    },
                                    "rows":[
                                        {
                                            "sourceColumnOffset": 1,
                                            "showTotals": False,
                                            "sortOrder": "ASCENDING",
                                            "valueBucket": {}
                                        }
                                    ],
                                    "values":[
                                            {
                                            "summarizeFunction": "SUM",
                                            "sourceColumnOffset": 3
                                        }
                                    ],
                                    "valueLayout" : "HORIZONTAL"
                                }
                            }
                        ]
                    },
                    "start": {
                        "sheetId": sheetId,
                        "rowIndex": 0,
                        "columnIndex": 3
                    },
                    "fields": "pivotTable"
                }
            },
            {
                "updateCells": {
                    "rows": {
                        "values": [
                            {
                                "pivotTable": {
                                    "source": {
                                        "sheetId": sheetId_source,
                                        "startRowIndex": 0,
                                        "endRowIndex": len(data),
                                        "startColumnIndex": 0,
                                        "endColumnIndex":5
                                    },
                                    "rows":[
                                        {
                                            "sourceColumnOffset": 0,
                                            "showTotals": False,
                                            "sortOrder": "ASCENDING",
                                            "groupRule":{
                                                "dateTimeRule":{
                                                    "type": "YEAR"
                                                }
                                            }
                                        }
                                    ],
                                    "values":[
                                        {
                                            "summarizeFunction": "SUM",
                                            "sourceColumnOffset": 3
                                        },
                                        {
                                            "summarizeFunction": "SUM",
                                            "sourceColumnOffset": 4
                                        }
                                    ],
                                    "valueLayout" : "HORIZONTAL"
                                }
                            }
                        ]
                    },
                    "start": {
                        "sheetId": sheetId,
                        "rowIndex": 0,
                        "columnIndex": 6
                    },
                    "fields": "pivotTable"
                }
            },
            {
                "updateCells": {
                    "rows": {
                        "values": [
                            {
                                "pivotTable": {
                                    "source": {
                                        "sheetId": sheetId_source,
                                        "startRowIndex": 0,
                                        "endRowIndex": len(data),
                                        "startColumnIndex": 0,
                                        "endColumnIndex":5
                                    },
                                    "rows":[
                                        {
                                            "sourceColumnOffset": 0,
                                            "showTotals": False,
                                            "sortOrder": "ASCENDING",
                                            "groupRule":{
                                                "dateTimeRule":{
                                                    "type": "YEAR_MONTH"
                                                }
                                            }
                                        }
                                    ],
                                    "values":[
                                        {
                                            "summarizeFunction": "SUM",
                                            "sourceColumnOffset": 3
                                        },
                                        {
                                            "summarizeFunction": "SUM",
                                            "sourceColumnOffset": 4
                                        }
                                    ],
                                    "valueLayout" : "HORIZONTAL"
                                }
                            }
                        ]
                    },
                    "start": {
                        "sheetId": sheetId,
                        "rowIndex": 40,
                        "columnIndex": 0
                    },
                    "fields": "pivotTable"
                }
            },
            {
                "updateCells":{
                    "rows":{
                        "values":[
                            {
                                "pivotTable":{
                                    "source":{
                                        "sheetId": sheetId_source,
                                        "startRowIndex": 0,
                                        "endRowIndex": len(data),
                                        "startColumnIndex": 0,
                                        "endColumnIndex": 5
                                    },
                                    "rows":[
                                        {
                                            "sourceColumnOffset": 0,
                                            "showTotals": False,
                                            "sortOrder": "ASCENDING",
                                            "groupRule":{
                                                "dateTimeRule":{
                                                    "type": "YEAR_MONTH"
                                                }
                                            }
                                        },
                                       
                                    ],
                                    "columns": [
                                        {
                                            "sourceColumnOffset": 2,
                                            "showTotals": False,
                                            "sortOrder": "ASCENDING"
                                        }
                                    ],
                                    "values":[
                                        {
                                            "sourceColumnOffset": 3,
                                            "summarizeFunction": "SUM"
                                        }
                                    ],
                                    "valueLayout" : "HORIZONTAL"
                                }
                            }
                        ]
                    },
                    "start":{
                        "sheetId": sheetId,
                        "rowIndex": 40,
                        "columnIndex": 6
                    },
                    "fields": "pivotTable"
                }
            }
        ]
    }

    service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body=requests).execute()


def format_cells(service, spreadsheet_id, sheetId=2):
    """
    Set a custom datetime and format for a range in the sheet 2

    Arguments:
    - service : google sheets service
    - spreadsheet_id (string): id of the google spreadsheet
    - sheetId (int): id of the sheet where apply the formatting
    """

    requests = {
        "requests": [
            {
                "repeatCell": {
                    "range": {
                        "sheetId": sheetId,
                        "startRowIndex": 1,
                        "startColumnIndex": 0,
                        },
                        "cell": {
                            "userEnteredFormat": {
                                "horizontalAlignment" : "CENTER"
                            }
                        },
                        "fields": "userEnteredFormat.horizontalAlignment"
                }
            },
            {
                "repeatCell": {
                    "range": {
                        "sheetId": sheetId,
                        "startRowIndex": 1,
                        "startColumnIndex": 0,
                        "endColumnIndex": 1
                        },
                        "cell": {
                            "userEnteredFormat": {
                                "numberFormat": {
                                    "type": "DATE",
                                    "pattern": "dd/mm/yyyy"
                                    }
                            }
                        },
                        "fields": "userEnteredFormat.numberFormat"
                }
            },
            {
                "repeatCell":{
                    "range":{
                        "sheetId": sheetId,
                        "startRowIndex":1,
                        "startColumnIndex":3,
                        "endColumnIndex":5
                    },
                    "cell": {
                        "userEnteredFormat": {
                            "numberFormat": {
                                "type": "NUMBER"
                            }
                        }
                    },
                    "fields": "userEnteredFormat.numberFormat"
                }
            },
            {
                "repeatCell":{
                    "range":{
                        "sheetId": sheetId,
                        "startRowIndex":0,
                        "endRowIndex": 1,
                        "startColumnIndex":0,
                        "endColumnIndex":5
                    },
                    "cell": {
                        "userEnteredFormat": {
                            "backgroundColor": {
                                    "red": 0.0,
                                    "green": 0.0,
                                    "blue": 0.0
                            },
                            "horizontalAlignment" : "CENTER",
                            "textFormat": {
                                "foregroundColor": {
                                    "red": 1.0,
                                    "green": 1.0,
                                    "blue": 1.0
                                },
                                
                                "bold": True,
                                "fontSize": 11
                            }
                        }
                    },
                    "fields": "userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)"
                }
            }
        ]
    }

    service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body=requests).execute()