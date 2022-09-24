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