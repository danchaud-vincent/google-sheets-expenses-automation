# GOOGLE SHEETS EXPENSES AUTOMATION PROJECT

This project create a google sheets with python to analyze your expenses with pivot_tables. It uses the google sheets API.

# Tutorial 

### 1. Setup Google Cloud Project

1. Create a new project on https://console.cloud.google.com/.
2. Enable the Google sheets API
3. OAuth consent screen -> External -> Fill out required fields with an email, also add this email under Test users.
4. Create credentials: Credentials -> OAuth 2.0 Client IDs -> Desktop App. Download the json file and save it as credentials.json file in the project folder.

### 2. Setup your python environment

> Create a new env : `python -m venv my_env` or `conda create -n my_env`

### 3. Install python libraries

- install google api : `pip install --upgrade google-api-python-client google-auth-httplib2 google-auth-oauthlib`

### 4. Create a csv file for your expenses 

> Example:
Date;Item;debit;credit
07/09/2022;Restaurant;41;
07/09/2022;Groceries;72,55;
07/09/2022;Insurance;103,84;
05/09/2022;Bar;7,8;
05/09/2022;Subway;12,38;

### 5. Create a sheet

Run the following code line:
> `python create_sheet.py`

The first time you run this line of code, the browser will open and you'll have to login. 

It creates and prints the sheetID. `Copy the id and put it in update_sheets.py (line 10)`

### 6. Update the sheets file

Run the following code line:
> `python update_sheet.py`