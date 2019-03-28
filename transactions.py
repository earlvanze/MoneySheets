from __future__ import print_function
import pickle
import os.path
import csv
import json
import traceback
from tkinter import filedialog
from tkinter import *
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

# Google Sheets stuff

# You should change these to match your own spreadsheet
GSHEET_ID = '10PmGsjxMXvIMDIig1QiS-YVYxqOClZEvEu8B9Z69MeA'
RANGE_NAME = 'Transactions!A:I'
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
]

def parse_csv():
    root = Tk()
    root.filename =  filedialog.askopenfilename(title = "Select file",filetypes = (("CSV files","*.csv"),("All files","*.*")))
#    root.filename = "Untitled.csv"
    print(root.filename)
    with open(root.filename, 'r') as csvfile:
        fieldnames = ("Name", "Current balance", "Account", "Transfers", "Description", "Merchant",
                      "Category", "Date", "Time", "Amount", "Currency", "Check #")
        reader = csv.DictReader(csvfile, fieldnames)
        next(reader, None)  # skip the headers

        output_data = []

        for row in reader:
            if not row["Name"]:
                try:
                    data = []
                    data.append(row["Date"])
                    data.append(row["Account"])
                    data.append(row["Transfers"])
                    data.append(row["Description"])
                    data.append(row["Merchant"])
                    if float(row["Amount"]) > 0:
                        data.append(row["Amount"])
                        data.append("")
                    else:
                        data.append("")
                        data.append(row["Amount"])
                    data.append("")
                    data.append(row["Category"])
                    print(data)
                    output_data.append(data)
                except:
                    traceback.print_exc()
                    continue
            else:
                next(reader, None)  # not transaction data, skip the row


    result = append_to_gsheet(output_data, GSHEET_ID, RANGE_NAME)
    return result


def append_to_gsheet(output_data=[], gsheet_id = GSHEET_ID, range_name = RANGE_NAME):
    creds = None
    # The file token.pickle stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server()
        # Save the credentials for the next run
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    service = build('sheets', 'v4', credentials=creds)

    # Call the Sheets API
    body = {
        'values': output_data
    }
    try:
        result = service.spreadsheets().values().append(
            spreadsheetId=gsheet_id, range=range_name,
            valueInputOption='USER_ENTERED', body=body).execute()
        message = ('{0} rows updated.'.format(DictQuery(result).get('updates/updatedRows')))
        return message
    except Exception as err:
        traceback.print_exc()
        return json.loads(err.content.decode('utf-8'))['error']['message']


# Used to search for keys in nested dictionaries and handles when key does not exist
# Example: DictQuery(dict).get("dict_key/subdict_key")
class DictQuery(dict):
    def get(self, path, default = None):
        keys = path.split("/")
        val = None

        for key in keys:
            if val:
                if isinstance(val, list):
                    val = [ v.get(key, default) if v else None for v in val]
                else:
                    val = val.get(key, default)
            else:
                val = dict.get(self, key, default)

            if not val:
                break;

        return val


print(parse_csv())
