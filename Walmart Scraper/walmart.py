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
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from bs4 import BeautifulSoup

options = webdriver.ChromeOptions()
options.add_argument("--user-data-dir=/Users/Vanze/Library/Application Support/Google/Chrome/Profile 6")
browser = webdriver.Chrome('./chromedriver', options = options)

delay = 30

# Google Sheets stuff
# You should change these to match your own spreadsheet
if os.path.exists('gsheet_id.txt'):
    with open('gsheet_id.txt', 'r') as file:
       json_repr = file.readline()
       data = json.loads(json_repr)
       GSHEET_ID = data["GSHEET_ID"]
       RANGE_NAME = data["RANGE_NAME"]
else:
    GSHEET_ID = '10PmGsjxMXvIMDIig1QiS-YVYxqOClZEvEu8B9Z69MeA'
    RANGE_NAME = 'Transactions!A:I'
    
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
]

def get_urls(filename):
    urls = []
    with open(filename, 'r') as file:
        urls = file.readlines()
        return urls


def get_items(urls):
    for url in urls:
        browser.get(url)
        try:
            order_details = WebDriverWait(browser, delay).until(EC.presence_of_element_located((By.CLASS_NAME, "order-details")))
            items = order_details.find_elements_by_class_name('order-item-left')
            for item in items:
                product_name = item.find_elements_by_class_name('product_name')
                print(product_name)
        except:
            traceback.print_exc()
            pass

#    result = append_to_gsheet(output_data, GSHEET_ID, RANGE_NAME)
#    return result


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

urls = get_urls("urls.txt")
get_items(urls)
browser.quit()
