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

def parse_csv():
    root = Tk()
    root.filename =  filedialog.askopenfilename(title = "Select file",filetypes = (("CSV files","*.csv"),("All files","*.*")))
#    root.filename = "Untitled.csv"
    print(root.filename)
    with open(root.filename, 'r') as csvfile:
        fieldnames = ("Name", "Current balance", "Account", "Transfers", "Description", "Merchant",
                      "Category", "Date", "Time", "Memo", "Amount", "Currency", "Check #")
        reader = csv.DictReader(csvfile, fieldnames)
        next(reader, None)  # skip the headers

        output_data = []

        for row in reader:
            if not row["Name"]:
                try:
#                    if "ORIG CO NAME:" in row["Description"]:
#                        continue
                    data = []
                    data.append(row["Date"])
                    data.append(row["Account"])
                    data.append(row["Description"].split('WEB ID:')[0])
                    data.append(row["Merchant"])

                    # Split Amount column to Incomes and Expenses columns
                    try:
                        if float(row["Amount"].replace(',', '')) > 0:
                            data.append(row["Amount"])
                            data.append("")
                        else:
                            data.append("")
                            data.append(row["Amount"])
                    except ValueError:
                        print("Error with row: ", row["Amount"])
                        pass

                    # Business Category
                    if "88 Madison Joint Account" in row["Account"]:
                        data.append("88 Madison Ave")
                    elif "88 Madison Ave" in row["Account"]:
                        data.append("88 Madison Ave")
                    elif "110 N Saddle Dr" in row["Account"]:
                        data.append("110 N Saddle Dr")
                    elif "90 Madison Ave" in row["Account"]:
                        data.append("90 Madison Ave")
                    elif "ECO Systems Checking" in row["Account"]:
                        data.append("724 3rd Ave")
                    elif "Your Second Home Checking" in row["Account"]:
                        data.append("Your Second Home")
                    elif "Dover Holdings Checking" in row["Account"]:
                        data.append("3880 Dover St")
                    else:
                        data.append("")

                    # Autofill Category column
                    # Venmo cleaning transactions
                    if row["Account"] == "88 Madison Joint Account" and "VENMO PAYMENT" in row["Description"]:
                        data[2] = data[2].replace(" WEB ID: 3264681992", "")
                        if float(row["Amount"]) == -25:
                            data[2] += " - 1 hr cleaning"
                        else:
                            data[2] += " - {0} hrs cleaning".format(float(row["Amount"]) / -20.0) # number of hours @ $20/hr
                        data[3] = "Florence Odongo"
                        data.append("Cleaning & Maintenance")

                    # Airbnb or Homeaway Income
                    elif row["Account"] == "88 Madison Joint Account" and "AIRBNB PAYMENTS" in row["Description"]:
                        data[3] = "Airbnb"
                        data.append("Rental")
                    elif row["Account"] == "88 Madison Joint Account" and "VRBO" in row["Description"]:
                        data.append("Rental")

                    # Mortgage Transactions
                    elif "DITECH" in row["Description"]:
                        data[3] = "Ditech Financial"
                        data.append("Mortgage")
                    elif "NewRez" in row["Description"]:
                        data[3] = "NewRez LLC"
                        data.append("Mortgage")
                    elif "RoundPoint" in row["Description"]:
                        data[3] = "RoundPoint Mortgage Servicing Corporation"
                        data[6] = "724 3rd Ave"
                        data.append("Mortgage")

                    # Subscriptions: PriceLabs, Smartbnb, Arcadia, Netflix, TWC, BillFixers, Comcast, RedPocket, Tello
                    elif "PRICELABS" in row["Description"]:
                        data[3] = "PriceLabs"
                        data.append("Advertising")
                    elif "SMARTBNB" in row["Description"]:
                        data[3] = "Smartbnb"
                        data.append("Advertising")
                    elif "ARCADIA" in row["Description"]:
                        data[3] = "Arcadia Power"
                        data.append("Utilities")
                    elif "XCEL" in row["Description"]:
                        data[3] = "Xcel Energy"
                        data.append("Utilities")
                    elif "NETFLIX" in row["Description"]:
                        data[3] = "Netflix"
                        data.append("Subscriptions")
                    elif "TWC" in row["Description"]:
                        data[3] = "Spectrum"
                        data.append("Utilities")
                    elif "COMCAST" in row["Description"]:
                        data[6] = "Your Second Home"
                        data.append("Utilities")
                    elif "BILLFIXERS" in row["Description"]:
                        data[3] = "BillFixers"
                        data.append("Utilities")
                    elif "RED POCKET" in row["Description"]:
                        data.append("Utilities")
                    elif "TELLO" in row["Description"]:
                        data.append("Utilities")
                    elif "LegalShield" in row["Description"]:
                        data[6] = "88 Madison Ave"
                        data.append("Legal & Professional")
                    elif "ROOMONITOR.COM" in row["Description"]:
                        data.append("Subscriptions")

                    # Propane for 110 N Saddle
                    elif "POLAR GAS" in row["Description"]:
                        data[6] = "110 N Saddle Dr"
                        data.append("Utilities")

                    # Trash for 110 N Saddle
                    elif "Doyle" in row["Description"]:
                        data[6] = "110 N Saddle Dr"
                        data.append("Utilities")

                    # Proper.Insure for 110 N Saddle
                    elif "Premium Finance" in row["Description"]:
                        data[6] = "110 N Saddle Dr"
                        data.append("Insurance")

                    # Handy Cleaning
                    elif "HANDY.COM" in row["Description"]:
                        data[6] = "Personal"
                        data.append("Cleaning & Maintenance")

                    # Gas/Fuel
                    elif "CONOCO" in row["Description"]:
                        data[3] = "Conoco"
                        data.append("Automobile > Gas/Fuel")

                    # Gas/Fuel
                    elif "SEI" in row["Description"]:
                        data[3] = "SEI Fuels"
                        data.append("Automobile > Gas/Fuel")

                    elif "Target" in row["Merchant"] or "Walmart" in row["Merchant"] or "Amazon" in row["Merchant"]:
                        data.append("Supplies")

                    elif "Instacart" in row["Merchant"]:
                        data.append("Food & Dining > Groceries")

                    # Aurora Salary
                    elif "GUSTO" in row["Description"]:
                        data[3] = "Aurora Insight"
                        data[6] = "Aurora"
                        data.append("Salary/Wages")

                    # Your Second Home Payroll
                    elif "GUSTO" in row["Description"] and "DESCR:NET" in row["Description"]:
                        data[3] = "Gusto"
                        data[6] = "Your Second Home"
                        data.append("Salary/Wages")

                    # Your Second Home Payroll Taxes
                    elif "GUSTO" in row["Description"] and "DESCR:TAX" in row["Description"]:
                        data[3] = "Gusto"
                        data[6] = "Your Second Home"
                        data.append("Taxes")

                    # Your Second Home Payroll
                    elif "GUSTO" in row["Description"] and "DESCR:FEE" in row["Description"]:
                        data[3] = "Gusto"
                        data[6] = "Your Second Home"
                        data.append("Accounting")

                    # Automatic Payments
                    elif "AUTOPAY" in row["Description"] or "AUTOMATIC PAYMENT" in row["Description"]:
                        data[3] = "Payment"
                        data[6] = "Personal"
                        data.append("Transfer")

                    # Bright Money
                    elif "Bright Money" in row["Description"]:
                        data[3] = "Bright Money"
                        data[6] = "Personal"
                        data.append("Transfer")

                    # Nexo
                    elif "Currency Cloud" in row["Description"]:
                        data[3] = "Nexo"
                        data[6] = "Personal"
                        data.append("Transfer")

                    # Gemini
                    elif "GEMINI TRUST CO" in row["Description"]:
                        data[3] = "Gemini"
                        data[6] = "Crypto"
                        data.append("Investments")

                    elif row["Transfers"]:
                        data.append("Transfer")

                    # Renatus OES
                    elif "TEAM ELEVATE OES" in row["Description"]:
                        data[3] = "Team Elevate"
                        data[6] = "Consulting"
                        data.append("Events")

                    # Nationwide Car Insurance
                    elif "NATIONWIDE" in row["Description"]:
                        data[6] = "Consulting"
                        data.append("Insurance")

                    # Nationwide Pet Insurance
                    elif "NATIONWIDE PET INS" in row["Description"]:
                        data[6] = "Personal"
                        data.append("Insurance")

                    # Zoom
                    elif "ZOOM.US" in row["Description"]:
                        data[3] = "Zoom Video Communications"
                        data[6] = "Consulting"
                        data.append("Subscriptions")

                    # Primerica Life
                    elif "PRIMERICA LIFE" in row["Description"]:
                        data[3] = "Primerica Life"
                        data[6] = "Personal"
                        data.append("Insurance")

                    # Prudential
                    elif "PRUDENTIAL" in row["Description"]:
                        data[3] = "Prudential"
                        data[6] = "Personal"
                        data.append("Insurance")

                    # Northwestern Mutual
                    elif "NORTHWESTERN MUTUAL" in row["Description"]:
                        data[3] = "Northwestern Mutual"
                        data[6] = "Personal"
                        data.append("Insurance")

                    # State Farm Renter's
                    elif "STATE FARM INSURANCE" in row["Description"] and "Chase Freedom Unlimited" in row["Account"]:
                        data[3] = "State Farm"
                        data[6] = "Personal"
                        data.append("Insurance")

                    # Northwestern Mutual
                    elif "LIBERTY MUTUAL" in row["Description"]:
                        data[3] = "Liberty Mutual"
                        data.append("Insurance")

                    # TDAmeritrade
                    elif "TDAmeritrade" in row["Account"]:
                        data[3] = "TDAmeritrade"
                        data[6] = "Personal"
                        data.append("Investments")

                    # Meerkat
                    elif "MEERKAT" in row["Description"]:
                        data[3] = "Meerkat Pest Control"
                        data[6] = "724 3rd Ave"
                        data.append("Cleaning & Maintenance")

                    # Meerkat
                    elif "MEERKAT" in row["Description"]:
                        data[3] = "Meerkat Pest Control"
                        data[6] = "724 3rd Ave"
                        data.append("Cleaning & Maintenance")

                    # Cozy.co Rents
                    elif "Cozy Services" in row["Description"]:
                        data[3] = "Cozy.co"
                        data[6] = "724 3rd Ave"
                        data.append("Rental")

                    # Distributions to Kenny
                    elif "Online Transfer from CHK ...8700" in row["Description"] and "88 Madison" in row["Account"]:
                        data[3] = "Kenny Low"
                        data[6] = "88 Madison Ave"

                    # Distributions to Kenny
                    elif "Online Transfer from CHK ...8700" in row["Description"] and "90 Madison" in row["Account"]:
                        data[3] = "Kenny Low"
                        data[6] = "90 Madison Ave"

                    # Distributions to Earl
                    elif "Online Transfer from CHK ...0000" in row["Description"] and "88 Madison" in row["Account"]:
                        data[3] = "Earl Co"
                        data[6] = "88 Madison Ave"

                    # Distributions to Earl
                    elif "Online Transfer from CHK ...0000" in row["Description"] and "90 Madison" in row["Account"]:
                        data[3] = "Earl Co"
                        data[6] = "90 Madison Ave"

                    # Distributions to Jen
                    elif "Online Transfer from CHK ...8852" in row["Description"] and "88 Madison" in row["Account"]:
                        data[3] = "Jen Yuan"
                        data[6] = "88 Madison Ave"

                    # Distributions to Jen
                    elif "Online Transfer from CHK ...8852" in row["Description"] and "90 Madison" in row["Account"]:
                        data[3] = "Jen Yuan"
                        data[6] = "90 Madison Ave"

                    # Upstart
                    elif row["Amount"] == "-191.83" and "Capital One 360" in row["Account"]:
                        data[3] = "Upstart"
                        data[6] = "Personal"
                        data.append("Transfer")


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
