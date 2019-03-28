# MoneyWiz-to-Google-Sheet-Exporter
This simple script takes a .CSV file of transactions exported from <a href="https://wiz.money">MoneyWiz 2 or 3</a> and automatically appends the transactions to your own copy of this custom <a href="https://docs.google.com/spreadsheets/d/10PmGsjxMXvIMDIig1QiS-YVYxqOClZEvEu8B9Z69MeA/copy"><b>Google Sheet</b></a> I made for business bookkeeping purposes. No coding experience necessary. You just need to know how to follow directions below.


## Prerequisites for running the app on your desktop
From https://developers.google.com/sheets/api/quickstart/python

To run this program, you'll need:

Python 3.4 or greater.

This was developed and tested using Python 3.6.4 on macOS High Sierra and tested in Python 3.7.3 on Windows 10 
but it should work in any other Python environment as long as the necessary Python modules are installed.
If Python complains that one of the modules is missing, just install that module and let me know that this
documentation is missing a package.

The <a href="https://pypi.python.org/pypi/pip">pip</a> package management tool.

A Google account.

Clone this repository using git in Terminal (Mac or Linux) by typing:
```
git clone https://github.com/earlvanze/MoneyWiz-to-Google-Sheet-Exporter moneysheets && cd moneysheets
```
Windows users need to install git or GitHub Desktop or simply download the repository as a zip and unzip the package.

## Step 1: Turn on the Google Sheets API
Use <a href="https://console.developers.google.com/start/api?id=sheets.googleapis.com">this wizard</a> to create or
select a project in the Google Developers Console and automatically turn on the API.

Click Continue, then Go to credentials.

On the Add credentials to your project page, click the Cancel button.

At the top of the page, select the OAuth consent screen tab.

Select an Email address, enter a Product name if not already set, and click the Save button.

Select the Credentials tab, click the Create credentials button and select OAuth client ID.

Select the application type Other, enter the name "Google Sheets API Quickstart", and click the Create button.

Click OK to dismiss the resulting dialog.

Click the file_download (Download JSON) button to the right of the client ID.

Move this file to your working directory and rename it client_secret.json.


## Step 2: Install the Google Client Library and other non-standard python modules
Run the following command in Terminal to install the necessary libraries using pip:
```
pip install --upgrade googleapiclient
```
See the library's <a href="https://developers.google.com/api-client-library/python/start/installation">installation page</a> for the alternative installation options.

## Step 3: Make a copy of spreadsheet and update transactions.py
Make your own copy of the Google spreadsheet linked above and replace GSHEET_ID
with your own Google Sheet's ID (derived from the URL https://docs.google.com/spreadsheets/d/{SPREADSHEET_ID})

If you renamed the specific sheet 'Transactions!A:I' inside the Google Sheets file or modified the number of columns,
replace RANGE_NAME with the new name along with the new range of columns.

## Step 4: Run the program
Run the sample using the following command
```
python transactions.py
```
An Open File dialog box will come up (on Windows or Linux using Tkinter module). Select the correct CSV file of transactions you want to copy to Google Sheets.

The append_to_gsheet() code will attempt to open a new window or tab in your default browser. If this fails, copy the URL from the console and manually open it in your browser.

If you are not already logged into your Google account, you will be prompted to log in.
If you are logged into multiple Google accounts, you will be asked to select one account to use for the authorization.

Click the Accept button.
The sample will proceed automatically, and you may close the window/tab.

If you need help, type:
```
python main.py --help
```
