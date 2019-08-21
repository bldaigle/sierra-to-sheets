import gspread
import time
import psycopg2
import sys
sys.path.append('/home/sierra/config') # refer to the sierra_config folder for credentials
import sierra_config # imports database and email credentials

from oauth2client.service_account import ServiceAccountCredentials

from pprint import pprint
from googleapiclient import discovery

scope = [
    "https://spreadsheets.google.com/feeds",
    "https://www.googleapis.com/auth/drive",
    "https://www.googleapis.com/auth/drive.file",
    "https://www.googleapis.com/auth/spreadsheets"
]

# Get credentials for the Google Sheets API
creds = ServiceAccountCredentials.from_json_keyfile_name('/home/sierra/scripts/open_orders/creds.json', scope)

# Create an empty Google Sheet
sheet = discovery.build('sheets', 'v4', credentials=creds) # Instantiate a new sheet
data = {'properties': {'title': 'Open Orders at Kenyon [%s]' % time.ctime()}} # Specify content for the new sheet
res = sheet.spreadsheets().create(body=data).execute() # Load the content
sheet_id = res['spreadsheetId'] # Get the id of the new sheet
print('Created "%s"' % res['properties']['title']) # Print out a confirmation of the new sheet

# Get data from the Sierra database
fields = ('Order Record Number', 'Bib Record Number', 'Title', 'Vendor', 'Created Date', 'Updated Date')
connection = psycopg2.connect("dbname=" + sierra_config.sierra_dbname + " user=" + sierra_config.sierra_dbuser + " host=" + sierra_config.sierra_dbhost + " port=" + sierra_config.sierra_dbport + " password=" + sierra_config.sierra_dbpass + " sslmode=" + sierra_config.sierra_dbssl)
cursor = connection.cursor()
cursor.execute(open("/home/sierra/scripts/open_orders/kenyon_open_orders.sql", "r").read())
rows = cursor.fetchall()
connection.close()
rows.insert(0, fields)
sierra_data = {'values': [row[:6] for row in rows]}

# Populate new sheet with data from Sierra database
sheet.spreadsheets().values().update(spreadsheetId=sheet_id, range='A1', body=sierra_data, valueInputOption='RAW').execute()
print('Wrote data to Sheet:')
rows = sheet.spreadsheets().values().get(spreadsheetId=sheet_id, range='Sheet1').execute().get('values', [])
for row in rows:
    print(row)

# Move new sheet to a shared folder
drive_service = discovery.build('drive', 'v3', credentials=creds)
folder_id = '1LPHaG147zoha1s3BsMcC0-epuQEdwuZm'
## Retrieve existing parent of the file (My Drive)
file = drive_service.files().get(fileId=sheet_id, fields='parents').execute()
previous_parents = ", ".join(file.get('parents'))
## Move file to the appropriate folder
file = drive_service.files().update(fileId=sheet_id, supportsAllDrives='true', addParents=folder_id, removeParents=previous_parents, fields='id, parents').execute()


