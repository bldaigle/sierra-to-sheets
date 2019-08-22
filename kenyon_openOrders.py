import gspread # suite of tools for working with Google Sheets
import psycopg2 # for connecting to the Sierra PostgreSQL database
import pytz # for handling timezones
import sys # for importing other files in the file system
sys.path.append('/home/sierra/config') # makes a local config file accessible to the program
import sierra_config # imports database and email credentials
from datetime import datetime # for working with dates and times
from oauth2client.service_account import ServiceAccountCredentials # for authentication with the Google Sheets API
from googleapiclient import discovery # required by Google Drive API
import smtplib # for sending emails
from email.mime.multipart import MIMEMultipart # for constructing emails appropriately
from email.mime.base import MIMEBase # for constructing emails appropriately
from email.mime.text import MIMEText # for constructing emails appropriately
from email.utils import formatdate # for constructing emails appropriately
from email import encoders # for encoding email messages

# Establish some variables for dates and times in the Eastern Timezone
tz = pytz.timezone('America/New_York')
ohio_now = datetime.now(tz).strftime("%a %b %d, %Y %I:%M%p")

# Sets the Google API scopes required for the program
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
data = {'properties': {'title': 'Open Orders at Kenyon [%s]' % ohio_now}} # Give new sheet a title
res = sheet.spreadsheets().create(body=data).execute() # Add title to the new sheet
sheet_id = res['spreadsheetId'] # Get the id of the new sheet
print('Created "%s"' % res['properties']['title']) # Print out a confirmation of the new sheet

# Get data from the Sierra database
fields = ('Order Record Number', 'Bib Record Number', 'Title', 'Vendor', 'Created Date', 'Updated Date') # Specify the fields (and also the headers) for the report
connection = psycopg2.connect("dbname=" + sierra_config.sierra_dbname + " user=" + sierra_config.sierra_dbuser + " host=" + sierra_config.sierra_dbhost + " port=" + sierra_config.sierra_dbport + " password=" + sierra_config.sierra_dbpass + " sslmode=" + sierra_config.sierra_dbssl)
cursor = connection.cursor() 
cursor.execute(open("/home/sierra/scripts/open_orders/kenyon_open_orders.sql", "r").read())
rows = cursor.fetchall()
connection.close()
rows.insert(0, fields) # Insert the fields/headers into the first row of the Google Sheet
sierra_data = {'values': [row[:6] for row in rows]}

# Populate new sheet with data from Sierra database
sheet.spreadsheets().values().update(spreadsheetId=sheet_id, range='A1', body=sierra_data, valueInputOption='RAW').execute() # Gathers the data
print('Wrote data to Sheet')
rows = sheet.spreadsheets().values().get(spreadsheetId=sheet_id, range='Sheet1').execute().get('values', [])
for row in rows:
    print(row)

# Format the sheet
format_reqs = { "requests": [
        {
            "updateSheetProperties": { # Freeze the first row in the sheet
                "properties": {
                    "gridProperties": {
                        "frozenRowCount": 1
                    },
                },
                "fields": "gridProperties.frozenRowCount"   
            }
        },
        {
            "updateDimensionProperties": { # Set the first two column widths to 150px
                "range": {
                    "sheetId": 0,
                    "dimension": "COLUMNS",
                    "startIndex": 0,
                    "endIndex": 2
                },
                "properties": {
                    "pixelSize": 150
                },
                "fields": "pixelSize"
            }   
        },
        {
            "updateDimensionProperties": { # Set the third column width to 350px
                "range": {
                    "sheetId": 0,
                    "dimension": "COLUMNS",
                    "startIndex": 2,
                    "endIndex": 3
                },
                "properties": {
                    "pixelSize": 350
                },
                "fields": "pixelSize"
            }   
        },
        {"repeatCell": { # Bold the header row
            "range": {
                "sheetId": 0,
                "startRowIndex": 0,
                "endRowIndex": 1
            },
            "cell": {
                "userEnteredFormat": {
                    "textFormat": {
                        "bold": True
                    }   
                }   
            },
            "fields": "userEnteredFormat.textFormat.bold"
            }
        }
    ]}
sheet.spreadsheets().batchUpdate(spreadsheetId=sheet_id, body=format_reqs).execute() # Send formatting requirements to the API


# Move new sheet from private drive to a shared folder
drive_service = discovery.build('drive', 'v3', credentials=creds)
folder_id = '1LPHaG147zoha1s3BsMcC0-epuQEdwuZm' # Specify the ID of the shared Google drive
file = drive_service.files().get(fileId=sheet_id, fields='parents').execute() # Retrieve existing parent of the file (My Drive)
previous_parents = ", ".join(file.get('parents'))
file = drive_service.files().update(fileId=sheet_id, supportsAllDrives='true', addParents=folder_id, removeParents=previous_parents, fields='id, parents').execute() # Move file to the appropriate folder

# Configure email header information
emailhost = sierra_config.email_host
emailuser = sierra_config.email_user
emailpass = sierra_config.email_pass
emailport = sierra_config.email_port
emailsubject = 'Kenyon Open Orders Report'
emailmessage = '''
    <html>
        <head></head>
        <body>
            <p>Hello!</p>
            <p>A report of current open library orders for Kenyon has been posted to the <a href="https://drive.google.com/drive/u/2/folders/1LPHaG147zoha1s3BsMcC0-epuQEdwuZm">Kenyon Open Orders</a> shared drive. Have a look and let us know if you have any questions about the report.</p>
            <p>Thanks!</p>
        </body>
    </html>
'''
emailfrom = sierra_config.email_from
emailto = ['email_recipient.1@sample.edu','email_recipient2@sample.edu']

# Create the email message
msg = MIMEMultipart('alternative')
msg['From'] = emailfrom
if type(emailto) is list:
    msg['To'] = ', '.join(emailto)
else:
    msg['To'] = emailto
msg['Date'] = formatdate(localtime = True)
msg['Subject'] = emailsubject
msg.attach(MIMEText(emailmessage, 'html'))

#Send the email
smtp = smtplib.SMTP(emailhost, emailport)
smtp.ehlo() #for Google connection
smtp.starttls()
smtp.login(emailuser, emailpass)
smtp.sendmail(emailfrom, emailto, msg.as_string())
smtp.quit()
