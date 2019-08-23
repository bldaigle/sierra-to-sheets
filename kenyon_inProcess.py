# /usr/bin/env python3
# Author: Ben Daigle
# Date: August 23, 2019

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
creds = ServiceAccountCredentials.from_json_keyfile_name('/home/sierra/scripts/in_process/creds.json', scope)

# Create an empty Google Sheet
sheet = discovery.build('sheets', 'v4', credentials=creds) # Instantiate a new sheet
data = {'properties': {'title': 'Items In Process at Kenyon [%s]' % ohio_now}} # Give new sheet a title
res = sheet.spreadsheets().create(body=data).execute() # Add title to the new sheet
sheet_id = res['spreadsheetId'] # Get the id of the new sheet
print('Created "%s"' % res['properties']['title']) # Print out a confirmation of the new sheet

# Get data from the Sierra database
fields = ('Item Record Number', 'Title', 'Barcode', 'Call Number', 'Created Date', 'Updated Date') # Specify the fields (and also the headers) for the report
connection = psycopg2.connect("dbname=" + sierra_config.sierra_dbname + " user=" + sierra_config.sierra_dbuser + " host=" + sierra_config.sierra_dbhost + " port=" + sierra_config.sierra_dbport + " password=" + sierra_config.sierra_dbpass + " sslmode=" + sierra_config.sierra_dbssl)
cursor = connection.cursor() 
cursor.execute(open("/home/sierra/scripts/in_process/kenyon_in_process.sql", "r").read())
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
                    "endIndex": 1
                },
                "properties": {
                    "pixelSize": 150
                },
                "fields": "pixelSize"
            }   
        },
        {
            "updateDimensionProperties": { # Set the first two column widths to 150px
                "range": {
                    "sheetId": 0,
                    "dimension": "COLUMNS",
                    "startIndex": 2,
                    "endIndex": 6
                },
                "properties": {
                    "pixelSize": 150
                },
                "fields": "pixelSize"
            }   
        },
        {
            "updateDimensionProperties": { # Set the second column width to 350px
                "range": {
                    "sheetId": 0,
                    "dimension": "COLUMNS",
                    "startIndex": 1,
                    "endIndex": 2
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
folder_id = '1Jw1NXa6l11d85VFYn7Rx4Aq_PZnczgQK' # Specify the ID of the shared Google drive
file = drive_service.files().get(fileId=sheet_id, fields='parents').execute() # Retrieve existing parent of the file (My Drive)
previous_parents = ", ".join(file.get('parents'))
file = drive_service.files().update(fileId=sheet_id, supportsAllDrives='true', addParents=folder_id, removeParents=previous_parents, fields='id, parents').execute() # Move file to the appropriate folder

# Configure email header information
emailhost = sierra_config.email_host
emailuser = sierra_config.email_user
emailpass = sierra_config.email_pass
emailport = sierra_config.email_port
emailsubject = 'Kenyon Items In Process Report'
emailmessage = '''
    <html>
        <head></head>
        <body>
        <div style="width: 100%!important;margin:0;padding:0;">
            <div style="padding:10px;line-height:18px;font-family:'Lucida Grande',Verdana,Arial,sans-serif;font-size:12px;color:#444444;">
                <div style="margin-top:25px;">
                    <table width="100%" cellpadding="0" cellspacing="0" border="0" role="presentation">
                        <tbody>
                            <tr>
                                <td colspan="2" width="100%" style="padding:15px 0;border-top:1px dotted #c5c5c5">
                                    <tbody>
                                        <tr>
                                        
                                            <td valign="top" style="padding:0 15px 0 15px;width:40px">
                                                <img width="40" height="40" alt="Five Colleges of Ohio Logo" style="height:auto;line-height:100%;outline:none;text-decoration:none;border-radius:5px" src="https://www.ohio5.org/sites/default/files/oh5-logo-sm_1.jpg">
                                            </td>
                                            <td width="100%" style="padding:0;margin:0;" valign="top">
                                                <p style="font-family:'Lucida Grande','lucida Sans Unicode','Lucida Sans',Verdana,Tahoma,sans-serif;font-size:15px;line-height:18px;margin-bottom:0;margin-top:0;padding:0;color:#1b1d1e">
                                                    <strong>CONSORT Library Admin</strong>
                                                </p>
                                                <p>''' + ohio_now + '''</p>
                                                <div style="color:#2b2e2f;font-family:'Lucida Sans Unicode','Lucida Grande','Tahoma',Verdana,sans-serif;font-size:14px;line-height:22px;margin:15px 0;">
                                                    <p>Hello there!</p>
                                                    <p>A report of Kenyon items with the status "In Process" has been posted to the <a href="https://drive.google.com/drive/u/2/folders/1Jw1NXa6l11d85VFYn7Rx4Aq_PZnczgQK">Kenyon College Reports</a> shared drive. Have a look and let us know if you have any questions about the report.</p>
                                                    <p>Thanks!</p>
                                                    <p>Ben Daigle & Matt McNemar</p>
                                                </div
                                            </td>
                                        </tr>
                                    </tbody>
                                </td>
                        </tbody>
                    </table>
                </div>
            </div>
        </div>
        </body>
    </html>
'''
emailfrom = sierra_config.email_from
emailto = ['email_recipient.1@sample.edu', 'email_recipient.2@sample.edu']

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
