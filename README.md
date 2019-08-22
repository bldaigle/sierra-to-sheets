# kenyon_openOrders
An integration between the Google Sheets API and the Sierra library system that automates the creation and delivery of open library orders.

## Project Components
This project consists of the following structural components:
* Automation server (just an AWS micro-instance for hosting scripts)
* Sierra PostgreSQL Database
* Google Sheets API
* Google Shared Drive

### Automation Server
All files are hosted on a t2.micro instance in Amazon Web Services (AWS). These include:
* `sierra_config.txt` - A restricted file that contains sensitive login information. Permissions on this file are set to 0400
* `kenyon_open_orders.sql` - A PostgreSQL query for extracting data from the Sierra database
* `kenyon_openOrders.py` - A Python program that fetches data from the Sierra database and posts it to a shared Google Sheet
* `creds.json` - A restricted file that contains Google API credentials
A cron job executes the Python program every week on Monday at 8:00 AM ET

### Sierra PostgreSQL Database
The Sierra product offers direct access to its underlying PostgreSQL database. The `kenyon_open_orders.sql` file contains the query used to extract data about open library orders. The tables and database UML diagrams are documented but are proprietary and therefore behind an authentication wall for customers only.

### Google Sheets API
This project uses Google Sheets API v4 which is well documented at on the [Google Sheets API Reference](https://developers.google.com/sheets/api/).

### Google Shared Drive
A Google Shared Drive (formerly Team Drive) serves as the destination for the resulting reports. When a new report is posted to the shared drive, an email notification is sent to any recipients specified in the Python program.
