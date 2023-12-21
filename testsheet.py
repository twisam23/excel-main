import gspread
from oauth2client.service_account import ServiceAccountCredentials

# Set scope and credentials from the credentials.json file obtained from Google Developer Console
scope = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive'
]

creds = ServiceAccountCredentials.from_json_keyfile_name('excelapp-main\e-store-408305-2d24bf10c72e.json', scope)
client = gspread.authorize(creds)

# Open the spreadsheet by its title
spreadsheet_title = 'testdata'  # Replace with your spreadsheet title
sheet = client.open(spreadsheet_title).sheet1  # Use sheet by index or name, here using the default 'Sheet1'

# Access data from the sheet
data = sheet.get_all_records()

# Print the data
print(data)
