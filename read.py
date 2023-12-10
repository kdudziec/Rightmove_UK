from googleapiclient.discovery import build
from google.oauth2 import service_account


SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
SERVICE_ACCOUNT_FILE = 'keys.json'     # Downloaded keys from google cloud as 'python-projects-341918-f77861f2f4f6.json'

# creds = None
creds = service_account.Credentials.from_service_account_file(
        SERVICE_ACCOUNT_FILE, scopes=SCOPES)
SAMPLE_SPREADSHEET_ID = '15ZdwC4U83OvcObbrOYJ_FL2xLKTEYerWwmS-hHAvzHo' # Comes from my spreadsheet url; 'https://docs.google.com/spreadsheets/d/15ZdwC4U83OvcObbrOYJ_FL2xLKTEYerWwmS-hHAvzHo/edit#gid=2036120036'
service = build('sheets', 'v4', credentials=creds)

# Call the Sheets API
sheet = service.spreadsheets()
result = sheet.values().get(spreadsheetId=SAMPLE_SPREADSHEET_ID,
                                    range="Properties!A2:D18").execute()
values = result.get('values', []) # Return values, if not return empty list

# aoa = [["02/01/2022 19:35:26", "ChesHill on Mission | 33 Seneca Ave, San Francisco", "£2,995", "https://www.zillow.com/b/cheshill-on-mission-san-francisco-ca-BKTtQD/"]]
# request = sheet.values().update(spreadsheetId=SAMPLE_SPREADSHEET_ID,
#                                 range="Properties!A2", valueInputOption="USER_ENTERED", body={"values":aoa}).execute()
#


# request = service.spreadsheets().values().update(spreadsheetId=spreadsheet_id, range=range_, valueInputOption=value_input_option, body=value_range_body)
# request = service.spreadsheets().values().append(spreadsheetId=spreadsheet_id, range=range_, valueInputOption=value_input_option, insertDataOption="INSERT_ROWS", body=value_range_body)

aoa = [["02/01/2022 19:35:26", "ChesHill on Mission | 33 Seneca Ave, San Francisco", "£2,995", "https://www.zillow.com/b/cheshill-on-mission-san-francisco-ca-BKTtQD/"]]
request = sheet.values().append(spreadsheetId=SAMPLE_SPREADSHEET_ID,
                                range="Properties!A2", valueInputOption="USER_ENTERED", insertDataOption="INSERT_ROWS", body={"values":aoa}).execute()
print(request)

