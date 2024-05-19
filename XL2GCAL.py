import pandas as pd
import os.path
from google.oauth2 import service_account
from googleapiclient.discovery import build
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request



# Define the scope and load the credentials
SCOPES = ['https://www.googleapis.com/auth/calendar']
SERVICE_ACCOUNT_FILE = '/path/to/credentials.json'

creds = None
  # The file token.json stores the user's access and refresh tokens, and is
  # created automatically when the authorization flow completes for the first
  # time.
if os.path.exists("token.json"):
    creds = Credentials.from_authorized_user_file("token.json", SCOPES)
  # If there are no (valid) credentials available, let the user log in.
if not creds or not creds.valid:
    if creds and creds.expired and creds.refresh_token:
      creds.refresh(Request())
    else:
      flow = InstalledAppFlow.from_client_secrets_file(
          SERVICE_ACCOUNT_FILE, SCOPES
      )
      creds = flow.run_local_server(port=0)
    # Save the credentials for the next run
    with open("token.json", "w") as token:
      token.write(creds.to_json())



# Build the service
service = build('calendar', 'v3', credentials=creds)

# Read events from the Excel file
excel_file = '/path/to/excelfile.xlsx'
df = pd.read_excel(excel_file)

# Ensure 'Start Date' and 'End Date' are in datetime format
df['Start Date'] = pd.to_datetime(df['Start Date'])
df['End Date'] = pd.to_datetime(df['End Date'])

# Combine the date and time columns
df['Start DateTime'] = df.apply(lambda row: pd.to_datetime(f"{row['Start Date'].date()} {row['Start Time']}"), axis=1)
df['End DateTime'] = df.apply(lambda row: pd.to_datetime(f"{row['End Date'].date()} {row['End Time']}"), axis=1)

# Function to create an event
def create_event(service, calendar_id, event):
    event = service.events().insert(calendarId=calendar_id, body=event).execute()
    print(f'Event created: {event.get("htmlLink")}')

# Calendar ID (you can use 'primary' for the primary calendar)
calendar_id = 'calendar_id under configuration options in google calendar'

# Iterate over the rows in the DataFrame and create events
for index, row in df.iterrows():
    event = {
        'summary': row['Subject'],
        'description': row.get('Description', ''),
        'start': {
            'dateTime': row['Start DateTime'].isoformat(),
            'timeZone': 'America/New_York',  # Change this if your events are in a different time zone
        },
        'end': {
            'dateTime': row['End DateTime'].isoformat(),
            'timeZone': 'America/New_York',  # Change this if your events are in a different time zone
        },
    }
    create_event(service, calendar_id, event)

print('All events have been added to the calendar.')
