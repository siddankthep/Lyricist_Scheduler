import datetime
import os.path
import pandas as pd

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

from utils import (
    get_work_days,
    flatten_schedule,
    create_event,
    get_sheet_url,
    get_all_calendar_ids,
)

# If modifying these scopes, delete the file token.json.
SCOPES = ["https://www.googleapis.com/auth/calendar"]


def main():
    """Shows basic usage of the Google Calendar API."""
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
            flow = InstalledAppFlow.from_client_secrets_file("credentials.json", SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open("token.json", "w") as token:
            token.write(creds.to_json())

    current_month = datetime.datetime.now().strftime("%B")

    # This is the link to the Google Sheet template, insert your own link here
    sheet_url = "https://docs.google.com/spreadsheets/d/1gAeJxdwRQoP3Ol5_eqLnmpUMG7Un-dIMLpv8_XI72oE/edit?usp=sharing"

    sheet_link_csv = get_sheet_url(sheet_url)

    sheet = pd.read_excel(
        sheet_link_csv, sheet_name=f"{current_month} Schedule", header=1
    )
    month_schedule = flatten_schedule(sheet)

    try:
        service = build("calendar", "v3", credentials=creds)

        all_calendars = get_all_calendar_ids(service)

        # Change this to the calendar ID you want to add the events to
        calendar_id = "primary"

        for work_day in get_work_days(month_schedule):
            create_event(work_day, service, calendar_id)

    except HttpError as error:
        print(f"An error occurred: {error}")


if __name__ == "__main__":
    main()
