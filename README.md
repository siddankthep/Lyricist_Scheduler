# Lyricist_Scheduler

Automatically imports a schedule from an Excel sheet and create events on Google Calendar using the Google Calendar API.

# Setup

## Python

### Version

This works with Python 3.12.2

### Libraries

```
pip install -r requirements.txt
```

## Spreadsheet

You can either use the spreadsheet locally or use a spreadsheet on Google Sheets. The spreadsheet should have the same format as the [Spreadsheet template](https://docs.google.com/spreadsheets/d/1gAeJxdwRQoP3Ol5_eqLnmpUMG7Un-dIMLpv8_XI72oE/edit?usp=sharing).

The overall format of the spreadsheet should follow these rules:

1. The Excel file can have multiple sheets, each sheet representing a month. Each sheet should be named `[Month] Schedule`.
2. The columns represent the day and date, and the rows represent the time of each shift. A staff's name should be inserted to the cell corresponding to their shift.
3. If a staff is not working the full shift, a note should be included under the staff's name, e.g `Your Name (11:00 - 16:00)` (The note should be on a new line in the cell).

## Google Workspace

Follow the [Python Quickstart](https://developers.google.com/calendar/api/quickstart/python) guide from Google to set up your Google Workspace account and enable the Google Calendar API.

The Google client library is already included in `requirements.txt` so you can skip that part.

**After downloading the JSON file and saving it as `credentials.json`, you are ready to insert your schedule to Google Calendar.**

# Usage

Make sure that you have already had your `credentials.json` file in the directory.

In [`schedule.py`](/schedule.py), change the following fields:

- `sheet_url`: the url to your Google Spreadsheet
- `calendar_id`: the id of the calendar that you would like to create events on. You can access the id by using the `all_calendar` dictionary, e.g `calendar_id = all_calendars["Your calendar name"]`

After, simply run the python file. You will be prompted to authorize using your Google Account. Happy scheduling!

# TODO

- [] Update the flatten_schedule function for flexible spreadsheets (different shifts)
