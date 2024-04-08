import pandas as pd
import datetime


class WorkDay:
    def __init__(self, staff, day, date, shift, location):
        self.staff = staff
        self.day = day
        self.date = date
        self.start, self.end = self.get_start_end_time(shift)
        self.shift = shift
        self.location = location

    def get_start_end_time(self, shift: str):
        start_time_str, end_time_str = shift.split(" - ")
        start_time = datetime.datetime.strptime(start_time_str, "%H:%M")
        end_time = datetime.datetime.strptime(end_time_str, "%H:%M")

        # Convert the date from 'DD/MM' format to a date object
        # Assuming the year is this year. Replace with the correct year if needed.
        year = datetime.date.today().year
        date = datetime.datetime.strptime(f"{self.date}/{year}", "%d/%m/%Y").date()

        # Combine the date with the start and end times
        start_datetime = datetime.datetime.combine(date, start_time.time())
        end_datetime = datetime.datetime.combine(date, end_time.time())

        return start_datetime.isoformat(), end_datetime.isoformat()

    def __str__(self):
        return f"{self.staff} is working on {self.day}, {self.date} from {self.start} to {self.end} at {self.location}"

    def __repr__(self):
        return f"{self.staff} is working on {self.day}, {self.date} from {self.start} to {self.end} at {self.location}"


def flatten_schedule(month_schedule: pd.DataFrame):
    new_headers = ["SHIFT"] + [
        "MONDAY",
        "MONDAY",
        "TUESDAY",
        "TUESDAY",
        "WEDNESDAY",
        "WEDNESDAY",
        "THURSDAY",
        "THURSDAY",
        "FRIDAY",
        "FRIDAY",
        "SATURDAY",
        "SATURDAY",
        "SUNDAY",
        "SUNDAY",
    ] * 5

    month_schedule = month_schedule.drop(month_schedule.columns[-1], axis=1)
    weeks = []
    num_row, _ = month_schedule.shape
    for i in range(7, num_row, 6):
        week = pd.concat(
            [month_schedule.iloc[0:1], month_schedule.iloc[i : i + 6]],
            axis=0,
            ignore_index=True,
        ).iloc[0:, 1:]
        weeks.append(week)

    flat = month_schedule.copy()
    for week in weeks:
        flat = pd.concat([flat, week], axis=1)
    flat = flat.set_axis(new_headers, axis=1)

    flat.iloc[1:2] = flat.iloc[1:2].ffill(axis=1)

    return flat.iloc[0:7]


def get_work_days(month: pd.DataFrame, staff_name: str = None) -> "list[WorkDay]":
    work_days = []

    num_row, num_col = month.shape
    for row in range(2, num_row):
        for col in range(1, num_col):
            if isinstance(month.iloc[row, col], float):
                month.iloc[row, col] = str(month.iloc[row, col])
            if month.iloc[row, col] != "X" and month.iloc[row, col] != "nan":
                if "\n" in month.iloc[row, col]:
                    name, shift = month.iloc[row, col].split("\n")
                    shift = shift[1:-1]
                else:
                    name = month.iloc[row, col]
                    shift = month.iloc[row, 0]

                if staff_name and staff_name != name:
                    continue

                day = month.columns[col].lower().capitalize()
                date = month.iloc[1, col]
                location = month.iloc[0, col]

                work_day = WorkDay(name, day, date, shift, location)
                work_days.append(work_day)

                # print(f"{name} is working on {day}, {date} during {shift} at {location}")
    return work_days


def create_event(work_day: WorkDay, service, calendar_id, attendees=False):
    created_events_description = []
    created_events = service.events().list(calendarId=calendar_id).execute()

    for event in created_events["items"]:
        created_events_description.append(event["description"])

    if work_day.__str__() not in created_events_description:
        if attendees:
            event = {
                "summary": work_day.staff,
                "location": work_day.location,
                "description": work_day.__str__(),
                "start": {
                    "dateTime": work_day.start,
                    "timeZone": "Asia/Bangkok",
                },
                "end": {
                    "dateTime": work_day.end,
                    "timeZone": "Asia/Bangkok",
                },
                "attendees": [
                    {"email": "anhthy7102003@gmail.com"},
                ],
                "reminders": {
                    "useDefault": False,
                    "overrides": [
                        {"method": "email", "minutes": 3 * 60},
                        {"method": "popup", "minutes": 45},
                    ],
                },
            }

        else:
            event = {
                "summary": work_day.staff,
                # "id": work_day.id,
                "location": work_day.location,
                "description": work_day.__str__(),
                "start": {
                    "dateTime": work_day.start,
                    "timeZone": "Asia/Bangkok",
                },
                "end": {
                    "dateTime": work_day.end,
                    "timeZone": "Asia/Bangkok",
                },
            }

        event = (
            service.events()
            .insert(
                calendarId=calendar_id,
                body=event,
            )
            .execute()
        )
        print("Event created: %s" % (event.get("htmlLink")))
    else:
        print(
            f"Event for {work_day.staff} on {work_day.day}, {work_day.date} already exists"
        )
