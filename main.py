import pandas as pd
from datetime import datetime
from pathlib import Path
import uuid

# Settings
NAME = "Avigail"

INPUT_FILE = Path('shifts.xlsx')
OUTPUT_FILE = Path('shifts.ics')
df = pd.read_excel(INPUT_FILE, header=None)

shifts = []
shifts_offset = [[1, 2], [3,4]]
for row_index, row in df.iloc[0::5].iterrows():
    for column_index, date in row.items():
        for shift in shifts_offset:
            date = df.iloc[row_index, column_index]
            name = df.iloc[row_index + shift[0], column_index]
            if pd.isnull(name):
                continue
            time = df.iloc[row_index + shift[1], column_index].replace(" ", "")
            start_time = datetime.strptime(f"{date.date()} {time.split('-')[0]}", "%Y-%m-%d %H:%M")
            end_time = datetime.strptime(f"{date.date()} {time.split('-')[1]}", "%Y-%m-%d %H:%M")
            
            print(f"{name=}, {time=} {start_time=} {end_time=}")
            if name == NAME:
                shifts.append([start_time, end_time])

print(f"Found {len(shifts)} for {NAME}")

with OUTPUT_FILE.open("w") as ics:
    ics.write("""BEGIN:VCALENDAR
VERSION:2.0
PRODID:-//ical.marudot.com//iCal Event Maker
CALSCALE:GREGORIAN
BEGIN:VTIMEZONE
TZID:Asia/Jerusalem
LAST-MODIFIED:20231222T233358Z
TZURL:https://www.tzurl.org/zoneinfo-outlook/Asia/Jerusalem
X-LIC-LOCATION:Asia/Jerusalem
BEGIN:DAYLIGHT
TZNAME:IDT
TZOFFSETFROM:+0200
TZOFFSETTO:+0300
DTSTART:19700327T020000
RRULE:FREQ=YEARLY;BYMONTH=3;BYDAY=-1FR
END:DAYLIGHT
BEGIN:STANDARD
TZNAME:IST
TZOFFSETFROM:+0300
TZOFFSETTO:+0200
DTSTART:19701025T020000
RRULE:FREQ=YEARLY;BYMONTH=10;BYDAY=-1SU
END:STANDARD
END:VTIMEZONE
""")
    
    for shift in shifts:
        start, end = shift
        ics.write("BEGIN:VEVENT\n")
        ics.write(f"DTSTAMP:{datetime.now().strftime(r'%Y%m%dT%H%M%SZ')}\n")
        ics.write(f"UID:{uuid.uuid4()}\n")
        ics.write(f"DTSTART;TZID=Asia/Jerusalem:{start.strftime(r'%Y%m%dT%H%M%S')}\n")
        ics.write(f"DTEND;TZID=Asia/Jerusalem:{end.strftime(r'%Y%m%dT%H%M%S')}\n")
        ics.write("SUMMARY:Shift\n")
        ics.write("CLASS:PRIVATE\n")
        ics.write("BEGIN:VALARM\n")
        ics.write("ACTION:DISPLAY\n")
        ics.write("DESCRIPTION:Shift\n")
        ics.write("TRIGGER:-P1D\n")
        ics.write("END:VALARM\n")
        ics.write("END:VEVENT\n")
    
    ics.write("END:VCALENDAR")