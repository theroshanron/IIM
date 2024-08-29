import pandas as pd
from icalendar import Calendar, Event
from datetime import datetime, timedelta

# Function to parse time
def parse_time(date_obj, time_str):
    start_time_str, end_time_str = time_str.split(" to ")
    start_datetime_str = f"{date_obj.strftime('%Y-%m-%d')} {start_time_str}"
    start_datetime = datetime.strptime(start_datetime_str, "%Y-%m-%d %I:%M %p")
    
    end_datetime_str = f"{date_obj.strftime('%Y-%m-%d')} {end_time_str}"
    end_datetime = datetime.strptime(end_datetime_str, "%Y-%m-%d %I:%M %p")

    return start_datetime, end_datetime

# Load the schedule from an Excel file
excel_file_path = r'C:\Users\ABC\Downloads\Q3_Schedule.xlsx'  # Update this with the path to your Excel file
df = pd.read_excel(excel_file_path)

# Create a calendar
cal = Calendar()
cal.add('prodid', '-//Your Product//Your Calendar//EN')
cal.add('version', '2.0')

for index, row in df.iterrows():
    event = Event()
    course_name = row["Subject"]
    code = row["Code"]
    faculty_name = row["Faculty Name"]
    event_name = f"{course_name} - {code} - {faculty_name}"
    date = row["Date"]
    time = row["Time"]
    description = "\n".join([f"{col}: {row[col]}" for col in df.columns if col not in ["Date", "Time"]])
    
    start_datetime, end_datetime = parse_time(date, time)
    event.add('summary', event_name)
    event.add('dtstart', start_datetime)
    event.add('dtend', end_datetime)
    event.add('description', description)
    event.add('dtstamp', datetime.now())

    cal.add_component(event)

# Write to ics file
with open('academic_schedule.ics', 'wb') as f:
    f.write(cal.to_ical())

print("ICS file created successfully.")
