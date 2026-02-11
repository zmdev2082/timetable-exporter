from datetime import datetime, timedelta
from icalendar import Calendar, Event
import uuid
import logging
from zoneinfo import ZoneInfo

import pandas as pd

class IcalGenerator:
    REQUIRED_FIELDS = ['summary', 'location']
    VALID_FIELDS = ['summary', 'location', 'description', 'dtstart', 'dtend', 'duration', 'category', 'attendee', 'organizer', 'url']
    def __init__(self, columns, timezone: str = 'Australia/Sydney'):

        self.columns = columns
        self.timezone = ZoneInfo(timezone)
        # Check if all required fields are present in the config
        for field in self.REQUIRED_FIELDS:
            if field not in self.columns:
                raise ValueError(f"Missing required field in config: {field}")

        # check if start and end or date_start, time_start and duration are present
        if not (("dtstart" in self.columns and "dtend" in self.columns) or
            ("dtstart" in self.columns and "duration" in self.columns)):
            raise ValueError("Event time fields need to be configured as: (dtstart & dtend) or (dtstart & duration)")


    def add_event_property(self, event, key, value, entry):

        if key in ['dtstart', 'dtend']:
            if value is None or (isinstance(value, float) and pd.isna(value)):
                return

            # Normalize strings to datetime
            if isinstance(value, str):
                value = datetime.strptime(value, '%Y-%m-%d %H:%M:%S')

            # Accept pandas Timestamp
            if isinstance(value, pd.Timestamp):
                value = value.to_pydatetime()

            # Ensure timezone-aware
            if isinstance(value, datetime) and value.tzinfo is None:
                value = value.replace(tzinfo=self.timezone)
        
        elif key == 'duration':
            if isinstance(value, str):
                try:
                    td = pd.to_timedelta(value)
                    value = timedelta(seconds=int(td.total_seconds()))
                except Exception:
                    parts = value.strip().split(":")
                    if len(parts) == 2:
                        hours, minutes = parts
                        seconds = "0"
                    elif len(parts) == 3:
                        hours, minutes, seconds = parts
                    else:
                        raise ValueError(f"Unsupported duration format: {value}")
                    value = timedelta(hours=int(hours), minutes=int(minutes), seconds=int(seconds))
            elif isinstance(value, (int, float)):
                value = timedelta(hours=value)
        
        elif key not in self.VALID_FIELDS:
            return

        event.add(key.upper(), value)

    def generate_ical(self, timetable_data: list[dict[str, str]], company: str="TSSAMME") -> Calendar:
        cal = Calendar()
        cal.add('prodid', '-//'+ company + '//timetable-exporter//EN')
        cal.add('version', '2.0')

        for entry in timetable_data:
            event = Event()
            event.add('UID', str(uuid.uuid4()))  # Add a unique identifier for each event
            event.add('DTSTAMP', datetime.now())
            for key, column in self.columns.items():
                value = entry.get(column)
                self.add_event_property(event, key, value, entry)
            cal.add_component(event)
        
        return cal
    