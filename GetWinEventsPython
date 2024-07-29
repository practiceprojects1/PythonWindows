#GetWinEventsPython
import win32evtlog  #Import the pywin32 library
import pandas as pd
from datetime import datetime, timedelta

# Configuration
log_name = "Security"  # Change to the desired log name, e.g., "System" or "Security"
csv_file_path = r"C:\Users\94eke\OneDrive\Desktop\EventsLastMinute.csv"

def fetch_events(log_name, start_time, end_time):
    events_list = []
    
    try:
        # Open the event log
        log_handle = win32evtlog.OpenEventLog(None, log_name)
        
        # Read the events from the log
        events = win32evtlog.ReadEventLog(log_handle, win32evtlog.EVENTLOG_SEQUENTIAL_READ | win32evtlog.EVENTLOG_FORWARDS_READ, 0)
        
        for event in events:
            event_time = event.TimeGenerated.Format()
            event_time = datetime.strptime(event_time, "%Y-%m-%dT%H:%M:%S.%fZ")  # Convert to datetime object
            
            if start_time <= event_time <= end_time:
                events_list.append({
                    "TimeCreated": event_time,
                    "EventID": event.EventID,
                    "EventType": event.EventType,
                    "Message": event.StringInserts
                })
                
    except Exception as e:
        print(f"Failed to fetch events: {e}")
    
    return events_list

def main():
    # Get the current time and calculate the time 1 minute ago
    end_time = datetime.now()
    start_time = end_time - timedelta(minutes=1)
    
    # Fetch events from the last minute
    events = fetch_events(log_name, start_time, end_time)
    
    # Convert to DataFrame and save to CSV
    if events:
        df = pd.DataFrame(events)
        df.to_csv(csv_file_path, index=False)
        print(f"Successfully exported events to {csv_file_path}")
    else:
        print("No events found for the specified time range.")

if __name__ == "__main__":
    main()
