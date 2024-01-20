import pandas as pd
from datetime import datetime, timedelta

def convert_to_timedelta(time_str):
    # Convert float to string and then to timedelta
    try:
        hours, minutes = map(int, time_str.split(':'))
        return timedelta(hours=hours, minutes=minutes)
    except (ValueError, AttributeError):
        return pd.NaT  # Return NaT for invalid time values
    
def analyze_excel_file(file_path):
    # Load the Excel file into a DataFrame
    df = pd.read_csv(file_path)

    # Drop rows with missing 'Time In' or 'Time Out'
    df = df.dropna(subset=['Time', 'Time Out'])

    # Convert 'Time In' and 'Time Out' columns to datetime format
    df['Time In'] = pd.to_datetime(df['Time'], errors='coerce')
    df['Time Out'] = pd.to_datetime(df['Time Out'], errors='coerce')

    # Convert 'Timecard Hours' to timedelta format
    df['Timecard Hours (as Time)'] = df['Timecard Hours (as Time)'].astype(str).apply(convert_to_timedelta)

    # Dictionary to store results
    results = {
        '7_consecutive_days': set(),
        'less_than_10_hours_but_more_than_1_hour_between_shifts': set(),
        'more_than_14_hours_single_shift': set()
    }

    # Iterate through each employee
    for employee, group in df.groupby('Employee Name'):
        consecutive_days_count = 1
        previous_date = None

        # Calculate time_between_shifts for the entire group
        time_between_shifts = group['Time In'] - group['Time Out'].shift(fill_value=pd.to_datetime('now'))

        # Iterate through each row of the employee's time entries
        for index, row in group.iterrows():
            # Check for 7 consecutive days
            current_date = row['Time In'].date()
            if previous_date is not None and (current_date - previous_date).days == 0:
              consecutive_days_count += 0
            elif previous_date is not None and (current_date - previous_date).days == 1:
                consecutive_days_count += 1
            else:
                consecutive_days_count = 1

            if consecutive_days_count == 7:
                results['7_consecutive_days'].add((employee, row['Position ID']))

            # Check for less than 10 hours between shifts and greater than 1 hour
            if any(1 < hours < 10 for hours in time_between_shifts.dt.total_seconds() / 3600):
                results['less_than_10_hours_but_more_than_1_hour_between_shifts'].add((employee, row['Position ID']))

            # Check for more than 14 hours in a single shift
            shift_hours = row['Timecard Hours (as Time)'].total_seconds() // 3600
            if shift_hours > 14:
                results['more_than_14_hours_single_shift'].add((employee, row['Position ID']))

            previous_date = current_date

    # Print the results
    with open('output.txt', 'w') as f:
        for key, value in results.items():
            f.write(f'\nEmployees with {key}:\n')
            for item in value:
                f.write(f"Employee: {item[0]}, Position ID: {item[1]}\n")

file_path = 'Assignment_Timecard.csv'

# Call the function with the file path
analyze_excel_file(file_path)
