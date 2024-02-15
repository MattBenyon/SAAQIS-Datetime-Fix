import os
import pandas as pd

def make_0_23_hours(date_str):
    # Split the string to extract date and time components
    date_components, time_components = date_str.split()

    # Extract date components
    day, month, year = map(int, date_components.split('/'))

    # Extract time components
    hour, minute = map(int, time_components.split(':'))

    # Get the number of days in the current month
    days_in_month = {
        1: 31,  # January
        2: 28 if year % 4 != 0 or (year % 100 == 0 and year % 400 != 0) else 29,  # February (considering leap year)
        3: 31,  # March
        4: 30,  # April
        5: 31,  # May
        6: 30,  # June
        7: 31,  # July
        8: 31,  # August
        9: 30,  # September
        10: 31,  # October
        11: 30,  # November
        12: 31,  # December
    }

    # Check if the hour is 24
    if hour == 24:
        # Adjust the date to the next day
        day += 1
        if day > days_in_month[month]:
            day = 1
            month += 1
            if month > 12:
                month = 1
                year += 1

        # Adjust hour to 0
        hour = 0

    # Output the adjusted datetime
    adjusted_date_str = "{:02d}/{:02d}/{:04d} {:02d}:{:02d}".format(day, month, year, hour, minute)
    #print(adjusted_date_str)
    
    #adjusted_date_str = pd.to_datetime(adjusted_date_str)
    #print(adjusted_date_str)
    return adjusted_date_str



# folder path
dir_path = r'Data/'

# list to store files
res = []

# Iterate directory
for path in os.listdir(dir_path):
    # check if current path is a file
    if os.path.isfile(os.path.join(dir_path, path)):
        res.append(os.path.join(dir_path, path))
#print(res)



for file in res:
    # Read in the file, skip the first two rows of the excel file, the last 8 (summary stats)
    # tell command that 2 rows are headers (the units are the second one)
    df = pd.read_excel(file,skiprows=2, skipfooter=8, header=[0,1])
    # then join the 2 header rows
    df.columns = df.columns.map(' '.join)
    # rename the datetime column to remove the unnecessary part
    df = df.rename({"Date & Time Unnamed: 0_level_1": "Datetime"},axis=1)
    # apply the function to correct the datetime issue, then make it the right format nd dateitme object
    df["Datetime"] = pd.to_datetime(df["Datetime"].apply(make_0_23_hours), dayfirst=True).dt.strftime('%m/%d/%Y %H:%M')
    # Set the datetime as the index
    df = df.set_index("Datetime")
    
    filename = file.split("/")[1]
    
    df.to_excel(f"Output/{filename}")
                 
