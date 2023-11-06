from pathlib import Path
import pandas as pd
import os
import re
import plotly.express as px
from datetime import datetime
from openpyxl import load_workbook  # To get sheet names



def find_most_recent_csv_file(folder_path):
    csv_files = [f for f in os.listdir(folder_path) if f.startswith("job_list") and f.endswith(".csv")]
    
    # If there are no matching csv files, return None
    if not csv_files:
        return None
    
    # Sort the files based on modification time in descending order
    most_recent_csv = max(csv_files, key=lambda f: os.path.getmtime(os.path.join(folder_path, f)))
    return os.path.join(folder_path, most_recent_csv)

# Use the above function to get the path to the most recent CSV file
current_folder = os.path.dirname(os.path.abspath(__file__))
csv_file_path = find_most_recent_csv_file(current_folder)


if not csv_file_path:
    print("No matching CSV file found!")
else:
    # Load data directly from the CSV file
    def get_data_from_csv():
        df = pd.read_csv(
            csv_file_path,
            skiprows=2,
            usecols=range(69),  # equivalent to 'A:AE' in Excel
            nrows=6000,
        )
        return df

    # Load data using the cached function
    df = get_data_from_csv()




# Drop rows with empty values in the 'Scheduled OnSite' column
df = df.dropna(subset=['Job Number'])


# Remove rows where the first cell doesn't start with a letter followed by a number
df = df[df[df.columns[0]].apply(lambda x: re.match(r'^[A-Za-z][0-9]', str(x)) is not None)]

# Exclude rows with "HWARR" or "HALO" in the "Debtor" column
#df = df[~df["Source Code"].str.match("HALO Warranty", case=False, na=False)]
df = df[~df["Key Tag"].isna()]
extracted_data = df[['Job Number', 'Location', 'Due On Site Date/Time', 'Customer Name', 'Vehicle Registration', 'Key Tag', 'Driveable', 'Insurer', "Insured's Post Code", 'Vehicle Manufacturer', 'Vehicle Model', 'Entered Date/Time', 'Last Customer Contact Date/Time']].copy()

# Sort the extracted_data dataframe by the 'Due On Site Date/Time' column in ascending order and then by the 'Key Tag' column in ascending order
extracted_data = extracted_data.sort_values(by=['Key Tag', 'Due On Site Date/Time'])

# Reset index for cleaner output
extracted_data.reset_index(drop=True, inplace=True)

print(extracted_data.head())

    
extracted_data.to_csv('output_file_path.csv', index=False)