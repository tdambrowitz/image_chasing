import pandas as pd
import streamlit as st
import re
from io import BytesIO

# Function to process data
def process_data(uploaded_file):
    df = pd.read_csv(
        uploaded_file,
        skiprows=2,
        usecols=range(69),  # Adjust range according to the structure of your CSV file.
        nrows=6000,
    )
    df = df.dropna(subset=['Job Number'])
    df = df[df[df.columns[0]].apply(lambda x: re.match(r'^[A-Za-z][0-9]', str(x)) is not None)]
    df = df[~df["Key Tag"].isna()]

    # Format date columns here with pandas datetime functions if needed
    # Example: df['Due On Site Date/Time'] = pd.to_datetime(df['Due On Site Date/Time']).dt.strftime('%Y-%m-%d %H:%M')
    
    # Fill NaN values with 'N/A' for cleaner visuals
    df.fillna('N/A', inplace=True)

    # Capitalize or title case specific string columns
    # Example: df['Customer Name'] = df['Customer Name'].str.title()

    extracted_data = df[['Job Number', 'Location', 'Due On Site Date/Time', 'Customer Name', 'Vehicle Registration', 'Key Tag', 'Driveable', 'Insurer', "Insured's Post Code", 'Vehicle Manufacturer', 'Vehicle Model', 'Entered Date/Time', 'Last Customer Contact Date/Time']].copy()
    extracted_data = extracted_data.sort_values(by=['Key Tag', 'Due On Site Date/Time'])
    extracted_data.reset_index(drop=True, inplace=True)
    return extracted_data

def to_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Sheet1')

        workbook = writer.book
        worksheet = writer.sheets['Sheet1']

        # Define formats with a border and without
        bordered_format = workbook.add_format({
            'border': 1,
            'text_wrap': False,
            'valign': 'top'})
        borderless_format = workbook.add_format({
            'text_wrap': False,
            'valign': 'top'})

        # Define header format with a background color
        header_format = workbook.add_format({
            'bold': True,
            'text_wrap': False,
            'valign': 'top',
            'fg_color': '#D7E4BC',  # Use your header color
            'border': 2})

        # Set the format for alternating grey and white rows
        grey_fmt = workbook.add_format({'bg_color': '#f0f0f0', 'text_wrap': False, 'valign': 'top', 'border': 1,})
        white_fmt = workbook.add_format({'bg_color': '#ffffff', 'text_wrap': False, 'valign': 'top', 'border': 1,})

        # Set the column widths based on header lengths
        for i, header in enumerate(df.columns):
            column_len = len(header) + 2  # Adjust the column length here as necessary
            worksheet.set_column(i, i, column_len)

        # Write the header with the header format
        for col_num, header in enumerate(df.columns):
            worksheet.write(0, col_num, header, header_format)

        # Write data rows with job number borders
        for row_index, row in enumerate(df.values, start=1):
            for col_num, cell_value in enumerate(row):
                fmt = grey_fmt if (row_index) % 2 == 0 else white_fmt  # Choose format based on even/odd row

                worksheet.write(row_index, col_num, cell_value, fmt)

    output.seek(0)
    return output

# Streamlit interface setup
st.title('CSV File Processor for Collision Repair Data')


with st.expander("How do I run the report?"):
    st.write('1. Navigate to the "Job Listing" report (found under "Job Analysis" in BMS)')
    st.write('2. Set the "Date" selector to "Scheduled Onsite"')
    st.write('3. Set the "From Date" selector to the current Date')
    st.write('4. Set the "To Date" selector to as far out as you want to go (e.g. 4 weeks)')
    st.write('5. Click "Print" then close the excel file that opens (you can save it somewhere if you want)')
    st.write('6. Come back to this page and upload the file (it should start with "job_list" and end with ".csv")')


uploaded_file = st.file_uploader("Upload your CSV file", type=['csv'])

# When a file is uploaded, process and display the data
if uploaded_file is not None:
    # Process file
    processed_data = process_data(uploaded_file)
    
    # Show processed data
    st.write("Processed Data:")
    st.dataframe(processed_data.head())

    # Download button for processed data as Excel
    st.download_button(
        label="Download Excel file",
        data=to_excel(processed_data),
        file_name="processed_data.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
