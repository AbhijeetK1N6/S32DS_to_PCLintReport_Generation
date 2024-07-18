"""
import pandas as pd
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from datetime import datetime
import sys
import os

# Function to adjust column widths based on content length, with a maximum width constraint
def adjust_column_width(ws, max_width=50):
    for column in ws.columns:
        max_length = 0
        column_letter = column[0].column_letter  # Get the column letter
        for cell in column:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(str(cell.value))
            except TypeError:
                pass
        adjusted_width = (max_length + 2) * 1.2  # Adjusted width with a little padding
        # Apply maximum width constraint
        if adjusted_width > max_width:
            adjusted_width = max_width
        ws.column_dimensions[column_letter].width = adjusted_width

# Function to process CSV and save as Excel
def process_csv_to_excel(csv_file, output_folder):
    try:
        # Step 1: Read the CSV file into a pandas DataFrame
        df = pd.read_csv(csv_file)

        # Step 2: Remove duplicate rows
        df.drop_duplicates(inplace=True)

        # Step 3: Create a new Excel workbook and add the DataFrame to it
        wb = Workbook()
        ws = wb.active

        for r in dataframe_to_rows(df, index=False, header=True):
            ws.append(r)

        # Step 4: Adjust column widths with a maximum width of 50 (adjust as needed)
        adjust_column_width(ws, max_width=111)

        # Step 5: Apply filters to all columns (excluding header)
        ws.auto_filter.ref = ws.dimensions

        # Step 6: Generate today's date in DDMMYYYY format
        today_date = datetime.now().strftime("%d%m%Y")

        # Step 7: Construct the output file name with today's date
        excel_file = os.path.join(output_folder, f"FinalOutput_{today_date}.xlsx")

        # Step 8: Save the Excel workbook
        wb.save(excel_file)

        print(f"Excel workbook saved with filters applied and duplicates removed: {excel_file}")

    except FileNotFoundError:
        print(f"Error: File '{csv_file}' not found.")

# Example usage:
if __name__ == "__main__":
    if len(sys.argv) != 2:
        print("Usage: python script.py <path_to_csv_file>")
        sys.exit(1)

    input_csv_path = sys.argv[1]
    if not input_csv_path.lower().endswith('.csv'):
        print("Error: Input file must be a CSV file.")
        sys.exit(1)

    output_folder = os.path.dirname(input_csv_path)  # Output folder is the same as the input folder

    # Process CSV to Excel
    process_csv_to_excel(input_csv_path, output_folder)
"""

import pandas as pd
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from datetime import datetime
import os

# Function to adjust column widths based on content length, with a maximum width constraint
def adjust_column_width(ws, max_width=50):
    for column in ws.columns:
        max_length = 0
        column_letter = column[0].column_letter  # Get the column letter
        for cell in column:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(str(cell.value))
            except TypeError:
                pass
        adjusted_width = (max_length + 2) * 1.2  # Adjusted width with a little padding
        # Apply maximum width constraint
        if adjusted_width > max_width:
            adjusted_width = max_width
        ws.column_dimensions[column_letter].width = adjusted_width

# Function to process CSV and save as Excel
def process_csv_to_excel(csv_file):
    try:
        # Step 1: Read the CSV file into a pandas DataFrame
        df = pd.read_csv(csv_file)

        # Step 2: Remove duplicate rows
        df.drop_duplicates(inplace=True)

        # Step 3: Create a new Excel workbook and add the DataFrame to it
        wb = Workbook()
        ws = wb.active

        for r in dataframe_to_rows(df, index=False, header=True):
            ws.append(r)

        # Step 4: Adjust column widths with a maximum width of 111 (adjust as needed)
        adjust_column_width(ws, max_width=111)

        # Step 5: Apply filters to all columns (excluding header)
        ws.auto_filter.ref = ws.dimensions

        # Step 6: Generate today's date in DDMMYYYY format
        today_date = datetime.now().strftime("%d%m%Y")

        # Step 7: Construct the output file name with today's date
        output_folder = os.path.dirname(csv_file)
        excel_file = os.path.join(output_folder, f"FinalOutput_{today_date}.xlsx")

        # Step 8: Save the Excel workbook
        wb.save(excel_file)

        print(f"Excel workbook saved with filters applied and duplicates removed: {excel_file}")

    except FileNotFoundError:
        print(f"Error: File '{csv_file}' not found.")

# Example usage:
if __name__ == "__main__":
    input_csv_path = r"D:\SOFTWARE\PC-lintPlus_V2.0\pclp.windows.2.0\pclp\logs\3_vcast_test.csv"

    if not input_csv_path.lower().endswith('.csv'):
        print("Error: Input file must be a CSV file.")
        sys.exit(1)

    # Process CSV to Excel
    process_csv_to_excel(input_csv_path)


#
#1) Converts CSV to Excel Workbook
#2) Adjusts Column Widths (MAX 111)
#3) Removes Duplicate Rows
#4) Apply Filters to all columns
#5) Constructs the output file name in FileName_DDMMYYYY format
#6) Saves the Excel Workbook by self