#Code By : K1N6
#Source Code available on GitHub : https://github.com/AbhijeetK1N6/S32DS_to_PCLintReport_Generation/edit/main/MasterScript.py

import sys
import os
import re
import csv
import pandas as pd
from xml.etree import ElementTree
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from datetime import datetime

# Function to clean XML file and save as another XML file
def process_xml_file(input_file, output_file):
    try:
        # Read the entire XML file
        with open(input_file, 'r') as f:
            lines = f.readlines()

        # Remove first three lines (<?xml version="1.0" ?>, <doc>, and last line </doc>)
        lines = lines[3:-1]

        # Filter lines containing both "MISRA 2012" and either "required" or "mandatory"
        filtered_lines = []
        for line in lines:
            if "MISRA 2012" in line and ("required" in line or "mandatory" in line):
                # Extract the file path using regular expression
                match = re.search(r'<file>(D.*?\.c)</file>', line)
                if match:
                    filtered_lines.append(line.strip())

        # Remove duplicates
        filtered_lines = list(set(filtered_lines))

        # Write cleaned lines to the output file
        with open(output_file, 'w') as f:
            f.write('\n'.join(filtered_lines))

        print(f"Processed XML file saved as {output_file}")

    except FileNotFoundError:
        print(f"Error: File '{input_file}' not found.")

# Function to convert cleaned XML to CSV
def xml_to_csv(input_xml_file, csv_output_file):
    try:
        # Read contents of input_xml_file
        with open(input_xml_file, "r") as ptr:
            contents = ptr.readlines()

        # Modify contents to prepare XML structure
        contents = [line for line in contents if not line.strip().startswith("--- ") and line.strip().startswith("<")]
        contents.insert(0, "<lint>")
        contents.append("</lint>")

        # Write modified contents to a temporary XML file
        temp_xml_file = "3_vcast_test.xml"
        with open(temp_xml_file, "w") as ptr:
            ptr.writelines(contents)

        # Parse the temporary XML file
        tree = ElementTree.parse(temp_xml_file)
        root = tree.getroot()
        result = []

        # Extract data from XML into result list
        for item in root.findall("message"):
            line = {}
            for child in item:
                line[child.tag] = child.text
            result.append(line)

        # Define CSV output file and write the result
        fields = ['file', 'line', 'code', 'desc', 'type']
        with open(csv_output_file, "w", encoding='utf-8', newline='') as ptr:
            writer = csv.DictWriter(ptr, fieldnames=fields)
            writer.writeheader()
            writer.writerows(result)

        print(f"CSV file '{csv_output_file}' generated successfully.")

    except FileNotFoundError:
        print(f"Error: File '{input_xml_file}' not found.")
    except Exception as e:
        print(f"Error: {e}")

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

# Function to convert CSV to Excel
def csv_to_excel(input_csv_file, excel_output_file):
    try:
        # Read the CSV file into a pandas DataFrame
        df = pd.read_csv(input_csv_file)

        # Remove duplicate rows
        df.drop_duplicates(inplace=True)

        # Create a new Excel workbook and add the DataFrame to it
        wb = Workbook()
        ws = wb.active

        for r in dataframe_to_rows(df, index=False, header=True):
            ws.append(r)

        # Adjust column widths with a maximum width of 111
        adjust_column_width(ws, max_width=111)

        # Apply filters to all columns (excluding header)
        ws.auto_filter.ref = ws.dimensions

        # Generate today's date in DDMMYYYY format
        today_date = datetime.now().strftime("%d%m%Y")

        # Construct the output file name with today's date
        excel_file = os.path.join(os.path.dirname(input_csv_file), f"FinalOutput_{today_date}.xlsx")

        # Save the Excel workbook
        wb.save(excel_file)

        print(f"Excel workbook saved with filters applied and duplicates removed: {excel_file}")

    except FileNotFoundError:
        print(f"Error: File '{input_csv_file}' not found.")

#Code By : K1N6
#Source Code available on GitHub : https://github.com/AbhijeetK1N6/S32DS_to_PCLintReport_Generation/edit/main/MasterScript.py

# Main execution
if __name__ == "__main__":
    try:

        if len(sys.argv) < 2:
            print("Error: Please provide the XML file path as an argument.")
            sys.exit(1)
        
        input_xml_file = sys.argv[1]
        
        if not input_xml_file.lower().endswith('.xml'):
            print("Error: Input file must be an XML file.")
            sys.exit(1)

        # Step 1: XML Cleaning
        #input_xml_file = r'D:\\SOFTWARE\\PC-lintPlus_V2.0\\pclp.windows.2.0\\pclp\\logs\\1_StaticAnalysis.xml'
        cleaned_xml_file = os.path.join(os.path.dirname(input_xml_file), "2_StaticAnalysisCleaned.xml")
        process_xml_file(input_xml_file, cleaned_xml_file)
        print("Step 1: XML Cleaning completed successfully.")

        # Step 2: XML to CSV Conversion
        input_cleaned_xml = cleaned_xml_file
        csv_output_file = os.path.join(os.path.dirname(input_xml_file), "3_vcast_test.csv")
        xml_to_csv(input_cleaned_xml, csv_output_file)
        print("Step 2: XML to CSV Conversion completed successfully.")

        # Step 3: CSV to Excel Conversion
        input_csv_file = csv_output_file
        csv_to_excel(input_csv_file, input_csv_file.replace('.csv', '.xlsx'))
        print("Step 3: CSV to Excel Conversion completed successfully.")

        print("Automation process completed successfully.")

    except Exception as e:
        print(f"Error: {e}")
        
#Code By : K1N6
#Source Code available on GitHub : https://github.com/AbhijeetK1N6/S32DS_to_PCLintReport_Generation/edit/main/MasterScript.py
