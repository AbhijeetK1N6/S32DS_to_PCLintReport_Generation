import os
import re
import csv
import pandas as pd
from xml.etree import ElementTree
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from datetime import datetime
import tkinter as tk
from tkinter import filedialog
from tkinterdnd2 import DND_FILES, TkinterDnD
import sys

# Function to clean XML file and save as another XML file
def process_xml_file(input_file, output_file):
    try:
        with open(input_file, 'r') as f:
            lines = f.readlines()

        lines = lines[3:-1]

        filtered_lines = []
        for line in lines:
            if "MISRA 2012" in line and ("required" in line or "mandatory" in line):
                match = re.search(r'<file>(D.*?\.c)</file>', line)
                if match:
                    filtered_lines.append(line.strip())

        filtered_lines = list(set(filtered_lines))

        with open(output_file, 'w') as f:
            f.write('\n'.join(filtered_lines))

        print(f"Processed XML file saved as {output_file}")

    except FileNotFoundError:
        print(f"Error: File '{input_file}' not found.")

def xml_to_csv(input_xml_file, csv_output_file):
    try:
        with open(input_xml_file, "r") as ptr:
            contents = ptr.readlines()

        contents = [line for line in contents if not line.strip().startswith("--- ") and line.strip().startswith("<")]
        contents.insert(0, "<lint>")
        contents.append("</lint>")

        temp_xml_file = "3_vcast_test.xml"
        with open(temp_xml_file, "w") as ptr:
            ptr.writelines(contents)

        tree = ElementTree.parse(temp_xml_file)
        root = tree.getroot()
        result = []

        for item in root.findall("message"):
            line = {}
            for child in item:
                line[child.tag] = child.text
            result.append(line)

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

def adjust_column_width(ws, max_width=50):
    for column in ws.columns:
        max_length = 0
        column_letter = column[0].column_letter
        for cell in column:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(str(cell.value))
            except TypeError:
                pass
        adjusted_width = (max_length + 2) * 1.2
        if adjusted_width > max_width:
            adjusted_width = max_width
        ws.column_dimensions[column_letter].width = adjusted_width

def csv_to_excel(input_csv_file, excel_output_file):
    try:
        df = pd.read_csv(input_csv_file)
        df.drop_duplicates(inplace=True)

        wb = Workbook()
        ws = wb.active

        for r in dataframe_to_rows(df, index=False, header=True):
            ws.append(r)

        adjust_column_width(ws, max_width=111)
        ws.auto_filter.ref = ws.dimensions

        today_date = datetime.now().strftime("%d%m%Y")
        excel_file = os.path.join(os.path.dirname(input_csv_file), f"FinalOutput_{today_date}.xlsx")
        wb.save(excel_file)

        print(f"Excel workbook saved with filters applied and duplicates removed: {excel_file}")

    except FileNotFoundError:
        print(f"Error: File '{input_csv_file}' not found.")

# Function to handle file selection
def select_file():
    input_xml_file = filedialog.askopenfilename(title="Select XML File", filetypes=[("XML files", "*.xml")])
    if input_xml_file:
        process_file(input_xml_file)

# Function to process the dropped file
def drop(event):
    input_xml_file = event.data.strip('{}')  # Clean up the path
    process_file(input_xml_file)

# Function to process the selected or dropped XML file
def process_file(input_xml_file):
    if not input_xml_file.lower().endswith('.xml'):
        print("Error: Input file must be an XML file.")
        return

    cleaned_xml_file = os.path.join(os.path.dirname(input_xml_file), "2_StaticAnalysisCleaned.xml")
    process_xml_file(input_xml_file, cleaned_xml_file)

    csv_output_file = os.path.join(os.path.dirname(input_xml_file), "3_vcast_test.csv")
    xml_to_csv(cleaned_xml_file, csv_output_file)

    csv_to_excel(csv_output_file, csv_output_file.replace('.csv', '.xlsx'))
    print("Thanks for using this tool.")
    print("MISRA Report Generated successfully, and the file has been saved in the 'logs' folder.")
    print("You can close all the windows now.")

# Main execution
if __name__ == "__main__":
    try:
        # Check if the script is run without arguments
        if len(sys.argv) < 2:
            # Set up the TkinterDnD root window
            root = TkinterDnD.Tk()
            root.title("AbhijeetK1N6")

            # Create a button for file selection
            select_button = tk.Button(root, text="Select\nXML File", command=select_file, bg='#4dff00')
            select_button.pack(pady=20)

            # Set up drag-and-drop for the window
            root.drop_target_register(DND_FILES)
            root.dnd_bind('<<Drop>>', drop)

            # Run the Tkinter main loop
            root.geometry("300x100")  # Set window size for better visibility
            
            root.mainloop()
        else:
            # If arguments are provided, process the first argument directly
            input_xml_file = sys.argv[1]
            process_file(input_xml_file)

    except Exception as e:
        print(f"Error: {e}")
