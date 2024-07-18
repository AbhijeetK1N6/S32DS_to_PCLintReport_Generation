#
# This Script Cleans the XML File:
#1) Removes 1st, 2nd and last line (Doc Tag Opening / Closing)
#2) Remove Duplicate Lines
#3) Removes Blank Lines
#4) Removes Non-imp lines (Files other than D directory / Files not ending on .c)
#5) Keeps only MISRA 2012 warnings (Required / Mandatory)


import os
import re

def process_xml_file(input_file, output_file):
    try:
        # Read the entire XML file
        with open(input_file, 'r') as f:
            lines = f.readlines()

        # Remove first three lines (<?xml version="1.0" ?>, <doc>, and last line </doc>)
        lines = lines[3:-1]  # Removing first 3 and last line

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

# Define input XML file path here
input_xml_file = r'D:\SOFTWARE\PC-lintPlus_V2.0\pclp.windows.2.0\pclp\logs\1_StaticAnalysis.xml'

# Output file path with fixed name "2_StaticAnalysisCleaned.xml"
output_xml_file = os.path.join(os.path.dirname(input_xml_file), "2_StaticAnalysisCleaned.xml")

# Call process_xml_file with the defined input and output file paths
process_xml_file(input_xml_file, output_xml_file)

