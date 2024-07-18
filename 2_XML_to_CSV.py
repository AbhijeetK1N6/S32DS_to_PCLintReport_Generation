#
# Converts XML Cleaned by 1st Step to CSV File and Remove left over duplicates
#

import os
import sys
import csv
from xml.etree import ElementTree

def main():
    input_xml_file = r"D:\SOFTWARE\PC-lintPlus_V2.0\pclp.windows.2.0\pclp\logs\2_StaticAnalysisCleaned.xml"

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
        csv_output_file = "3_vcast_test.csv"
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

if __name__ == "__main__":
    main()


