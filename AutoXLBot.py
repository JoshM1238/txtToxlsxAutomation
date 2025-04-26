"""
Reads values from 'values.txt' and writes them into 'output.xlsx',
one per row. Runs every set interval and auto-creates the input file
if missing. Handles locked Excel file errors gracefully.
"""

import time
from openpyxl import Workbook
import os

# File names
txt_file = "values.txt"
output_file = "output.xlsx"

# Interval between runs in minutes
interval_minutes = 1

# Converts string values to int or float if possible
def parse_value(item):
    try:
        if "." in item:  # Check if number has a decimal
            return float(item)
        else:
            return int(item)
    except ValueError:
        return item  # Leave as string if not a number

# Main function to read text data and write it to an Excel file
def run_script():
    # Check if the input text file exists, create if not
    if not os.path.exists(txt_file):
        print(f"'{txt_file}' not found. Creating a new blank file...")
        open(txt_file, "w").close()
        print(f"Blank file '{txt_file}' successfully created.")
        return

    # Read lines from the text file
    with open(txt_file, "r") as f:
        raw_lines = [line.strip() for line in f.readlines()]

    # Flatten all values from all lines
    all_values = []
    for line in raw_lines:
        all_values.extend(line.split())

    # Create a new Excel workbook and select the active sheet
    wb = Workbook()
    ws = wb.active
    ws.title = "Imported Data"

    # Write each value into a new row in column A
    for i, value in enumerate(all_values, start=1):
        ws.cell(row=i, column=1, value=parse_value(value))

    # Try saving the workbook, retry if Excel has the file open
    while True:
        try:
            wb.save(output_file)
            break
        except PermissionError:
            print(f"Cannot write to '{output_file}' — please close the file and press any key to continue")
            os.system("pause")

    print(f"Data written to '{output_file}' successfully.")


# Loop the script forever, running at the defined interval
while True:
    run_script()
    print(f"\nWaiting {interval_minutes} minutes before next run...")
    time.sleep(interval_minutes * 60)

