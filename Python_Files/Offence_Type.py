# This code is automated, which mean this code will read each Excel file within the LGA_Preprocessed folder and apply
# the same code in each file because we are removing the same thing from every file

# Imports path which allows as to automate the process
from pathlib import Path

# Identifies the path in which the Excel files are saved in this case its saved in LGA_Preprocessed
input_dir = Path.cwd() / '/Users/hanaarshadahmed/Desktop/LGA/Offence_Type'

# import openpyxl to allow us to pick a specific sheet within the Excel workbook and load the workbook
from openpyxl import load_workbook  # pip install openpyxl

# For every Excel file in VV folder perform the following functions
for path in list(input_dir.rglob("*.xlsx*")):
    # Loads the Excel file
    wb = load_workbook(filename=path)
    # Deletes sheets named; Victims, Premises type, Summary of offences, and alcohol related
    del wb["Victims"]
    del wb["Premises Type"]
    del wb['Summary of offences']
    del wb['Alcohol Related']

    # Code to delete specific items in a sheet
    # Changes in Sheet 'Offenders'
    sheet = wb['Offenders']
    # Deletes specific rows in Offenders sheet in the Excel file
    rows_to_delete = [1, 2, 3, 4, 36, 37, 38, 39, 40, 41, 42, 43, 44, 45, 46]
    for row_index in sorted(rows_to_delete, reverse=True):
        sheet.delete_rows(row_index)

    # Changes in Sheet 'Aboriginality'
    sheet = wb['Aboriginality']
    # Deletes specific rows in Aborginality sheet in the Excel file
    rows_to_delete = [1, 5, 11, 12, 13, 14, 15, 16, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31, 32, 33, 34,
                      35, 36, 37]
    for row_index in sorted(rows_to_delete, reverse=True):
        sheet.delete_rows(row_index)

    # Changes in Sheet 'Month'
    sheet = wb['Month']
    # Deletes specific rows in Month sheet in the Excel file
    rows_to_delete = [1, 2, 3, 4, 41, 42, 43, 44, 45, 46, 47]
    for row_index in sorted(rows_to_delete, reverse=True):
        sheet.delete_rows(row_index)

    # Changes in Sheet 'Time'
    sheet = wb['Time']
    # Deletes specific Time in Offenders sheet in the Excel file
    rows_to_delete = [1, 2, 3, 4, 41, 42, 43, 44, 45, 46, 47]
    for row_index in sorted(rows_to_delete, reverse=True):
        sheet.delete_rows(row_index)

    # Save the modified data into a new folder called LGA_Final
    output_dir = Path.cwd() / '/Users/hanaarshadahmed/Desktop/LGA/Offence_Type_Final'
    output_dir.mkdir(exist_ok=True)
    wb.save(output_dir / path.name)
