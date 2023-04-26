# import openpyxl to allow us to pick a specific sheet within the Excel workbook
import openpyxl

# Load the Excel file SALM Smoothed LGA Datafiles (ASGS 2022) - June quarter 2022 from LGA folder into a pandas dataframe
workbook = openpyxl.load_workbook('/Users/hanaarshadahmed/Desktop/LGA/Unemployment/SALM Smoothed LGA Datafiles (ASGS 2022) - June quarter 2022.xlsx')

# Deletes Smoothed LGA unemployment and Smoothed LGA labour force sheets from the Excel workbook
del workbook['Smoothed LGA unemployment']
del workbook['Smoothed LGA labour force']

# Deletes specific rows (i.e. 1,2,3) from the Smoothed LGA unemployment sheet
sheet = workbook['Smoothed LGA unemployment rates']
rows_to_delete = [1, 2, 3]
for row_index in sorted(rows_to_delete, reverse=True):
    sheet.delete_rows(row_index)

# Deletes columns from the Smoothed LGA unemployment sheet
sheet.delete_cols(2, 22)

# Save the modified dataframe back to a new Excel file called Final_Unemployment_2022 in the LGA folder
workbook.save('/Users/hanaarshadahmed/Desktop/LGA/Unemployment/Final_Unemployment_2022')
