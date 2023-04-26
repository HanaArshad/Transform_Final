# import openpyxl to allow us to pick a specific sheet within the Excel workbook
import openpyxl

# Open the Excel file called NSW_trends in LGA folder
workbook = openpyxl.load_workbook('/Users/hanaarshadahmed/Desktop/LGA/Trends/NSW_trends.xlsx')

# Deletes specific rows in the Excel sheet
sheet = workbook['New South Wales']
rows_to_delete = [1, 2, 3, 8126, 8127, 8128, 8129, 8130, 8131, 8132, 8133]
for row_index in sorted(rows_to_delete, reverse=True):
    sheet.delete_rows(row_index)
sheet.delete_cols(8, 3)

# Save the updates in Final_Trendd in the LGA folder
workbook.save('/Users/hanaarshadahmed/Desktop/LGA/Trends/Final_Trendd')

# imports pandas allows as to change values within the Excel File
import pandas as pd

# Reads the Final_trend file in the LGA Folder
location = pd.ExcelFile('/Users/hanaarshadahmed/Desktop/LGA/Trends/Final_Trendd')
df = pd.read_excel(location)

# Remove betting and gaming from the values within the Excel sheet
df = df.loc[~(df['Offence type'] == 'Betting and gaming offences')]

# Saves the final modifications in the "Final_Trend2_" Excel file in the LGA folder
output = '/Users/hanaarshadahmed/Desktop/LGA/Trends/Final_Trend2_.xlsx'
df.to_excel(output, index=False)
