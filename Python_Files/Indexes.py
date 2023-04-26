# Imports pandas allows as to change values within the Excel File
import pandas as pd

# Load the Excel file LGA_Indexes from file Indexes into a pandas dataframe
sheets = pd.read_excel('/Users/hanaarshadahmed/Desktop/LGA/Indexes /LGA_Indexes.xls', sheet_name='Table 1')

# Removes specific rows (i.e 0,1,2) from the Table 1 sheet
rows_to_remove = [0, 1, 2, 3, 4, 551]
sheets = sheets.drop(rows_to_remove)

# Removing specific columns (i.e 0,3) from the Table 1 sheet
var = sheets.columns
sheets = sheets.drop(sheets.columns[[0, 3, 5, 7, 9, 10]], axis=1)

# Renaming each column in the sheet to its appropriate name
sheets = sheets.rename(columns={'Australian Bureau of Statistics ': '2016 Local Government Area (LGA) Code'})
sheets = sheets.rename(columns={'Unnamed: 1': '2016 Local Government Area (LGA) Name'})
sheets = sheets.rename(columns={'Unnamed: 2': 'Index of Relative Socio-economic Disadvantage'})
sheets = sheets.rename(columns={'Unnamed: 3': 'Decile'})
sheets = sheets.rename(columns={'Unnamed: 4': 'Index of Relative Socio-economic Advantage and Disadvantage'})
sheets = sheets.rename(columns={'Unnamed: 5': 'Decile'})
sheets = sheets.rename(columns={'Unnamed: 6': 'Index of Economic Resources'})
sheets = sheets.rename(columns={'Unnamed: 7': 'Decile'})
sheets = sheets.rename(columns={'Unnamed: 8': 'Index of Education and Occupation'})
sheets = sheets.rename(columns={'Unnamed: 9': 'Decile'})
sheets = sheets.rename(columns={'Unnamed: 10': 'Usual Resident Population'})

# Save the modified dataframe back to a new Excel file called Python_Files.xlsx in the Final_Indexes folder
sheets.to_excel('/Users/hanaarshadahmed/Desktop/LGA/Indexes /Final_Indexes.xlsx', index=False)
