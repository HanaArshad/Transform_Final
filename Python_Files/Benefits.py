# imports pandas allows as to change values within the Excel File
import pandas as pd

# Load the Excel file into a pandas dataframe
file_path = '/Users/hanaarshadahmed/Desktop/LGA/Benefits/dss-payments-2020-lga-jun-2021-to-dec-2022-historic.csv'
# Reads the Excel file
df = pd.read_csv(file_path)
# Removes specific columns by its name and displays wanted columns
df = df.drop(
    ['ABSTUDY (Living allowance)', 'Commonwealth Seniors Health Card', 'Date', 'ABSTUDY (Non-living allowance)',
     'Austudy', 'LGA_code_2020', 'Age Pension', 'Age Pension', 'Carer Allowance',
     'Carer Allowance (Child Health Care Card only)', 'Carer Payment', 'Disability Support Pension',
     'Family Tax Benefit A', 'Family Tax Benefit B', 'Parenting Payment Partnered', 'Parenting Payment Single',
     'Pension Concession Card', 'Special Benefit', 'Commonwealth Rent Assistance', 'Partner Allowance',
     'Widow Allowance'], axis=1)

# Adding the values in column Youth Allowance (student and apprentice) and Youth Allowance (other) and saving this in
# a new column called Total Youth allowance
df["Total Youth Allowance"] = df["Youth Allowance (other)"] + df["Youth Allowance (student and apprentice)"]

# Removes Youth Allowance (student and apprentice) and Youth Allowance (other) after calculating the total youth
# allowance
df = df.drop(['Youth Allowance (other)', 'Youth Allowance (student and apprentice)'], axis=1)

# Save the modified dataframe back to a new Excel file called Final_Benefit in LGA folder
output_file_path = '/Users/hanaarshadahmed/Desktop/LGA/Benefits/Final_Benfit.csv'
df.to_csv(output_file_path, index=False)
