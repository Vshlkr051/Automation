# Python code to merge n number of sheets into single sheet if and only if "all sheet has same columns"
# incase we need to exclude one sheet, can be used 
# ====================================================
# First install "pandas" package into system using either of below
# >>> pip install pandas 
# >>> pip3 install pandas
# ====================================================

import pandas as pd

input_file = input("enter the file path contaning many sheets : ")
sheet_to_be_excluded = input("enter the sheet name that you want to exclude : ")

# read the Excel file
excel_file = pd.ExcelFile(input_file)

# create an empty DataFrame to store the merged data
merged_data = pd.DataFrame()

# loop through each sheet in the file
for sheet_name in excel_file.sheet_names:
    if sheet_name!=sheet_to_be_excluded:
        # read the sheet into a DataFrame
        sheet_data = excel_file.parse(sheet_name, skiprows=2)
        # append the sheet data to the merged data
        merged_data = merged_data.append(sheet_data)

# write the merged data to a new Excel file
merged_data.to_excel('merged_file.xlsx', index=False)