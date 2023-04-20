import openpyxl

dictionary = input("enter the dictionary file path : ")
file = input("enter the file path to translate : ")
target_file = input("enter the destination path with file name : ")

# Load "xlsx file 1"
wb1 = openpyxl.load_workbook(dictionary)
sheet1 = wb1.active

# Load "xlsx file 2"
wb2 = openpyxl.load_workbook(file)
sheet2 = wb2.active

# Create a dictionary to store the translation mappings from "xlsx file 1"
translation_dict = {}
for row in sheet1.iter_rows(values_only=True):
    translation_dict[row[0]] = row[1]

# Translate data in "xlsx file 2" based on the mappings
for row in sheet2.iter_rows():
    for cell in row:
        if cell.value in translation_dict:
            cell.value = translation_dict[cell.value]

# Save the translated data to a new Excel file
wb2.save(target_file)