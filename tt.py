import openpyxl

file_path = 'test.xlsx'
workbook = openpyxl.load_workbook(file_path)
active_sheet= workbook.active
print(active_sheet)