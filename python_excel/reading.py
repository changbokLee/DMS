import openpyxl

excel_file = openpyxl.load_workbook('국세청_발간책자 목록(2019년 누계).xlsx')
ws = excel_file['hello ']
print(ws['D5'].value)
