from openpyxl import load_workbook

filepath = "C:\\Users\chang\OneDrive\바탕 화면\pytho_excel\Development_Mentoring-For-Upskill\Python_control_excel"
wb =load_workbook(filepath)
sheet = wb.active
b1 = sheet['B1']
# sheet.cell[row =1, column =3].value =>가 많이쓰인다 이렇게 출력하면 더 좋음
# 출력하면 튜플로 된 리스트로 얻을 수 있다.
print(b1.value)