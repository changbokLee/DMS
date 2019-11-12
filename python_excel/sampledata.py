from openpyxl import Workbook
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.chart import(
    AreaChart,
    Reference,
    Series
)

wb = Workbook() # 워크북 만듬
ws = wb.create_sheet() # sheet 생성
ws.title = 'Fruit' 

data = [
    ['Apples', 10000, 5000, 8000, 6000],
    ['Pears',   2000, 3000, 4000, 5000],
    ['Bananas', 6000, 6000, 6500, 6000],
    ['Oranges',  500,  300,  200,  700],
]

# 헤더 추가함
ws.append(["Fruit", "2011", "2012", "2013", "2014"])
for row in data:
    ws.append(row)

tab = Table(displayName="Table1", ref="A1:E5")

# 테이블 스타일 적용
style = TableStyleInfo(name="TableStyleMedium9", showFirstColumn=False,
                       showLastColumn=False, showRowStripes=True, showColumnStripes=True)
tab.tableStyleInfo = style
ws.add_table(tab)

# 차트 만들기
chart = AreaChart()
chart.title ="Fruit"
chart.style= 13
chart.x_axis.title ="Years"
chart.y_axis.title = "Fruit"

wb.save(filename ="table.xlsx") 