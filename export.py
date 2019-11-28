import xlrd
import json


wb = xlrd.open_workbook('590aa102d809444f.xlsx')
zl = wb.sheet_by_name('原始表格')

items = {}
for row_index in range(1, zl.nrows):
    order_no = str(zl.cell(rowx=row_index, colx=0).value)
    address = str(zl.cell(rowx=row_index, colx=23).value).strip()
    items[address] = order_no

with open('xx.json', 'w') as f:
    f.write(json.dumps(items))
    f.close()

print(len(items))
