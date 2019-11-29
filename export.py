import xlrd
import json


class SetEncoder(json.JSONEncoder):
    def default(self, obj):
        if isinstance(obj, set):
            return list(obj)
        return json.JSONEncoder.default(self, obj)


wb = xlrd.open_workbook('590aa102d809444f.xlsx')
zl = wb.sheet_by_name('原始表格')

items = {}
for row_index in range(1, zl.nrows):
    order_no = str(zl.cell(rowx=row_index, colx=0).value)
    name = str(zl.cell(rowx=row_index, colx=18).value).strip()
    mobile = str(zl.cell(rowx=row_index, colx=19).value).strip()
    address = str(zl.cell(rowx=row_index, colx=23).value).strip()
    if address not in items:
        items[address] = {}
    name_mobile = name + mobile
    if name_mobile not in items[address]:
        items[address][name_mobile] = set()
    items[address][name_mobile].add(order_no)

with open('xx.json', 'w') as f:
    f.write(json.dumps(items, cls=SetEncoder))
    f.close()

print(len(items))
