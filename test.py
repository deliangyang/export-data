import re

a = '胡xxx13002341331北京市 海淀区  清华大学生命科学新馆'

reg_detail = re.compile(r'^([^\d]+)(\d+)([^$]+)')


def split(detail):
    print(reg_detail.findall(detail))


split(a)

a = [1, 3]
b = [1, 2]
print(a + b)


count = 0
with open('b', 'r') as f:
    for line in f.readlines():
        if line:
            count += int(line)
    f.close()

print(count)
