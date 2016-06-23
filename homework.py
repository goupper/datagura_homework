# coding: utf-8

import xlwt
import xlrd
from xlwt import Style, Alignment

files = open('text.txt', 'r')
file_lines = files.readlines()

wb = xlwt.Workbook()
ws = wb.add_sheet(u'测试xlwt')

align = Alignment()
align.horz = Alignment.HORZ_CENTER
align.vert = Alignment.VERT_CENTER
style = Style.XFStyle()
style.alignment = align
line_list = []
for res_key, res in enumerate(file_lines):
    # 以tab符号split
    line_list = res.split('\t')
# 读txt文件写入xls文件中
    for k, v in enumerate(line_list):
        ws.write(res_key, k, v, style)

wb.save('xlwt_demo.xls')
# 打开文件
workbook = xlrd.open_workbook('xlwt_demo.xls')
# 获取sheet名字
sheet = workbook.sheet_by_name(u'测试xlwt')
# 读文件写入txt， sheet.nrows表示所有的行数，sheet.row_values(key)表示读取一行数据
for key in xrange(0, sheet.nrows):
    with open('xlrd_demo.txt', 'a+') as f:
        # print sheet.row_values(key)
        for va in sheet.row_values(key):
            f.write(str(va).encode('utf-8'))



