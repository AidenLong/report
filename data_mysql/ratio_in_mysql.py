# -*- coding utf-8 -*- #
import xlrd

from utils import mysql_util

file_name = 'XJ Info_RedBook.xlsx'
db = mysql_util.get_db(host='localhost', user='root', password='root1234', port=3306, db='redbook')

# sql = 'INSERT INTO `year_mouth_group_ratio`(year_mouth_group, mouth, ratio) VALUES(%s, %s, %s)'
#
# wb = xlrd.open_workbook(file_name)
# sheet = wb.sheet_by_index(1)
# nrows = sheet.nrows
# ncols = sheet.ncols
# for row in range(21, nrows):
#     year_mouth_group = sheet.cell(row, 0).value
#     print(row)
#     for col in range(2, 15):
#         # print((year_mouth_group, col - 2, sheet.cell(row, col).value))
#         mysql_util.excete(db, sql, (year_mouth_group, col - 2, sheet.cell(row, col).value))

# 检测导入成功？
sql = 'select ratio from year_mouth_group_ratio where year_mouth_group = %s and mouth = %s'
data = mysql_util.select(db, sql, ('200912', 0))
print(data[0][0])

mysql_util.close(db)
