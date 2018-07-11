import openpyxl
from openpyxl.styles import PatternFill, Border, Side, Alignment, Protection, Font
from openpyxl import Workbook

'''
1.拿到txt文件
2.放在这个目录，运行Python文件
3.生成新的Excel文件，单元格自动居中
'''

con = input("输入你想转换的文件名:\n")
file = "{}.txt".format(con)
wb = Workbook()  # 引入workbook类
ws = wb.active
# 一个工作簿(workbook)在创建的时候同时至少也新建了一张工作表(worksheet)。你可以通过openpyxl.workbook.Workbook.active()调用得到正在运行的工作表。（注意：该函数调用工作表的索引(_active_sheet_index)，默认是0。除非你修改了这个值，否则你使用该函数一直是在对第一张工作表进行操作。）

sheet = wb['Sheet']  # 选中表格
column = 1  # 列
row = 3  # 行
count = 0
start = 0
end = 30
title = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L',
         'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z']

score = [
    "1:0",
    "0:0",
    "0:1",
    "2:0",
    "1:1",
    "0:2",
    "2:1",
    "2:2",
    "1:2",
    "3:0",
    "3:3",
    "0:3",
    "3:1",
    "3:2",
    "1:3",
    "2:3",
    "1:0",
    "0:0",
    "0:1",
    "2:0",
    "1:1",
    "0:2",
    "2:1",
    "2:2",
    "1:2",
    "3:0",
    "3:3",
    "0:3",
    "3:1",
    "3:2",
    "1:3",
    "2:3"]
coord = [
    "93-223",
    "273-239",
    "448-243",
    "92-320",
    "270-324",
    "448-317",
    "88-397",
    "293-397",
    "450-398",
    "88-475",
    "265-477",
    "444-478",
    "88-550",
    "86-635",
    "450-552",
    "449-636",
    "93-223",
    "273-239",
    "448-243",
    "92-320",
    "270-324",
    "448-317",
    "88-397",
    "293-397",
    "450-398",
    "88-475",
    "265-477",
    "444-478",
    "88-550",
    "86-635",
    "450-552",
    "449-636"]


def top_two_cell(cell_num):  # 前2个单元格内容填充并赋予格式
    cell1 = sheet['{}{}'.format(y, 1)]
    cell2 = sheet['{}{}'.format(y, 2)]
    cell1.value = score[count]
    cell2.value = coord[count]
    cell1.alignment = Alignment(
        horizontal='center', vertical='center')  # 对单元格居中
    cell2.alignment = Alignment(
        horizontal='center', vertical='center')  # 对单元格居中


def cell_value():
    cell = sheet['{}{}'.format(y, row)]  # 指定坐标
    cell.value = x  # 坐标赋值
    cell.alignment = Alignment(
        horizontal='center', vertical='center')  # 对单元格居中


with open(file, 'r') as f:
    s = f.read().split("\n")
    for y in title:
        top_two_cell(y)  # 前2个单元格内容填充
        row = 3  # 重置row的值
        for x in s[start: end]:
            cell_value()
            row += 1
        sheet.column_dimensions[y].width = 30
        start += 30
        end += 30  # 30个为一组
        count += 1

wb.save(filename='{}.xlsx'.format(file.split('.')[0]))
print("转换完成")
