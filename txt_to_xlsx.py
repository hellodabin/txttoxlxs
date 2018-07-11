import openpyxl
from openpyxl.styles import PatternFill, Border, Side, Alignment, Protection, Font

'''
1.拿到卖家的账号txt文件
2.放在这个目录，运行Python文件
3.得到txt文件和Excel文件，Excel文件居中并且按16个为单位分组
'''

file = "2018年7月10日 17.txt"
wb = openpyxl.load_workbook("test.xlsx")
sheet = wb['Sheet1']
column = 1  # 列
row = 3  # 行
count = 0
start = 0
end = 30
title = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L',
         'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z']

score = ["1:0", "0:0", "0:1", "2:0", "1:1", "0:2", "2:1", "2:2", "1:2", "3:0", "3:3", "0:3", "3:1", "3:2", "1:3", "2:3",
         "1:0", "0:0", "0:1", "2:0", "1:1", "0:2", "2:1", "2:2", "1:2", "3:0", "3:3", "0:3", "3:1", "3:2", "1:3", "2:3"]
coord = ["93-223", "273-239", "448-243", "92-320", "270-324", "448-317", "88-397", "293-397", "450-398", "88-475",
         "265-477", "444-478", "88-550", "86-635", "450-552", "449-636", "93-223", "273-239", "448-243", "92-320",
         "270-324", "448-317", "88-397", "293-397", "450-398", "88-475",
         "265-477", "444-478", "88-550", "86-635", "450-552", "449-636"]
with open(file, 'r') as f:
    s = f.read().split("\n")
    with open("new_file.txt", 'w+') as newf:
        for y in title:
            cell1 = sheet['{}{}'.format(y, 1)]
            cell2 = sheet['{}{}'.format(y, 2)]
            cell1.value = score[count]
            cell2.value = coord[count]
            cell1.alignment = Alignment(horizontal='center', vertical='center')  # 对单元格居中
            cell2.alignment = Alignment(horizontal='center', vertical='center')  # 对单元格居中
            row = 3  # 重置row的值
            for x in s[start: end]:
                cell = sheet['{}{}'.format(y, row)]  # 指定坐标
                cell.value = x  # 坐标赋值
                cell.alignment = Alignment(horizontal='center', vertical='center')  # 对单元格居中
                row += 1
            sheet.column_dimensions[y].width = 30
            start += 30
            end += 30  # 30个为一组
            count += 1

    wb.save(filename='{}.xlsx'.format(file))
