# 
# auto excel processing by SLQ

import os
import openpyxl # Is it different from "from openpyxl import Workbook" 
# from openpyxl import Workbook

os.chdir(r"C:\Users\shaolqi\Desktop")

wb = openpyxl.load_workbook('example.xlsx')

sheet = wb.get_sheet_by_name('Sheet1')
#---------------------------------------------------------------------------------
# choose the range of table you want
r_want = int(input('enter row number you want:'))       
c_want = int(input('enter column number you want:'))   
# tuple(sheet['A1':'C3'])

print(sheet.cell(row=r_want,column=c_want).value)
# for i in range(1,8):
#     print(i,sheet.cell(row=i,column=3).value)
    
# for i in range(sheet.max_row,sheet.max_column,2):
#     print (i,sheet.cell(row=i,column=2).value)

# for rowOfCellObjects in sheet['A1':'C3']:
#     for cellObj in rowOfCellObjects:
#         print(cellObj.coordinate, cellObj.value)
#     print('----END of ROW----')
