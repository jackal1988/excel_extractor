#!/usr/bin/env python3
# -*- coding: utf-8 -*-

# readCensusExcel.py - calculate population by SLQ
# for each county
import openpyxl, pprint, os
from openpyxl import Workbook


# from openpyxl.cell.cell import column_index_from_string

os.chdir(r"C:\Users\shaolqi\Desktop")
print('Opening workbook...')
wb = openpyxl.load_workbook('免征车辆购置税的新能源汽车车型目录（第十五批）.xlsx')
sheet = wb.get_sheet_by_name('Table 2')
# ------------create a new workbook to store the selected data
wbTemp = Workbook()
sheetTemp = wbTemp.create_sheet('newEnergyCarList', 0)
# print(type(sheet))
# ------------------------#customize the rectangle area you want-----------------------
# startRow    = int(input("enter the start row number you want:"))
# startColumn = input("enter the start column letter you want:")
# endRow      = int(input("enter the end row number you want:"))
# endColumn   = input("enter the end column letter you want:")
# startCell = startColumn + str(startRow)
# endCell   = endColumn + str(endRow)
startCell = 'A2'
endCell = 'J72'
recArea = sheet[startCell:endCell]  # recArea type is tuple
# listRecArea = list(recArea)
# print(recArea[0][1].value)
# Loop through recArea --------------------------------------------------------------------------------------
dictData = {}
lsEachDictKey = []
lsEachDictValueAll = []
for rowNum in range(0,len(recArea)):                     #iterate all row in recArea
    if rowNum == 0: #如果是第一行，则存为key
        for eachCell in recArea[rowNum]:
            # if eachCell.value != None:
                lsEachDictKey.append(eachCell.value)
            # else:
            #     continue
        tupEachDictKey = tuple(lsEachDictKey)
    else:
        for eachCell in recArea[rowNum]:
            lsEachDictValueAll.append(eachCell.value)

    # dictData.setdefault(tupEachDictKey,lsEachDictValue)
# pprint.pprint(lsEachDictKey)
#regroup lsEachDictValue-------------------------------------
lsEachDictValueAllRegroup = []
for c in range(0,len(lsEachDictKey)):    # Will there be any error in using variable c?
    while c < len(lsEachDictValueAll):
        lsEachDictValueAllRegroup.append(lsEachDictValueAll[c])
        c += len(lsEachDictKey)
    c += 1
pprint.pprint(lsEachDictValueAllRegroup)


# wbTemp.save('newEnergyCarList')
