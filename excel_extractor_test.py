#!/usr/bin/env python3
# -*- coding: utf-8 -*-

# readCensusExcel.py - calculate population by SLQ
# for each county
import openpyxl, pprint, os
from openpyxl import Workbook


# from openpyxl.cell.cell import column_index_from_string

os.chdir(r"C:\Users\Administrator\Desktop")
print('Opening workbook...')
wb = openpyxl.load_workbook('免征车辆购置税的新能源汽车车型目录（第十五批）.xlsx')
sheet = wb.get_sheet_by_name('Table 4')
# ------------create a new workbook to store the selected data
# wbTemp = Workbook()
# sheetTemp = wbTemp.create_sheet('newEnergyCarList', 0)
# print(type(sheet))
# ------------------------#customize the rectangle area you want-----------------------
# startRow    = int(input("enter the start row number you want:"))
# startColumn = input("enter the start column letter you want:")
# endRow      = int(input("enter the end row number you want:"))
# endColumn   = input("enter the end column letter you want:")
# startCell = startColumn + str(startRow)
# endCell   = endColumn + str(endRow)
startCell = 'a2'
endCell = 'j55'
recArea = sheet[startCell:endCell]  # recArea type is tuple
# listRecArea = list(recArea)
# print(recArea[0][1].value)
# Loop through recArea --------------------------------------------------------------------------------------

lsEachDictKey = []
lsEachDictValueAll = []
for rowNum in range(0,len(recArea)):                     #iterate all row in recArea
    if rowNum == 0: #如果是第一行，则存为key
        for eachCell in recArea[rowNum]:
                lsEachDictKey.append(eachCell.value)
        for m in range(0,len(lsEachDictKey)):  # 去掉key字符串里的换行符
            if lsEachDictKey[m] != None:
                lsEachDictKey[m] = lsEachDictKey[m].replace('\n','')
            else:
                continue
        tupEachDictKey = tuple(lsEachDictKey)
    else:
        for eachCell in recArea[rowNum]:
            lsEachDictValueAll.append(eachCell.value)
        for m in range(0,len(lsEachDictValueAll)):  # 去掉value字符串里的换行符,注意replace只能针对str型数据
            if lsEachDictValueAll[m] != None and type(lsEachDictValueAll[m]) == str:
                lsEachDictValueAll[m] = lsEachDictValueAll[m].replace('\n','')
            else:
                continue

    # dictData.setdefault(tupEachDictKey,lsEachDictValue)
pprint.pprint(lsEachDictKey)
# Regroup lsEachDictValue-------------------------------------
lsEachDictValueAllRegroup = []
for c in range(0,len(lsEachDictKey)):
    n = c
    while n < len(lsEachDictValueAll):
        lsEachDictValueAllRegroup.append(lsEachDictValueAll[n])
        n += len(lsEachDictKey)
    c += 1
pprint.pprint(lsEachDictValueAllRegroup)
stepLen = len(recArea)-1
# Auto fill up none company name value cell --------------------------------------
for m in range(stepLen,2*stepLen):   # This method may not be flawless.
    if lsEachDictValueAllRegroup[m] == None:
        lsEachDictValueAllRegroup[m] = lsEachDictValueAllRegroup[m-1]
# pprint.pprint(lsEachDictValueAllRegroup)

# Dictionary data structure-----------------------------------------------------------
dictData = {}
for n in range(0,len(lsEachDictKey)):
    dictData.setdefault(lsEachDictKey[n],lsEachDictValueAllRegroup[n*stepLen:(n+1)*stepLen])
dictData.pop(None)  # pop out None key-value
pprint.pprint(dictData)

