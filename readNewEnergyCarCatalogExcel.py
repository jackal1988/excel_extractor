#!/usr/bin/env python3
# -*- coding: utf-8 -*-

# readNewEnergyCarCatalogExcel.py  by SLQ

import openpyxl, pprint, sys
from tkinter.filedialog import askopenfilename
# from openpyxl import Workbook
# os.chdir(r"C:\Users\Administrator\Desktop")
# ---------------------Load workbook and worksheet----------------------------------------
print('正在打开工作簿......\n')
fname = askopenfilename()
wb = openpyxl.load_workbook(fname)
sheetAll = wb.sheetnames
print('此工作簿包含 ' + str(sheetAll) + ' 工作表\n')
sheetSelect = input('请输入要选择的表名。\n（请注意大小写）\n')

while sheetSelect not in sheetAll:
    sheetSelect = input('二货！输错啦！\n再输一遍吧。\n\n')
sheet = wb[sheetSelect]

# ------------------------customize the rectangle area you want-----------------------

startCell = input('输入起始单元格名称（字母+数字）\n')
endCell = input('输入结束单元格名称（字母+数字）\n')
recArea = sheet[startCell:endCell]  # recArea type is tuple
# listRecArea = list(recArea)
# print(recArea[0][1].value)
# -----------------------Loop through recArea -----------------------------------------------

lsEachDictKey = []
lsEachDictValueAll = []
for rowNum in range(0, len(recArea)):  # Iterate all row in recArea
    if rowNum == 0:  # save 1st line as key of dictionary
        for eachCell in recArea[rowNum]:
            if type(eachCell.value) == int:  # attribute column must be selected
                print('属性列不在选择范围内，请重选。\n')
                sys.exit()
            lsEachDictKey.append(eachCell.value)
        for m in range(0, len(lsEachDictKey)):  # standardize Key name, e.g. delete \n, replace () to （）, etc.
            if lsEachDictKey[m] is not None:
                lsEachDictKey[m] = lsEachDictKey[m].replace(' ', '')
                lsEachDictKey[m] = lsEachDictKey[m].replace('\n', '')
                lsEachDictKey[m] = lsEachDictKey[m].replace('(', '（')
                lsEachDictKey[m] = lsEachDictKey[m].replace(')', '）')
                lsEachDictKey[m] = lsEachDictKey[m].replace('动力蓄电池总质量', '动力蓄电池组总质量')
                lsEachDictKey[m] = lsEachDictKey[m].replace('动力蓄电池总能量', '动力蓄电池组总能量')
            else:
                continue
        # tupEachDictKey = tuple(lsEachDictKey)
    else:
        for eachCell in recArea[rowNum]:
            lsEachDictValueAll.append(eachCell.value)
        for m in range(0, len(lsEachDictValueAll)):  # delete '\n' in lsEachDictValueAll.
            if lsEachDictValueAll[m] is not None and type(lsEachDictValueAll[m]) == str:  # 'replace()' method only
                # effective towards str type data.
                lsEachDictValueAll[m] = lsEachDictValueAll[m].replace('\n', '')
            else:
                continue

    # dictData.setdefault(tupEachDictKey,lsEachDictValue)
# pprint.pprint(lsEachDictKey)
# Regroup lsEachDictValue-------------------------------------
lsEachDictValueAllRegroup = []
for c in range(0, len(lsEachDictKey)):
    n = c
    while n < len(lsEachDictValueAll):
        lsEachDictValueAllRegroup.append(lsEachDictValueAll[n])
        n += len(lsEachDictKey)
    c += 1
# pprint.pprint(lsEachDictValueAllRegroup)
stepLen = len(recArea) - 1
# Auto fill up none company name value --------------------------------------
for m in range(stepLen, 2*stepLen):   # This method may not be flawless.
    if lsEachDictValueAllRegroup[m] is None:
        lsEachDictValueAllRegroup[m] = lsEachDictValueAllRegroup[m - 1]
# pprint.pprint(lsEachDictValueAllRegroup)

# Build dictionary data structure-----------------------------------------------------------
dictData = {}
for n in range(0, len(lsEachDictKey)):
    dictData.setdefault(lsEachDictKey[n], lsEachDictValueAllRegroup[n * stepLen:(n + 1) * stepLen])
for eachKey in list(dictData.keys()):  # pop out 'None' key value
    # 被遍历的对象，在被遍历时其数据结构不能更改！！！所以用list(dictData.keys())代替dictData()
    if eachKey is None:
        dictData.pop(eachKey)
# -------------------------------write dictData to .py file
print('写入读取数据...\n')
resultFile = open('newCatalog4NEC.py', 'w', encoding='utf-8')
resultFile.write('allData = ' + pprint.pformat(dictData))
resultFile.close()
print('Done.')
