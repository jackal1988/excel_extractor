#!/usr/bin/env python3
# -*- coding: utf-8 -*-

# writeNewEnergyCarCatalogExcel.py  by SLQ

import os
os.chdir('C:\\Users\\shaolqi\\PycharmProjects\\excel_extractor')

import openpyxl, newCatalog4NEC
from openpyxl.utils import get_column_letter
# ---------------------create a new workbook to store newCatalog4NEC.py
destFileName = 'catalog_of_new_energy_car.xlsx'
wbUpdate = openpyxl.load_workbook(destFileName)
wsUpdate = wbUpdate.active


lsWs1stRowValueAll = []  # iterate all contents in first row of wsUpdate
for m in range(1, wsUpdate.max_column + 1):
    lsWs1stRowValueAll.append(wsUpdate[get_column_letter(m) + '1'].value)
tupSheetColValueAll = tuple(lsWs1stRowValueAll)
print(lsWs1stRowValueAll)

for eachKey in newCatalog4NEC.allData.keys():
    if eachKey in lsWs1stRowValueAll:
        pinPoint = lsWs1stRowValueAll.index(eachKey)
        for n in range(0, len(newCatalog4NEC.allData[eachKey])):
            wsUpdate.cell(row = n + 2, column = pinPoint + 1).value = newCatalog4NEC.allData[eachKey][n]
wbUpdate.save(filename='new_file.xlsx')
