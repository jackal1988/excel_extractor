#!/usr/bin/env python3
# -*- coding: utf-8 -*-

# writeNewEnergyCarCatalogExcel.py  by SLQ

import os
os.chdir('C:\\Users\\shaolqi\\PycharmProjects\\excel_extractor')

import openpyxl, newCatalog4NEC
from openpyxl.utils import get_column_letter
# ---------------------create a new workbook to store newCatalog4NEC.py
destFileName = '自用新能源车型目录.xlsx'
wbUpdate = openpyxl.load_workbook(destFileName)
wsUpdate = wbUpdate.active

lsWsColValueAll = []  # iterate all contents in first row of 
for m in range(1, wsUpdate.max_column+1):
    lsWsColValueAll.append(wsUpdate[get_column_letter(m) + '1'].value)
tupSheetColValueAll = tuple(lsWsColValueAll)
# print(lsWsColValueAll)
# lsKeyIn = []
# for n in newCatalog4NEC.allData.keys():
# #     if newCatalog4NEC.allData[n]
