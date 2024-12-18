# -*- coding: utf-8 -*-
import os
from openpyxl import Workbook

def splitLine(line: str):
    if len(line) < 20 or line[0] != "0" or line[1] != "1":
        return -1, -1
    key = ""
    for i in range(2, len(line) - 1):
        if line[i] == '2' and line[i + 1] == '1':
            val = line[i + 2 : len(line) - 6]
            return key, val
        key += line[i]
    return -1, -1


file = open("kv.txt", "r")
text = file.read()

kv = {}
for line in text.splitlines():
    key, value = line.split(" ")
    kv[key] = value

file.close()
file = open("input.txt", "r")
text = file.read()

data = {}
trash = []
for line in text.splitlines():
    key, val = splitLine(line)
    if key != -1:
        data.setdefault(key, set()).add(val)
    else:
        trash.append(line)

def fillExcel(codes, key: str):
    filepath = os.path.join("results", kv[key] + ".xlsx")
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"

    for code in codes:
        costil = ["01" + key + "21" + code]
        ws.append(costil)

    wb.save(filepath)

for key, codes in data.items():
    fillExcel(codes, key)

filepath = os.path.join("results", "trash.xlsx")
wb = Workbook()
ws = wb.active
ws.title = "Sheet1"

for s in trash:
    costil = [s]
    ws.append(costil)

wb.save(filepath)