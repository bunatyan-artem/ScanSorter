# -*- coding: utf-8 -*-
import os
from openpyxl import Workbook

for filename in os.listdir(os.path.join("results")):
    file_path = os.path.join("results", filename)
    os.unlink(file_path)

def splitLine(line: str):
    if len(line) < 20 or line[0] != "0" or line[1] != "1":
        return -1, -1
    GTIN = ""
    for i in range(2, len(line) - 1):
        if line[i] == '2' and line[i + 1] == '1':
            return GTIN, line[i + 2 : len(line) - 6]
        GTIN += line[i]
    return -1, -1


file = open("pc.txt", "r")
text = file.read()

cp = {}
pc = {}
for line in text.splitlines():
    product = line.split(" ")[0]
    GTINs = line.split(" ")[1:]
    
    for GTIN in GTINs:
        cp[GTIN] = product
    pc[product] = GTINs

file.close()
file = open("input.txt", "r")
text = file.read()

data = {}
trash = []
for line in text.splitlines():
    GTIN, code = splitLine(line)
    if GTIN != -1:
        data.setdefault(GTIN, set()).add(code)
    else:
        trash.append(line)

def fillExcel(product: str):
    filepath = os.path.join("results", product + ".xlsx")
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"

    isEmpty = True
    for GTIN in pc[product]:
        if GTIN not in data:
            continue
        isEmpty = False
        for code in data[GTIN]:
            ws.append(["01" + GTIN + "21" + code])

    wb.save(filepath)
    if isEmpty:
        file_path = os.path.join("results", product + ".xlsx")
        os.unlink(file_path)

for product, _ in pc.items():
    fillExcel(product)

if trash:
    filepath = os.path.join("results", "trash.xlsx")
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"

    for s in trash:
        ws.append([s])

    wb.save(filepath)