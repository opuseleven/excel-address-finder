#! python3

import os
import sys
import openpyxl

if len(sys.argv) < 2:
    print('Usage: python3 findaddress.py filename.xlsx')
    sys.exit()

filename = sys.argv[1]
path = os.path.join(os.getcwd(), filename)

print(path)

workbook = openpyxl.load_workbook(path)

namecol = null
citycol = null

for sheet in workbook.worksheets:
    titlerow = sheet[1]
    state = sheet.title
    for cell in titlerow:
        if cell.value = 'Name':
            namecol = cell.column
        if cell.value = 'City':
            citycol = cell.column
    if namecol = null:
        break
    for row in sheet.iter_rows():
        # search for address
