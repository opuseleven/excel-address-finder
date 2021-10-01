#! python3

import sys
import os
import openpyxl

if len(sys.argv) < 2:
    print('Usage: python3 test-results.py filename-address.xlsx')
    sys.exit()

filename = sys.argv[1]
path = os.path.join(os.getcwd(), filename)

workbook = openpyxl.load_workbook(path)

print('Counting...')

objcount = 0
addresscount = 0

for sheet in workbook.worksheets:
    titlerow = sheet[1]
    counter = 0
    addresscol = -1
    for cell in titlerow:
        if cell.value == 'Address':
            addresscol = counter
        counter += 1
    if addresscol == -1:
        print("Error: Couldn't find \"Address\" column.")
        break
    for row in sheet.iter_rows(min_row=2):
        if row[0].value:
            objcount += 1
            if row[addresscol].value != ' ':
                addresscount += 1
addstring = str(addresscount)
totalstring = str(objcount)
print('Found %s addresses out of %s'% (addstring, totalstring))
