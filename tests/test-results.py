#! python3

# TEST

import sys
import os
import openpyxl
import re

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
            pattern = "\d.+[A-Za-z]+,\s[A-Z]{2}\s\d{5}-?\d{0,4}.*"
            # if row[addresscol].value != ' ':
            if re.match(pattern, row[addresscol].value):
                addresscount += 1
addstring = str(addresscount)
totalstring = str(objcount)
percentage = (addresscount / objcount) * 100
percentage = round(percentage, 2)
stdcolor = '\033[0m' # white
color = '\033[31m' # red
if percentage > 59.9:
    color = '\033[33m' # orange
if percentage > 74:
    color = '\033[32m' # green
perstring = str(percentage)
print('Found %s addresses out of %s'% (addstring, totalstring))
printablestring = "{}% successful".format(perstring)
print(color + printablestring + stdcolor)
