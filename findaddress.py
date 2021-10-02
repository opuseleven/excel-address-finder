#! python3

import os
import sys
import openpyxl
from selenium import webdriver
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
import time

if len(sys.argv) < 2:
    print('Usage: python3 findaddress.py filename.xlsx')
    sys.exit()

filename = sys.argv[1]
path = os.path.join(os.getcwd(), filename)

print(path)

workbook = openpyxl.load_workbook(path)

driver_options = Options()
driver_options.headless = True
driver = webdriver.Firefox(options=driver_options)

def search(term):
    driver.get("https://www.duckduckgo.com/?q=%s" % term)
    print(term)
    time.sleep(2)
    list = driver.find_elements(By.CSS_SELECTOR, "p")
    addressfound = False
    for l in list:
        if l.text.startswith('Address:'):
            addressarr = l.text.split(': ')
            address = addressarr[1]
            addressfound = True
    if addressfound == False:
        print("Couldn't find that address.")
        address = " "
    print(address)
    return address

print("Searching for addresses:")
for sheet in workbook.worksheets:
    titlerow = sheet[1]
    state = sheet.title
    print('\n')
    print('state: %s'% state)
    namecol = -1
    citycol = -1
    addresscol = -1
    count = 0
    for cell in titlerow:
        if cell.value == 'Name':
            namecol = count
        if cell.value == 'City':
            citycol = count
        if cell.value == 'Address':
            addresscol = count
        count += 1
    if namecol == -1:
        break
    if addresscol == -1:
        sheet.insert_cols(5)
        addresscol = 4
        sheet['E1'].value = 'Address'
    for row in sheet.iter_rows(min_row=2):
        # search for address
        name = row[namecol].value
        city = row[citycol].value
        searchterm = str("%s %s %s"%(name, city, state))
        if searchterm.startswith('None'):
            address = ' '
        else:
            address = search(searchterm)
        row[addresscol].value = address

def tearDown(self):
    self.quit()

tearDown(driver)

# Write workbook
print("Writing file...")
split_filename = os.path.splitext(filename)
new_filename = split_filename[0] + '-address' + split_filename[1]
workbook.save(new_filename)
print("Done.")
