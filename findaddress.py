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
    driver.get("https://www.google.com/search?q=%s" % term)
    time.sleep(2)
    address = driver.find_element(By.CLASS_NAME, "LrzXr")
    return address.text

for sheet in workbook.worksheets:
    titlerow = sheet[1]
    state = sheet.title
    namecol = null
    citycol = null
    addresscol = null
    for cell in titlerow:
        if cell.value = 'Name':
            namecol = cell.column
        if cell.value = 'City':
            citycol = cell.column
        if cell.value = 'Address':
            addresscol = cell.column
    if namecol = null:
        break
    if addresscol = null:
        sheet.insert_cols(5)
        addresscol = sheet['E']
        sheet['E1'] = "Address"
    for row in sheet.iter_rows(min_row=2):
        # search for address
        name = row[namecol].value
        city = row[citycol].value
        searchterm = Str("%s %s %s"%(name, city, state))
        address = search(searchterm)
        row[addresscol] = address

def tearDown(self):
    self.quit()

tearDown(driver)

# Write workbook
split_filename = os.path.splitext(filename)
new_filename = split_filename[0] + '-address' + split_filename[1]
workbook.save(new_filename)
