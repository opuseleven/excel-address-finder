#! python3

import os
import sys
import openpyxl
from selenium import webdriver
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
import time
import re

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
    query = str("%s %s" % (term[0], term[1]))
    driver.get("https://www.duckduckgo.com/?q=%s&kz=1&kp=-2" % query)
    print(query)
    time.sleep(2)
    list = driver.find_elements(By.CSS_SELECTOR, "p")
    addressfound = False
    address = ""
    count = 0
    for l in list:
        count = count + 1
        cityName = term[1].split(', ')[0]
        pattern = str("\S+\sin\s%s" % cityName.title())
        if re.search(pattern, l.text):
            address = list[count].text
            addressfound = True
    if addressfound == False:
        backupResults = backupSearch(term)
        if backupResults != "":
            addressfound = True
            address = backupResults
        else:
            print("Couldn't find that address.")
            address = " "
    print(address)
    print("\n")
    return address

def backupSearch(term):
    query = str("%s %s" % (term[0], term[1]))
    driver.get("https://www.google.com/search?q=%s" % query)
    time.sleep(2)
    list = driver.find_elements(By.CSS_SELECTOR, "span")
    count = 0
    address = ""
    for l in list:
        count = count + 1
        # Fairly inclusive regex for US address
        pat = "\d.+[A-Za-z]+,\s[A-Z]{2}\s\d{5}-?\d{0,4}"
        if "Address:" in l.text:
            if re.match(pat, list[count + 1].text):
                address = list[count + 1].text
                break
            elif re.match(pat, list[count + 2].text):
                address = list[count + 2].text
                break
        if address = "":
            if re.match(pat, l.text):
                address = l.text
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
        location = str("%s, %s" % (city, state))
        searchterm = [name, location]
        if str(searchterm[0]).startswith('None'):
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
