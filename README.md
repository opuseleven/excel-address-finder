# excel-address-finder

A command line program to search for the addresses of a large number of US businesses/locations in a spreadsheet. The python script takes an excel document as input, recognizes columns labeled "Name", and "City". Then creates a new column labeled "Address". The program then scrapes the web for the addresses. Finally, the addresses are written to a new copy of the excel file "FileName-addresses.xlsx" with added "Address" column. 

Usage: python3 findaddress.py filename.xlsx

The results can be tested with test-results.py in the /tests directory. This script compares the number of addresses found to the total number of locations in the spreadsheet.
