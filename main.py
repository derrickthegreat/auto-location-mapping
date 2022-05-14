import pyautogui as pag
import pandas as pd
import os
import time

# User Input
input_entry_finished = False
default_path = 'google-earth.xlsx'

print('Google Earth Mapping Automation v1.0')
print('by Derrick Alvarez')
print()
print('Before you begin, please make sure to have\n 1) Google Earth Pro already open\n 2) Formatted Address List per README.md')
print()

while not input_entry_finished:
    # File Path & Name Insured Entry
    named_insured = input('Please enter the Named Insured [string]:')
    if not named_insured:
        named_insured = 'test'
        # named_insured = input('Please enter the named insured: ')
    print()
    file_path = input('Please enter file path [string]:')
    if not file_path:
        file_path = default_path
    print()
    rows = int(input('How many locations are there? [int]:'))

    # Review & Confirm
    print('You have entered:\n')
    print('1) Named Insured: ', named_insured)
    print('2) File Path: ', file_path)
    print('3) Number of Rows: ', rows, '\n')
    finished = input('Is this correct? [Y/n]: ')
    if finished == 'Y' or finished == 'y' or finished == '':
        input_entry_finished = True


# Pull data from Excel
excel_data = pd.read_excel(file_path)

# Launch Google Earth & set up folder 
with pag.hold('alt'):
    pag.press('tab')
with pag.hold(['ctrl', 'shift']):
    pag.press('n')
pag.write(named_insured)
pag.press('enter')
pag.press('tab')

time.sleep(1)

# Iterate through Excel
count = 0

def clear_search():
    with pag.hold('ctrl'):
        pag.press('a')
        pag.press('delete')

while count < 1:
    address = excel_data['address'][count]
    # Enter address
    pag.typewrite(address)
    pag.press('enter')
    time.sleep(1)

    # Save / Drop Pin
    with pag.hold(['ctrl', 'shift']):
        pag.press('p')
    pag.typewrite(address)
    pag.press('enter')
    time.sleep(1)

    clear_search()

    count += 1

print()
print('All addresses have been entered!')