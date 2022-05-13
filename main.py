import pag as pag
import pandas as pd
import subprocess


# User Input
input_entry_finished = False

while not input_entry_finished:
    # File Path & Name Insured Entry
    file_path = input('Please enter file path: ')
    if not file_path:
        file_path = 'address-list.xlsx'
    print()
    named_insured = input('Please enter insureds name: ')
    print()

    # Review & Confirm
    print('You have entered:')
    print('1) File path: ', file_path)
    print('Named Insured: ', named_insured)
    finished = input('Is this correct? [Y/n]: ')
    if finished == 'Y' | finished == 'y' | finished == False:
        input_entry_finished = True


# Pull data from Excel
excel_data = pd.read_excel(file_path)

# Launch Google Earth & set up folder 
subprocess.Popen("C:\Program Files\Google\Google Earth Pro\client\googleearth.exe")
pag.keyDown('ctrl')
pag.keyDown('shift')
pag.press('n')
pag.keyUp('ctrl')
pag.keyUp('shift')
pag.write(named_insured)
pag.press('enter')
pag.press('tab')

# Iterate through Excel
count = 0

while excel_data['address'][count]:
    address = excel_data['address'][count]
    # Enter address
    pag.write(address)
    pag.press('enter')
    # Save / Drop Pin
    pag.keyDown('ctrl')
    pag.keyDown('shift')
    pag.press('p')
    pag.keyUp('ctrl')
    pag.keyUp('shift')
    pag.write(address)
    pag.press('enter')
    # Clear Search
    pag.keyDown('ctrl')
    pag.press('a')
    pag.keyUp('ctrl')
    # End of Data Entry
    count += 1

print('All Addresses Entered.')