import pyautogui as pag
import pandas as pd
import time
import subprocess
import sys


# Functions
def toggle_sidebar():
    with pag.hold(['ctrl', 'alt']):
        pag.press('b')

def clear_search():
    with pag.hold('ctrl'):
        pag.press('a')
        pag.press('delete')

def sys_compatible():
    platforms = {
        'win32' : 'Windows',
        'linux' : 'Linux',
        'linux2' : 'Linux'
    }
    if sys.platform not in platforms:
        return False
    return [True, platforms[sys.platform]]

def automap():
    print('Auto Location Mapping v1.2.1','\n','by Derrick Alvarez','\n')
    print()
    print('Checking OS...')
    os = sys_compatible()
    if not os:
        return print('This OS is not supported yet.')
    elif os[1] == 'Linux':
        print('Running on Linux ;-)')
        google_earth = 'google-earth-pro'
    else:
        print('Running on Windows')
        google_earth = [r"C:\Program Files\Google Earth Pro\client\googleearth.exe"]
    
    input_entry_finished = False
    default_path = './google-earth.xlsx'
    
    # User Input
    while not input_entry_finished:
        print()
        #Name, file & # of locations
        named_insured = input('Please enter the Named Insured: ')
        if not named_insured:
            named_insured = input('Please enter the named insured: ')
        file_path = input('Please enter file path [string]: ')
        if not file_path:
            file_path = default_path
        num_of_locations = int(input('How many locations are there? [int]: '))
        if not num_of_locations:    
            num_of_locations = int(input('How many locations are there? [int]: '))
        print()

        # Review & Confirm
        print('You have entered:\n')
        print('1) Named Insured: ', named_insured)
        print('2) File Path: ', file_path)
        print('3) Number of Locations: ', num_of_locations, '\n')
        finished = input('Is this correct? [Y/n]: ')
        if finished == 'Y' or finished == 'y' or finished == '':
            input_entry_finished = True

    # Pull data from Excel
    excel_data = pd.read_excel(file_path)

    # Go to Google Earth & Set Up Folder
    subprocess.Popen(google_earth)
    time.sleep(3)
    toggle_sidebar()
    with pag.hold(['ctrl', 'shift']):
        pag.press('n')
    pag.write(named_insured)
    pag.press('enter')

    # Iterate through Excel
    count = 0

    while count < num_of_locations:
        address = excel_data['address'][count]
        # Enter address
        pag.typewrite(address)
        pag.press('enter')
        time.sleep(0.5)
        # Save / Drop Pin
        with pag.hold(['ctrl', 'shift']):
            pag.press('p')
        pag.typewrite(address)
        pag.press('enter')
        time.sleep(0.5)

        clear_search()
        count += 1

    toggle_sidebar()
    print('\nAll addresses have been entered!')
