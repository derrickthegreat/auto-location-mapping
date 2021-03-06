import pyautogui as pag
import pandas as pd
from progress.spinner import MoonSpinner
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

def main():
    print('Auto Location Mapping v1.2.4','\n','by Derrick Alvarez','\n')
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
        google_earth = r"C:\Program Files\Google\Google Earth Pro\client\googleearth.exe"
        print(google_earth)
    
    input_entry_finished = False
    default_path = './google-earth.xlsx'
    
    # User Input
    while not input_entry_finished:
        print()
        #Name, file & # of locations
        collection_name = input('Please enter the Collection Name: ')
        if not collection_name:
            collection_name = input('Please enter the Collection Name: ')
        file_path = input('Please enter file path: ')
        if not file_path:
            file_path = default_path
        print()

        # Review & Confirm
        print('You have entered:\n')
        print('1) Collection Name: ', collection_name)
        print('2) File Path: ', file_path)
        finished = input('Is this correct? [Y/n]: ')
        if finished == 'Y' or finished == 'y' or finished == '':
            input_entry_finished = True

    # Pull data from Excel
    excel_data = pd.read_excel(file_path)

    # Go to Google Earth & Set Up Folder
    subprocess.Popen(google_earth)
    time.sleep(5)
    toggle_sidebar()
    with pag.hold(['ctrl', 'shift']):
        pag.press('n')
    pag.write(collection_name)
    pag.press('enter')

    # Iterate through Excel
    count = 0

    print()
    with MoonSpinner('Processing...') as bar:
        for row in excel_data['address']:
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
            time.sleep(0.02)
            bar.next()

    toggle_sidebar()
    print('\nAll addresses have been entered!')

if __name__ == "__main__":
    main()