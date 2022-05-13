import pyautogui
import pandas
import time

# File path input

file_path = input('enter file path')

if not file_path:
    file_path = 'address-list.xlsx'

# Pull data from Excel
excel_data = pandas.read_excel('personal-budgetting-tool.xlsx')
total_rows = 100
count = 0
time.sleep(2)

# Iterate through Excel
for col in excel_data:
    # Enter address
    pyautogui.click(100, 500)
    pyautogui.write(excel_data['address'][count])
    pyautogui.press('enter')
    # Clear Search
    pyautogui.click(100, 500)
    # End of Data Entry
    if count == total_rows:
        break;
    count += 1

print('All Addresses Entered.')