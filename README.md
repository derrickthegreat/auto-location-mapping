# Auto Location Mapping v1.2.4

Auto location mapping tool to help map a list of address stored in a .xlsx file.

Currently works on Windows & Linux only.

## Prerequisites:

- [Google Earth](https://www.google.com/earth/versions/)
- [Python](https://www.python.org/downloads/)
- [pyautogui](https://pyautogui.readthedocs.io/en/latest/)
- [openpyxl](https://openpyxl.readthedocs.io/en/stable/)
- [pandas](https://pandas.pydata.org/docs/getting_started/index.html)

## Getting Started:

> #### Clone repo

```bash
git clone https://github.com/derrickthegreat/auto-location-mapping.git
```

> #### Install the tool

```bash
pip install auto-location-mapping
```
> #### Run the tool anywhere via cli

```bash
automap
```


## Additional Notes:

**Before running the script, please make sure that:** 

 The location spreadsheet includes a column titled *'address'* (case-sensitive) with the search terms.
The **File path** is based on where you open the `automap` command.
