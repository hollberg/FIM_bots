"""program_area.py
Load entries to the "Program Area" (NTEE) codes in FIMS
File Maintenance -> Grants -> Grant Code Maintenance -> Program Area -> (NEW)

Alt^F (file maintenance)
G G [Enter] (Grants)
G G [Enter] (Grant Code Maintenance)
P P [Enter] (Program Area)

Code | Description | Budget | Field | Active

new record is Alt^w
"""

# Standard
import time

#External Libraries
import pandas as pd
import pyautogui as bot

#Local External Modules
import gui_tools as gt


# Read Excel File
xl_path = r'P:\03_FinanceOperations\InformationManagement\F2RE\Cleanup\FIMS_NTEE_Load.xlsx'
df_ntee = pd.read_excel(xl_path)

print(df_ntee.head())

time.sleep(1)

for row in df_ntee.iterrows():
    time.sleep(2)
    ntee_code = row[1]['ntee']
    ntee_desc = row[1]['description']

    # Create new record ('Alt^W')
    bot.hotkey('alt', 'w')  # cursor will be in "Code" column
    time.sleep(.2)
    bot.typewrite(ntee_code, interval=.1)   # NTEE/Program Code
    time.sleep(.2)
    gt.tab_then_type(1, ntee_desc)          # Program Description
    time.sleep(.2)
    bot.press('tab')                        # Budget (ignore)
    time.sleep(.2)
    gt.tab_then_type(1, "S")                # Field ([P]rimary or [S]econdary)
    time.sleep(.2)
    bot.press('tab')                        # Active, defaults to Yes
    time.sleep(.2)
    bot.press('tab')                        # To next row, launches save dialog
    time.sleep(1)
    bot.press('Y')                        # Switch from "no" to "Yes"
