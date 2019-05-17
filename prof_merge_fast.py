"""prof_merge_fastpy
Pass 2 column list of FIMS IDs; merge first into the second"""

import pyautogui as bot
import pandas as pd
import xlrd
import time
import gui_tools

screen_res = str(bot.size()[0]) + 'x' + str(bot.size()[1])


def merge_profiles(copy_from_id, copy_to_id):
    """
    :param copy_from_id:
    :param copy_to_id:
    :return:
    """

    # Click File Maint
    #click_item('file_maint')
    bot.press('alt', interval=.1)
    bot.press('f', interval=.1)

    # Click "File Maintenance" -> "Profiles"
    #click_item('profiles')
    bot.press('enter', interval=.1)

    # click_item('combine_2_profiles')
    gui_tools.press_x_times('c', 3, .2)    # Press "C" 3 times

    bot.press('enter', interval=1)
    time.sleep(3)

    # Enter "From", then "To" FIMS IDs to merge
    bot.typewrite(copy_from_id)

    # Tab to next entry
    bot.press('tab', interval=.5)

    bot.typewrite(copy_to_id, interval=.5)

    gui_tools.press_x_times('tab', 2, .25)  # Press 'tab' twice

    bot.press('enter', interval=1)  # Selects "OK" button

    #Do you want to delete copy from?
    bot.press('tab')    # Toggles from "No" selected to "Yes:
    bot.press('enter')  # Clicks/selects the "yes" button

    print('Sleeping for 5 seconds')
    time.sleep(.6)

    # Choose "yes" on the "Ready to Proceed" window
    bot.press('tab')    # Toggles from "No" selected to "Yes:
    bot.press('enter')  # Clicks/selects the "yes" button

    # Wait/check until the window labeled "Message" with text
    # "Profiles Combined!" appears. Then click "OK
    time.sleep(30)
    for i in range(20):
        print('Checking for "profiles combined" box')
        xy = bot.locateCenterOnScreen(
            'img/message_box_profiles_combined_1920x1080.png', grayscale=True)
        if xy is not None:  # Then merge is complete
            bot.press('enter')
            print('Merged ' + copy_from_id + ' into ' + copy_to_id)
            break   # Exit loop
        else:
            time.sleep(30)
        # time.sleep(45)


    # Click "OK"
    bot.press('enter')
    return True

time.sleep(3)

# Read in Excel file of "From->To" mappings
FROM_TO_FILE = r'P:\03_FinanceOperations\InformationManagement\Data Quality\MergeTemplates\to_merge.xlsx'
df_from_to = pd.read_excel(FROM_TO_FILE)

for index, row in df_from_to.iterrows():
    from_id = str(row.From_ID)
    to_id = str(row.To_ID)
    print(f'Merging id {from_id} to id -> {to_id}')
    merge_profiles(from_id, to_id)
