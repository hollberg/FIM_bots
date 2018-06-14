"""profile_merge.py
Pass 2 column list of FIMS IDs; merge first into the second"""

import pyautogui as bot
import pandas as pd
import xlrd
import time
import gui_tools as gt

def merge_profiles(copy_from_id, copy_to_id):
    """

    :param copy_from_id:
    :param copy_to_id:
    :return:
    """

    # Click File Maint
    if gt.click_pic(r'img/fims_menu_File_Maintenance.png', sleep=1,
                    logging=True) is None:
        gt.click_pic(r'img/fims_menu_File_Maintenance_selected.png', sleep=1,
                     logging=True)

    # Click "File Maintenance" -> "Profiles"
    gt.click_pic(r'img/fims_file_maint_profiles.png', sleep=1,
                 logging=True)

    # Click "Combine 2 Profiles"
    gt.click_pic(r'img/fims_filemaint_profile_combine2.png', sleep=1,
                 logging=True)

    bot.typewrite(copy_from_id)

    # Tab to next entry
    bot.press('tab')

    time.sleep(.5)

    bot.typewrite(copy_to_id)
    time.sleep(.5)

    bot.press('tab')
    bot.press('tab')
    time.sleep(.5)

    bot.press('enter')  # Selects "OK" button

    time.sleep(1)

    #Do you want to delete copy from?
    bot.press('tab')    # Toggles from "No" selected to "Yes:
    bot.press('enter')  # Clicks/selects the "yes" button

    print('Sleeping for 5 seconds')
    time.sleep(.5)

    # Choose "yes" on the "Ready to Proceed" window
    bot.press('tab')    # Toggles from "No" selected to "Yes:
    bot.press('enter')  # Clicks/selects the "yes" button

    # Wait/check until the window labeled "Message" with text
    # "Profiles Combined!" appears. Then click "OK
    time.sleep(30)
    for i in range(5):
        print('Checking for "profiles combined" box')
        xy = bot.locateCenterOnScreen(
            'img/message_box_profiles_combined.png', grayscale=True,
             region = (0,0, 3840, 2160))
        if xy is not None:  # Then merge is complete
            bot.press('enter')
            print('Merged ' + copy_from_id + ' into ' + copy_to_id)
            return True
            # break   # Exit loop
        else:
            time.sleep(30)


    # Click "OK"
    bot.press('enter')
    return True




time.sleep(5)

merge_to = '31863'
copy_froms = ['28073', '47839', '34741', '31553']
for copy in copy_froms:
    merge_profiles(copy, merge_to)
