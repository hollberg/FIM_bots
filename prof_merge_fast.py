"""prof_merge_fastpy
Pass 2 column list of FIMS IDs; merge first into the second"""

import pyautogui as bot
import pandas as pd
import xlrd
import time
import gui_tools as gt

screen_res = str(bot.size()[0]) + 'x' + str(bot.size()[1])

# Dict of dicts, storing (x,y) location of menu item for a given screen
# resolution (assumes FIMS is running at full screen).
# Format is: {'MENU LOCATION': {ScreenResolution: (X,Y)}}
menu_locations = {'tools': {'1366x768': (420,32)},
                  'system_utils': {'1366x768': (484,181)},
                  'admin_utils': {'1366x768': (785,609)},
                  'fin_utils': {'1366x768': (103,665)},
                  'change_net_account': {'1366x768': (786, 503)},
                  'file_maint': {'1366x768': (353, 33)},
                  'profiles': {'1366x768': (349,57)},
                  'combine_2_profiles': {'1366x768': (579,325)},
                  }

def click_item(menu_item, screen_res = screen_res, menu_locations=menu_locations):
    """
    Given a menu item name, click the mouse at the x,y coords
    :param menu_item:
    :return:
    """
    x,y = menu_locations[menu_item][screen_res]
    bot.click(x,y)
    time.sleep(.5)
    return True

def merge_profiles(copy_from_id, copy_to_id):
    """

    :param copy_from_id:
    :param copy_to_id:
    :return:
    """

    # Click File Maint
    click_item('file_maint')

    # Click "File Maintenance" -> "Profiles"
    click_item('profiles')

    # Click "Combine 2 Profiles"
    click_item('combine_2_profiles')

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
    time.sleep(.6)

    # Choose "yes" on the "Ready to Proceed" window
    bot.press('tab')    # Toggles from "No" selected to "Yes:
    bot.press('enter')  # Clicks/selects the "yes" button

    # Wait/check until the window labeled "Message" with text
    # "Profiles Combined!" appears. Then click "OK
    time.sleep(30)
    for i in range(5):
        print('Checking for "profiles combined" box')
        xy = bot.locateCenterOnScreen(
            'img/message_box_profiles_combined_1366x768.png', grayscale=True,
             region = (609, 320, 761, 455))
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

from_tos = [('43871', '28799'),
     ]


for from_to in from_tos:
    merge_from, merge_to = from_to
    print(f'merge {merge_from} to {merge_to}')
    merge_profiles(merge_from, merge_to)

# merge_to = '39432'
# copy_froms = ['44570']
# for copy in copy_froms:
#     merge_profiles(copy, merge_to)
