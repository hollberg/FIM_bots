"""changeGL.py
Automate remapping a FIMS GL Account code via
Tools -> System Utilities -> Admin Utilities -> Finance Utilities -> Change Natural Account
"""

import pyautogui as bot
import pandas as pd
import xlrd
import time
import gui_tools as gt
import yrdata

def map_gl(fisc_yr, gl_old, gl_new):

    time.sleep(1)

    pics_to_click = [
        'menu_tools_highlighted.png',
        'menu_tools_systemutilities.png',
        'menu_tools_systemutilities_adminutilities.png',
        'menu_tools_systemutilities_adminutilities_financeUtilities.png',
        'menu_tools_systemutilities_adminutilities_financeUtilities_ChangeNaturalAccount.png'
        ]

    photo_folder = 'img/GLChng/'

    # Tools -> System Utilities -> Admin Utilities -> Finance Utilities -> Change Natural Account
    for pic in pics_to_click:
        path = photo_folder + pic
        gt.click_pic(path, sleep=1, logging=True)

    time.sleep(1)

    # # Confirm message box appears
    # msg_box_xy = bot.locateOnScreen(photo_folder +
    #                                 'message_box_ChangeNaturalAccountNumber.png')
    #
    #
    # if msg_box_xy is None:  #Box not found!
    #     print('***ERROR*** Message box did not appear')
    #
    # time.sleep(1)

    # Click the "OK" button on the message box
    bot.press('enter')

    time.sleep(1)

    # Enter data in "Global Change for GL Natural Account" window
    # steps:    1) Enter "Old" Account number; /tab
    #           2) Enter "New" Account number; /tab*3
    #           3) Select FY: Type number "2" n times (once = 2000, twice=2001...)
    #           4) /tab*3, type "Enter" (Selects "OK" button)

    bot.typewrite(gl_old)
    bot.press('tab')

    time.sleep(1)

    bot.typewrite(gl_new)
    # Tab 3 times to select GL Year option
    bot.press('tab')
    bot.press('tab')
    bot.press('tab')

    time.sleep(1)

    # Enter "2" enough times to cycle to desired fiscal yr
    number_of_2s = fisc_yr - 1999
    for _ in range(number_of_2s):
        bot.press('2', pause=.2)

    time.sleep(1)

    # Tab 3 times to select "Proceed" button
    bot.press('tab')
    bot.press('tab')
    bot.press('tab')

    # "Proceed" button is highlighted; hit "enter" to run process
    bot.press('enter')

    # Process will run for a long time. Check to see when completion message box
    # "Global Change for GL Natural Account" window appears - then click "Enter"
    for i in range(30):
        print('Cycle Number ' + str(i))
        time.sleep(30)

        gl_error_box_xy = \
            bot.locateOnScreen(photo_folder +
                               'Error_box.png')

        if gl_error_box_xy is not None:
            bot.press('enter')

        done_msg_box_xy = \
            bot.locateOnScreen(photo_folder +
                               'message_box_Global_Change_for_GL_nat_account.png')

        if done_msg_box_xy is not None: #box appeared
            bot.press('enter')  # Selects "OK" button on window
            break # Exit for loop

for entry in yrdata.mapping:
    gl_yr = entry[0]
    old_gl = entry[1]
    new_gl = entry[2]
    map_gl(fisc_yr=gl_yr, gl_old=old_gl, gl_new=new_gl)