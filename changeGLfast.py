"""changeGL.py
Automate remapping a FIMS GL Account code via
Tools -> System Utilities -> Admin Utilities -> Finance Utilities -> Change Natural Account
"""

import pyautogui as bot
import pandas as pd
import xlrd
import time
import yrdata

bot.FAILSAFE = False # disables the fail-safe

screen_res = str(bot.size()[0]) + 'x' + str(bot.size()[1])

# Dict of dicts, storing (x,y) location of menu item for a given screen
# resolution (assumes FIMS is running at full screen).
# Format is: {'MENU LOCATION': {ScreenResolution: (X,Y)}}
menu_locations = {'tools': {'1366x768': (420,32)},
                  'system_utils': {'1366x768': (493,189)},
                  'admin_utils': {'1366x768': (789,595)},
                  'fin_utils': {'1366x768': (1044,665)},
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
    time.sleep(1)
    print('clicked ' + menu_item)
    return True


def map_gl(fisc_yr, gl_old, gl_new):

    time.sleep(1)

    click_item('tools')
    click_item('system_utils')
    click_item('admin_utils')
    click_item('fin_utils')
    click_item('change_net_account')

    time.sleep(1)

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

    err_count = 0
    # Process will run for a long time. Check to see when completion message box
    # "Global Change for GL Natural Account" window appears - then click "Enter"
    for i in range(1000):
        print('Cycle Number ' + str(i))
        time.sleep(5)

        # Check to see if an error box pops up; click "OK" if it does
        gl_error_box_xy = \
            bot.locateOnScreen('img/' +
                               'error_global_change_for_gl.png',
                               region=(479,307, 888,467))

        if gl_error_box_xy is not None:
            print('Error found converting GL Account ' + gl_old)
            bot.screenshot('img/errors/' + gl_old + '_' + str(err_count) + '.png')
            err_count += 1
            bot.press('enter')

        done_msg_box_xy = \
            bot.locateOnScreen('img/message_global_change_done_1366x768.png',
                               region=(473,307, 888,467))

        if done_msg_box_xy is not None: #box appeared
            bot.press('enter')  # Selects "OK" button on window
            print('Completed conversion of GL Account ' + gl_old + ' to ' + gl_new)
            break # Exit for loop


for entry in yrdata.mapping:
    gl_yr, old_gl, new_gl = entry
    # gl_yr = entry[0]
    # old_gl = entry[1]
    # new_gl = entry[2]
    map_gl(fisc_yr=gl_yr, gl_old=old_gl, gl_new=new_gl)