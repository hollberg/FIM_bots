"""profile_merge.py
Pass 2 column list of FIMS IDs; merge first into the second"""

import pyautogui as bot
import pandas as pd
import xlrd
import time

region = (0,0, 1920,1080)

# Capture screenshots
##img_fims_home = bot.screenshot('fims_home.png', region=region)

# Click File Maint, make screenshot
file_maint_xy = bot.locateCenterOnScreen('fims_menu_File_Maintenance.png')
bot.click(file_maint_xy)
time.sleep(1)

# Get Screenshot of "File Maintenance" -> "Profiles"
##img_fims_file_maint_selected = bot.screenshot('fims_file_maint_selected.png', region=region)

# Click "File Maintenance" -> "Profiles"
file_maint_profiles_xy = bot.locateCenterOnScreen('fims_file_maint_profiles.png')
bot.click(file_maint_profiles_xy)
time.sleep(1)
# Get Screenshot of FIMS Main Menu; File Maintenance -> Profiles
# img_fims_filemaint_profiles_selected = bot.screenshot(
#  'fims_filemaint_profile_selected.png', region = region)

# Click "Combine 2 Profiles"
combine_profiles_xy = bot.locateCenterOnScreen('fims_filemaint_profile_combine2.png')
bot.click(combine_profiles_xy)
time.sleep(1)

# Enter copy from profile
copy_from_xy = bot.locateCenterOnScreen('fims_combine2_profile_copy_from.png')
copy_to_xy = bot.locateCenterOnScreen('fims_combine2_profile_copy_to.png')

one = '47229'
two = '31863'

#bot.click(copy_from_xy)
bot.typewrite(one)

# Tab to next entry
bot.press('tab')
time.sleep(.5)

#bot.click(copy_to_xy)
bot.typewrite(two)
time.sleep(.5)
#
# ok_button_xy = bot.locateCenterOnScreen('fims_button_ok.png')
# bot.click(ok_button_xy)
bot.press('tab')
bot.press('tab')
time.sleep(.5)

bot.press('enter')  # Selects "OK" button

time.sleep(1)

#Do you want to delete copy from?
bot.press('tab')    # Toggles from "No" selected to "Yes:
bot.press('enter')  # Clicks/selects the "yes" button

time.sleep(.5)

# Choose "yes" on the "Ready to Proceed" window
bot.press('tab')    # Toggles from "No" selected to "Yes:
bot.press('enter')  # Clicks/selects the "yes" button


# Wait/check until the window labeled "Message" with text
# "Profiles Combined!" appears. Then click "OK

#
# yes_button_xy = bot.locateCenterOnScreen('fims_button_yes.png')
# bot.click(yes_button_xy)
# time.sleep(1)
#
# # Prompts "Question" box. "Do you want to Delete "Copy From" Profile Record? (Y/N)
# new_yes_xy = bot.locateCenterOnScreen('fims_button_yes.png')
# bot.click(new_yes_xy)
# time.sleep(1)
#
# # Propts "Ready to Proceed?" window. "About to combine Profile XXX into YYY! Do
# # you really want to Combine these Profiles?"
# even_more_yes_xy = bot.locateCenterOnScreen('fims_button_yes.png')
# bot.click(even_more_yes_xy)
#
#
#
