"""contact_record_create.py
Read data from external file; build a contact record in FIMS
Follow toolbar path:
View -> Contacts ->
"""
#Imports
# Standard Library
import time

# Local Modules
import gui_tools as gt

# External Libraries
import pyautogui as bot
import pandas as pd

IMG_DIR = 'img/contact_record_create/'

def new_contact(fims_id:str,
                tickle_date:str = None, tickle_who:str = None,
                contact_date:str = None, next_action:str = None,
                contact_type:str = None, contact_time:str = None,
                contact_priority:int = None, solicitor:str = None,
                staff:str = None, fund:str = None,
                contact_comment:str = None, contact_grant:str = None):
    """
    From a full-screen FIMS launch screen, open and populate a new Contact record
    :return: None
    """
    # Select "View"      -> "Contacts" menu path via Alt-V -> "C" -> <enter>
    # bot.hotkey('alt', 'v')
    # time.sleep(.2)
        # bot.typewrite(['c', 'enter'], interval=.5)
    # Launch "Contacts" window - shortcut ctrl^alt^T
    # bot.hotkey('ctrl', 'alt', 't')

    # See if contact window has launched


    # Wait until Contacts window appears
    # TODO - reinstate code below, but get updated
    # new_record_xy = None
    # while new_recor   d_xy is None:   # The screen hasn't opened yet, sleep
    #     time.sleep(2)
    #     new_record_xy = bot.locateCenterOnScreen(IMG_DIR + 'new_record.png')
    time.sleep(10)

    # Assume contacts window is a pop-under; alt-tab to view it
    #bot.hotkey('alt', 'tab', interval=0.1)

    # Create new record ('Alt^W')
    print('About to alt W')
    bot.hotkey('alt', 'W')
    # Move mouse and click to activate window
    # bot.click(400,400)
    # # Click the "New" icon (piece of paper) to launch new contact window
    # bot.click(new_record_xy[0], new_record_xy[1])
    # time.sleep(2)

    # New Contact record window launches with "Grant" field highlighted.
    gt.tab_then_type(2, fims_id)    # tab twice to get to FIMS ID
    gt.tab_then_type(4, tickle_date)    # Tab 4 times to get to "Tickle Date"
    gt.tab_then_type(1, tickle_who) # Tab to get to 'Tickle Who'
    gt.tab_then_type(1, contact_date)   # Tab to get to contact date
    gt.tab_then_type(1, next_action)    # Tab to get to Next Action (date)
    gt.tab_then_type(1, contact_type)   # Tab once to get to Contact Type
    gt.tab_then_type(1, contact_time)   # Tab to get to Time
    gt.tab_then_type(1, contact_priority)   # Tab to get to priority
    gt.tab_then_type(3, solicitor)  # Tab **3 times** to get to solicitor
    gt.tab_then_type(1, staff)  # Tab to get to staff
    gt.tab_then_type(1, fund)   # Tab to get to fund
    gt.tab_then_type(1, contact_comment)    # Tab to get to comment
    # gt.tab_then_type(1, contact_grant)    # Tab to get to grant

    # # Click the save (floppy disk) icon
    # gt.click_pic(IMG_DIR + 'new_contact_save_icon.png')

    print('About to save')
    bot.hotkey('alt', 'S')  # Save keyboard shortcut
    time.sleep(3)

    #print('About to alt f4')
    # Select "File" -> Close Tab via "Alt" -> F -> F -> <Alt-F4>
    # bot.press(['alt', 'f', 'f', 'enter'], interval=1)
    # bot.hotkey('alt', 'f4')

    # Cleanup - Move mouse to lower left corner)
    # bot.moveTo(0, bot.size()[1])
    return None


def main():
    time.sleep(2)

    contact_filepath = r'P:\03_FinanceOperations\InformationManagement\Databases\FIM_Bots_data'
    contact_file = r'FIM_Bots_contacts_Spark_1_20190514.xlsx'
    df_contacts = pd.read_excel(contact_filepath + r'\\' + contact_file)
    df_contacts.fillna('', inplace=True)

    # Launch "Contacts" window - shortcut ctrl ^ alt ^ T
    bot.hotkey('ctrl', 'alt', 't')
    time.sleep(30)

    for index, row in df_contacts.iterrows():
        fims_id = str(row['FIMSID'])
        staff = row['StaffCode']
        contact_comment = gt.clean_text(row['Comment']) + '\n' \
        + '<Bulk entry 20190516>'

        contact_type = row['CntctType']
        eff_date = pd.to_datetime(row['EffDate'])
        contact_date = f'{str(eff_date.month)}/{str(eff_date.day)}/{str(eff_date.year)}'

        print(f'preparing to'
              f' load for {fims_id}')

        new_contact(fims_id=fims_id, staff=staff, contact_type=contact_type,
                    contact_date=contact_date, contact_comment=contact_comment)

        time.sleep(2)


main()
