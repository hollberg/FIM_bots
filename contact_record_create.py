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
    :return:
    """

    # Select "View" -> "Contacts" menu path via Alt-V -> "C" -> <enter>
    bot.hotkey('alt', 'v')
    time.sleep(.2)
    bot.typewrite(['c', 'enter'], interval=.5)

    # Wait until Contacts window appears
    new_record_xy = None
    while new_record_xy is None:   # The screen hasn't opened yet, sleep
        time.sleep(2)
        new_record_xy = bot.locateCenterOnScreen(IMG_DIR + 'new_record.png')

    # Click the "New" icon (piece of paper) to launch new contact window
    bot.click(new_record_xy[0], new_record_xy[1])
    time.sleep(2)

    # New Contact record window launches with "Grant" field highlighted.
    # tab twice to get to FIMS ID
    gt.tab_then_type(2, fims_id)

    #Tab 4 times to get to "Tickle Date"
    gt.tab_then_type(4, tickle_date)

    # Tab to get to 'Tickle Who'
    gt.tab_then_type(1, tickle_who)

    # Tab to get to contact date
    gt.tab_then_type(1, contact_date)

    # Tab to get to Next Action (date)
    gt.tab_then_type(1, next_action)

    # Tab once to get to Contact Type
    gt.tab_then_type(1, contact_type)

    # Tab to get to Time
    gt.tab_then_type(1, contact_time)

    # Tab to get to priority
    gt.tab_then_type(1, contact_priority)

    # Tab **3 times** to get to solicitor
    gt.tab_then_type(3, solicitor)

    # Tab to get to staff
    gt.tab_then_type(1, staff)

    # Tab to get to fund
    gt.tab_then_type(1, fund)

    # Tab to get to comment
    gt.tab_then_type(1, contact_comment)

    # Tab to get to grant
    # gt.tab_then_type(1, contact_grant)

    # Click the save (floppy disk) icon
    gt.click_pic(IMG_DIR + 'new_contact_save_icon.png')

    # Wait for "Save (Y/N)" window to appear
    time.sleep(1)

    # Select "File" -> Close Tab via "Alt" -> F -> F -> <enter>
    bot.press(['alt', 'f', 'f', 'enter'], interval=1)
    bot.hotkey('alt', 'f4')

    # Cleanup - Move mouse to lower left corner)
    bot.moveTo(0, bot.size()[1])


def get_contact_data(file_path):
    """
    Read Excel file with contact import data
    :param file_path:
    :return:
    """
    df_contacts = pd.read_excel(file_path)
    # print(df_contacts.columns)
    return df_contacts

def main():
    time.sleep(2)

    contact_filepath = r'P:\03_FinanceOperations\InformationManagement\Databases\FIM_Bots_data'
    contact_file = r'FIM_Bots_contacts_20180926.xlsx'
    df_contacts = get_contact_data(contact_filepath + r'\\' + contact_file)

    for index, row in df_contacts.iterrows():
        fims_id = str(row['FIMSID'])
        staff = row['StaffCode']
        contact_comment = '<Cleanup of Donor "Comment" field - MMH - 20180927>' + '\n' + \
                          gt.clean_text(row['Comment']) #Strip special chars
        contact_type = row['CntctType']
        eff_date = pd.to_datetime(row['EffDate'])
        contact_date = f'{str(eff_date.month)}/{str(eff_date.day)}/{str(eff_date.year)}'


        print(f'preparing to load for {fims_id}')
        # print(contact_comment)

        new_contact(fims_id=fims_id, staff=staff, contact_type=contact_type,
                    contact_date=contact_date, contact_comment=contact_comment)

        time.sleep(2)




    # new_contact(fims_id = '52545', contact_type='COTH', tickle_date='1/1/2018',
    #             solicitor='CMTG', staff='MH', fund='2004',
    #             contact_comment = r"""Now there buddy!
    #             What's your name?
    #
    #             I just skipped two lines.
    #             Now I'm done""")
    #
    # new_contact(fims_id='52545', contact_type='COTH', tickle_date='1/1/2018',
    #             solicitor='CMTG', staff='MH', fund='2004',
    #             contact_comment=r"""Now there buddy!
    #             What's your name?
    #
    #             I just skipped two lines.
    #             Now I'm done""")

main()