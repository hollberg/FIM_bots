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
    gt.click_pic(IMG_DIR + 'new_record.png')

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
    gt.tab_then_type(1, contact_grant)

    # Click the save (floppy disk) icon
    gt.click_pic(IMG_DIR + 'new_contact_save_icon.png')

    # Wait for "Save (Y/N)" window to appear
    time.sleep(1)

    # Select "File" -> Close Tab via "Alt" -> F -> F -> <enter>
    bot.press(['alt', 'f', 'f', 'enter', 'enter'], interval=.4)

    # Cleanup - Move mouse to lower left corner)
    bot.moveTo(0, bot.size()[1])


def main():
    time.sleep(2)
    new_contact(fims_id = '52545', contact_type='COTH', tickle_date='1/1/2018',
                solicitor='CMTG', staff='MH', fund='2004',
                contact_comment = r"""Now there buddy!
                What's your name?
                
                I just skipped two lines.
                Now I'm done""")

    new_contact(fims_id='52545', contact_type='COTH', tickle_date='1/1/2018',
                solicitor='CMTG', staff='MH', fund='2004',
                contact_comment=r"""Now there buddy!
                What's your name?

                I just skipped two lines.
                Now I'm done""")

main()