import pyautogui as bot
import time
import re

# Issues with 4k monitors? Per comments at
# https://github.com/asweigart/pyautogui/issues/33
from ctypes import windll
user32 = windll.user32
user32.SetProcessDPIAware()

def click_pic(path_to_photo, sleep = 0, logging = False, region = None):
    """given the filepath to a lossless imagage format (like *.png), find the
    (x,y) coordinates of that image on the screen, and click the mouse there """
    if logging:
        print('Looking for location of ' + path_to_photo)
    if region:
        xy = bot.locateCenterOnScreen(path_to_photo, region=region, grayscale=True)
    else:
        xy = bot.locateCenterOnScreen(path_to_photo, grayscale=True)
    if xy is None:
        print(path_to_photo + ' image not found')
        return None

    else:
        if logging:
            print('Photo ' + path_to_photo + ' located at (' + str(xy[0])
                  + ',' + str(xy[1]))

        bot.click(xy)

        if sleep != 0:
            time.sleep(abs(sleep))

        return True


def screencapt(output_file, region=(0,0, 1920, 1080)):
    """

    :param output_file:
    :param region:
    :return:
    """
    bot.screenshot(output_file, region=region)
    return True

def get_screenres():
    screen_res = str(bot.size()[0]) + 'x' + str(bot.size()[1])
    return screen_res


def track_mouse():
   """
   Utility function to keep running log of current mouse X,Y position
   :return:
   """
   for i in range(1999):
      time.sleep(.1)
      x,y = bot.position()
      print(str(x) + ", " + str(y))


def press_x_times(input_key:str, repetitions:int, interval=.2):
    """
    Press 'key' characters 'repetitions' number of times
    :param input_key: String, key(s) to press
    :param repetitions: Int, number of times to repeat
    :return: None
    """
    for _ in range(repetitions):
        bot.press(input_key)
        time.sleep(interval)

    return None

def tab_then_type(tab_count:int = 1, input_value:str=None):
    """
    Convenience script. Tab "tab_count" times to advance selected cell.
    Then, type in the 'input_value' into selected field
    :param tab_count: Number of times to 'tab' and advance selected input
    :param input_value: Value to paste into selected cell
    :return: None
    """
    press_x_times('tab', tab_count, interval=1)
    if input_value:
        bot.typewrite(input_value)
        print(str(input_value))

    return None


def clean_text(input_text:str)->str:
    """
    Given a text string, clean out all non-printing characters except 'enter'
    :param input_text:
    :return: Cleaned text
    """
    re_chars_and_hard_return = re.compile('[^A-Za-z0-9 !@#$%^&*(){}[]/<>-+=_;:"\',.\n]')
    cleaned_text = re_chars_and_hard_return.sub('',input_text)
    return cleaned_text
