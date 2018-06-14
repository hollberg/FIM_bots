import pyautogui as bot
import time

# Issues with 4k monitors? Per comments at
# https://github.com/asweigart/pyautogui/issues/33
from ctypes import windll
user32 = windll.user32
user32.SetProcessDPIAware()

def click_pic(path_to_photo, sleep = 0, logging = False):
    """given the filepath to a lossless imagage format (like *.png), find the
    (x,y) coordinates of that image on the screen, and click the mouse there """
    print('Looking for location of ' + path_to_photo)
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

