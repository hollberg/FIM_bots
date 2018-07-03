import pyautogui as bot
import time
from collections import namedtuple

# resolution_location = namedtuple(['resolution, x_loc, y_loc'])

# x, y = bot.size()
# print(str(x) + ", " + str(y))

for i in range(2000):
   time.sleep(.1)
   x,y = bot.position()
   print(str(x) + ", " + str(y))


