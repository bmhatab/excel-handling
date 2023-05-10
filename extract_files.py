import pyautogui as pg
import time
import pandas as pd 
import openpyxl

time.sleep(5)

for i in range(1,35):

    pg.press('enter')
    time.sleep(2)
    pg.pres('down')
    time.sleep(2)
    pg.hotkeys('ctrl','x')
    time.sleep(2)
    pg.press('backspace')
    pg.hotkeys('shift','del')
    time.sleep(2)
    pg.press('enter')
    time.sleep(2)
    pg.hotkeys('ctrl','v')
    time.sleep(2)
    pg.press('down')
    print(i)


