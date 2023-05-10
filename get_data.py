import pyautogui as pg
import time
import pandas as pd 
import openpyxl


branch_names = pd.read_excel('branches.xlsx')

# counter for loop report
count = 0

for branch in branch_names['branch'].tolist():
    # Goto branch drop down menu
    time.sleep(5)
    pg.moveTo(822,333,1)
    time.sleep(2)

    # Select 
    pg.leftClick()
    time.sleep(2)

    # Select the down key 
    pg.press('down')
    time.sleep(2)
    pg.press('enter')

    # Goto export
    pg.moveTo(1151,668,1)
    pg.leftClick()
    time.sleep(2)
    pg.moveTo(1125,683,1)
    pg.leftClick()

    # Scroll to end
    pg.moveTo(572,331,1)
    pg.leftClick()
    time.sleep(2)
    pg.scroll(-20) # 20 clicks down 

    # Select New  folder 
    time.sleep(2)
    pg.moveTo(608,492,1)
    pg.leftClick()
    time.sleep(2)

    # Goto make new folder 
    pg.moveTo(581,530,1)
    pg.leftClick()
    time.sleep(1)

    # Rename folder 
    pg.write(str(branch_names['branch'][count]))
    time.sleep(1)
    pg.press('enter')
    time.sleep(1)

    # Goto OK
    pg.moveTo(709,530,1)
    pg.leftClick()
    time.sleep(55)

    # Escape and quit excel 
    pg.press('enter')
    pg.moveTo(1338,83,1)
    time.sleep(2)
    pg.press('right')
    pg.press('enter')
    pg.leftClick()
    time.sleep(4)

    # Status report 
    print(count)
    count += 1