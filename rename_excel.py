import pyautogui as pg
import time
import pandas as pd 
import openpyxl


branch_names = pd.read_excel('branches.xlsx')

# counter for loop report
count = 0

for branch in branch_names['branch'].tolist():
    #Choose and Enter Folder -- Manually select first folder
    time.sleep(3)
    pg.press('enter')
    time.sleep(1)
    pg.press('down')
    time.sleep(2)
    pg.rightClick()
    pg.press('m')
    time.sleep(1)
    pg.write(str(branch_names['branch'][count]))
    time.sleep(1)
    pg.press('enter')
    time.sleep(1)
    pg.press('backspace')
    time.sleep(1)
    pg.press('down')
    print(count)
    count +=1

    