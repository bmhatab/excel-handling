import pyautogui as pg
import time
import pandas as pd 
import openpyxl


#branch_names = pd.read_excel('branches.xlsx')

# counter for loop report
count = 0
time.sleep(3)
for file in range(1,35):
    time.sleep(4)
    #Choose and Enter Folder -- Manually select first folder
    
    pg.press('enter')
    time.sleep(10)
    pg.press('enter')
    time.sleep(5)
    pg.press('enter')
    time.sleep(5)

    #Enable editing
    pg.moveTo(819,71,1)
    pg.leftClick()
    time.sleep(2)

    #Inside excel file moveto file
    pg.moveTo(25,37,1)
    pg.leftClick()
    time.sleep(1)

    #Move to save as
    pg.moveTo(45,237,1)
    pg.leftClick()
    time.sleep(1)

    #Move to current folder
    pg.moveTo(1025,224,1)
    pg.doubleClick()
    pg.moveTo(63,11)
    time.sleep(2)

    #Move to save as type
    pg.moveTo(658,365,1)
    pg.leftClick()
    pg.moveTo(287,7,1)
    pg.leftClick()
    time.sleep(1)
    pg.press('enter')

    #Quit excel
    time.sleep(4)
    pg.press('enter')
    pg.moveTo(1347,9)
    pg.leftClick()
    time.sleep(2)

    #Back to folder
    pg.moveTo(318,11,1)
    pg.leftClick()

    #Delete xlx file
    time.sleep(2)
    pg.hotkey('shift','del')
    time.sleep(1)
    pg.press('enter')

    print(count)
    
    #Go to next
    time.sleep(2)
    pg.press('down')




    
    