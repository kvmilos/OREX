# this will be a clicker
# in a loop, it should click mouse on 750, 150, then wait 2 seconds, click F9, wait 2 seconds, click T, wait 2 seconds, click Space, wait 2 seconds
# if it's interrupted by Escape, it should print 'Quitting' and break the loop

import pyautogui
import time as t

def clicker():
    print('Press ESC to quit')
    print('Clicking in 3 seconds')
    t.sleep(3)
    try:
        while True:
            t.sleep(1)
            if pyautogui.press('esc'):
                print('Quitting')
                break
            pyautogui.click(850, 150)
            t.sleep(2)
            pyautogui.click(750, 200)
            t.sleep(2)
            pyautogui.press('t')
            t.sleep(4) 
            pyautogui.press('space')  
            t.sleep(1)
    except KeyboardInterrupt:
        print('Quitting')

clicker()