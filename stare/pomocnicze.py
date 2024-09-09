import pandas as p
import pyautogui
import pydirectinput as pdi
import time as t
import re



def get_dic():
    df = p.read_csv('https://raw.githubusercontent.com/kvmilos/OREX/main/kontrahenci.csv')
    dic = dict(zip(df['kod'], df['pozycja']))
    return dic

def slownik(rez, dic, dlugi=False) -> str:
    if rez != None:
        if not dlugi:
            if dic[int(rez)]:
                return str(dic[int(rez)])
            else:
                return 'Nie ma'
        else:
            if dic[int(rez)]:
                return '201-2-1-'+str(dic[int(rez)])
            else:
                return 'Nie ma'
    else:
        return 'Nie ma'
        

def myszka():
    print('Press any key to quit')
    try:
        while True:
            x, y = pyautogui.position()
            positionStr = 'X: ' + str(x).rjust(4) + ' Y: ' + str(y).rjust(4)
            print(positionStr, end='')
            print('\b' * len(positionStr), end='', flush=True)
    except KeyboardInterrupt:
        print('\n')


def rozr():
    pdi.keyDown('alt')
    pdi.press('r')
    t.sleep(.2)
    pdi.press('t')
    t.sleep(.2) 
    pdi.press('z')
    t.sleep(.2)
    pdi.keyUp('alt')


def find_reservation(text, spaces=False):
    if not spaces:
        pattern = r'1[012]\d{5}'
    else:
        pattern = r'1\s*[012]\s*\d{5}'
    if re.findall(pattern, text):
        return str(re.findall(pattern, text)[0])
    else:
        return None
    

myszka()