import time as t
import pyautogui as pdi

def rozr():
    pdi.press('up')
    pdi.press('up')
    pdi.press('up')
    t.sleep(0.2)
    for _ in range(4):
        pdi.keyDown('alt')
        pdi.press('r')
        pdi.keyUp('alt')
        t.sleep(0.1)
        pdi.press('down')
        pdi.press('enter')
        t.sleep(0.1)

def wpisz():
    ile = int(input("Ile wpisów? "))
    print("5 sekund na zmianę okna") 
    t.sleep(5) 
    for _ in range(ile):
        pdi.press('enter')
        t.sleep(2)
        pdi.click(3251, 385)
        t.sleep(0.2)
        pdi.keyDown('alt')
        pdi.press('r')
        pdi.keyUp('alt')
        t.sleep(2)
        rozr()
        t.sleep(0.1)
        pdi.keyDown('alt')
        pdi.press('z')
        pdi.keyUp('alt')
        t.sleep(1)
        pdi.press('enter')
        t.sleep(0.5)
        pdi.press('enter')
        t.sleep(1)
        pdi.keyDown('alt')
        pdi.press('b')
        pdi.keyUp('alt')
        t.sleep(3)
        pdi.press('down')
        t.sleep(0.5)



def main():
    wpisz()


if __name__ == "__main__":
    main()