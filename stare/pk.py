import pandas as p
import pydirectinput as pdi
import time as t

df = p.read_csv('https://raw.githubusercontent.com/kvmilos/OREX/main/kontrahenci.csv')
dic = dict(zip(df['kod'], df['pozycja']))

def pk():
    df2 = p.read_csv('pk/pk.csv', sep=';', header=None)
    df2 = df2[[0, 1]]
    df2.columns = ['rez', 'kwota']

    print('5 sekund na zmianę okna')
    t.sleep(5)

    for _, row in df2.iterrows():
        pdi.press('Enter')
        t.sleep(0.1)
        pdi.press('left')
        t.sleep(0.1)
        pdi.press('right')
        t.sleep(0.1)
        pdi.press('backspace')
        t.sleep(0.1)
        pdi.press('backspace')
        t.sleep(0.1)
        pdi.press('backspace')
        t.sleep(0.1)
        pdi.press('backspace')
        t.sleep(0.1)
        pdi.press('backspace')
        t.sleep(0.1)
        pdi.press('backspace')
        t.sleep(0.1)
        pdi.press('backspace')
        t.sleep(0.1)
        pdi.write(str(row['rez']))
        t.sleep(0.1)
        pdi.press('Enter')
        t.sleep(0.1)
        pdi.press('Enter')
        t.sleep(0.1)
        pdi.write(str(row['kwota']))
        t.sleep(0.1)
        pdi.press('Enter')
        t.sleep(0.1)
        pdi.write('201-2-1-' + str(dic[int(row['rez'])]))
        t.sleep(0.1)
        pdi.press('Enter')
        t.sleep(0.1)
        pdi.press('Enter')
        t.sleep(0.1)
        pdi.write('763-7')


def nry():
    df2 = p.read_csv('pk/x.csv', header=None)
    print('5 sekund na zmianę okna')
    t.sleep(5)
    for _, row in df2.iterrows():
        pdi.press('Space')
        t.sleep(0.05)
        pdi.write(str(dic[int(row[0])]))
        t.sleep(0.05)
        pdi.press('down')

def rozr():
    start = int(input('Od jakiego numeru? '))
    koniec = int(input('Do jakiego numeru? '))
    ile = koniec - start + 1
    print('5 sekund na zmianę okna')
    t.sleep(5)
    for _ in range(ile):
        pdi.keyDown('alt')
        t.sleep(0.2)
        pdi.press('r')
        t.sleep(0.2)
        pdi.press('t')
        t.sleep(0.2)
        pdi.press('z')
        t.sleep(0.2)
        pdi.keyUp('alt')
        t.sleep(0.2)
        pdi.press('down')
        t.sleep(0.2)




def main():
    rozr()

if __name__ == '__main__':
    main() 