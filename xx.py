import pydirectinput as pdi
import time as t
import pandas as p
df = p.read_csv('https://raw.githubusercontent.com/kvmilos/OREX/main/kontrahenci.csv')
dic = dict(zip(df['kod'], df['pozycja']))

def rozr():
    pdi.keyDown('alt')
    pdi.press('r')
    t.sleep(.4) 
    pdi.press('t')
    t.sleep(.4) 
    pdi.press('z')
    t.sleep(.4)
    pdi.keyUp('alt')

def zamien():
    with open ('xx.txt', 'r') as f:
        p = f.read()
        p = p.replace('\n', ' ')

    with open('yy.txt', 'w') as f:
        f.write(p)


def wpisz():
    lista1 = input('Lista nrki i kwoty (albo tylko 1 nr bez kwoty): ')
    if lista1 == 'q':
        return
    lista1 = lista1.replace('-', ' ').replace(';', ' ').replace('(', ' ').replace(')', ' ').replace('/', ' ')
    print(lista1)
    lista = [lista1.split()[i:i+2] for i in range(0, len(lista1.split()), 2)]
    print(lista)
    print("5 sekund na zmianę okna") 
    t.sleep(5) 

    for i, row in enumerate(lista): 
        pdi.write(str(row[1]))
        pdi.press('tab')
        pdi.write('201-2-1-'+str(dic[int(row[0])]))
        if i != len(lista)-1:
            pdi.press('enter')
    
    for i in range(len(lista)-1):
        pdi.press('up')

    pdi.press('tab')
    pdi.press('tab')
    pdi.press('tab')
    pdi.press('tab')
    #35trans(len(lista))
    

def trans(ile=None):
        if ile == None:
            ile = int(input('ile transakcji: '))
        print("5 sekund na zmianę okna")
        t.sleep(5)
        for i in range(ile):
            rozr()
            if i != ile-1:
                pdi.press('down')
            if i == 0:
                pdi.press('tab')

def main():
    co = input('Co? t/n/z \n')
    if co == 't':
        trans()
    elif co == 'n':
        wpisz()
    elif co == 'z':
        zamien()
    else:
        print('nie rozumiem')
    main()

if __name__ == '__main__':
    main()