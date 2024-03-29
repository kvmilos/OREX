from pomocnicze import slownik, find_reservation, rozr
import pandas as p
import pydirectinput as pdi
import time as t
import os
import re
from tkinter import Tk
import random

BNP_OPIS = r'(\^20)(.*?\n?.*?\n?.*?\n?.*?\n?.*?)\^(32)'
PATTERN1 = r'(^|\b|\D|00000)(1 *[012] *\d *\d *\d *\d *\d)(\D|\b)'
BNP_PATTERN = r'(^|\D|00000)(1\s*[012]\s*\d\s*\d\s*\d\s*\d\s*\d)(\b|\D)'
BNP_PATTERN2 = r'(\b|\D|00000)(1\s*[012]\s*(\d\s*){5})(\D)+((\d\s*){1,5}(( ?\. ?|,)\d\d?)?)(\D|\b)'
SANTANDER_OPIS = r'\?20.*?\n?.*?\n?.*?\n?.*?\?\s*3\s*[12]'
SANTANDER_PATTERN = r'(1\s*[012]\s*\d\s*\d\s*\d\s*\d\s*\d)(\b|\D)'
SANTANDER_PATTERN2 = r'(1\s*[012]\s*(\d\s*){5})(\D)+((\d\s*){1,5}((\.|,)\d\d?)?)(\D|\b)'
url = 'https://raw.githubusercontent.com/kvmilos/OREX/main/kontrahenci.csv'
url += '?random=' + str(random.randint(1, 1000000))
df = p.read_csv(url)
dic = dict(zip(df['kod'], df['pozycja']))

def przeksiegowanie(plik):
    df2 = p.read_excel(plik)
    print('5 sekund na zmianę okna')
    t.sleep(5)
    for i in range(len(df2)):
        if i != 0:
            pdi.press('enter')
        t.sleep(0.05)
        pdi.write(str(df2['from'][i]) + ' na ' + str(df2['to'][i]))
        pdi.press('enter')
        t.sleep(0.05)
        pdi.press('enter')
        t.sleep(0.05)
        pdi.write(str(df2['amount'][i]))
        t.sleep(0.05)
        pdi.press('enter')
        t.sleep(0.05)
        pdi.write(slownik(df2['from'][i], dic, dlugi = True))
        t.sleep(0.05)
        pdi.press('enter')
        t.sleep(0.05)
        pdi.press('enter')
        t.sleep(0.05)
        pdi.write(slownik(df2['to'][i], dic, dlugi = True))
        if i%50 == 0:
            pdi.keyDown('ctrl')
            pdi.press('s')
            pdi.keyUp('ctrl')
            t.sleep(60)


def pk_prawy(plik):
    df = p.read_excel(plik, header=None)
    df2 = p.DataFrame()
    df2['Kwota'] = df[0]
    df2['Rezerwacja'] = df[1]
    df2['Konto'] = df2.apply(lambda x: slownik(x['Rezerwacja'], dic, dlugi = True), axis=1)
    df2 = df2.astype(str)
    print(sum([float(x.replace(' ', '')) for x in df2['Kwota']]))
    print('5 sekund na zmianę okna')
    t.sleep(5)
    for i, row in df2.iterrows():
        pdi.press('enter')
        t.sleep(0.05)
        pdi.press('enter')
        t.sleep(0.05)
        pdi.write(row['Kwota'])
        t.sleep(0.05)
        pdi.press('enter')
        t.sleep(0.05)
        pdi.write('764-98')
        t.sleep(0.05)
        pdi.press('enter')
        t.sleep(0.05)
        pdi.press('enter')
        t.sleep(0.05)
        pdi.write(row['Konto'])
        if i != len(df2)-1:
            pdi.press('enter')

def pk_lewy(plik):
    df = p.read_excel(plik, header=None)
    df2 = p.DataFrame()
    df2['Kwota'] = df[0]
    df2['Rezerwacja'] = df[1]
    df2['Konto'] = df2.apply(lambda x: slownik(x['Rezerwacja'], dic, dlugi = True), axis=1)
    df2 = df2.astype(str)
    print(sum([float(x.replace(' ', '')) for x in df2['Kwota']]))
    print('5 sekund na zmianę okna')
    t.sleep(5)
    for i, row in df2.iterrows():
        pdi.press('enter')
        t.sleep(0.05)
        pdi.press('enter')
        t.sleep(0.05)
        pdi.write(row['Kwota'])
        t.sleep(0.05)
        pdi.press('enter')
        t.sleep(0.05)
        pdi.write(row['Konto'])
        t.sleep(0.05)
        pdi.press('enter')
        t.sleep(0.05)
        pdi.press('enter')
        t.sleep(0.05)
        pdi.write('763-98')
        if i != len(df2)-1:
            pdi.press('enter')

def anulacje(plik):
    df2 = p.read_excel(plik)
    print('5 sekund na zmianę okna')
    t.sleep(5)
    for i in range(len(df2)):
        pdi.press('enter')
        t.sleep(0.05)
        pdi.press('left')
        t.sleep(0.05)
        pdi.press('right')
        t.sleep(0.05)
        for _ in range(7):
            pdi.press('backspace')
        t.sleep(0.05)
        pdi.write(str(df2['number'][i]))
        pdi.press('enter')
        t.sleep(0.05)
        pdi.press('enter')
        t.sleep(0.05)
        pdi.write(str(df2['amount'][i]))
        t.sleep(0.05)
        pdi.press('enter')
        t.sleep(0.05)
        if df2['number'][i] in dic:
            pdi.write(slownik(df2['number'][i], dic, dlugi = True))
        else:
            pdi.write('139-5')
        t.sleep(0.05)
        pdi.press('enter')
        t.sleep(0.05)
        pdi.press('enter')
        t.sleep(0.05)
        pdi.write('763-7')
        t.sleep(0.05)

        

def zamien_przelewy(plik):
    df = p.read_csv(plik, sep=',')
    df2 = p.DataFrame()
    df2['Kwota'] = df['Kwota']/100
    df2['Rezerwacja'] = df.apply(lambda x: find_reservation(x['Opis']), axis=1)
    df2['Konto'] = df2.apply(lambda x: slownik(x['Rezerwacja'], dic, dlugi = True), axis=1)
    plik2 = plik.replace('.csv', '_zmienione.csv')
    df2.to_csv(plik2, index=False)


def fvzp(plik):
    df = p.read_excel(plik, header=None)
    df2 = p.DataFrame()
    df2['Kwota'] = df[1]
    df2['Rezerwacja'] = df[0]
    df2['Konto'] = df2.apply(lambda x: slownik(x['Rezerwacja'], dic, dlugi = True), axis=1)
    df2 = df2.astype(str)
    print('Suma kwot to:', sum([float(x.replace(' ', '')) for x in df2['Kwota']]))
    print('5 sekund na zmianę okna')
    t.sleep(5)
    for _, row in df2.iterrows():
        pdi.write(row['Kwota'])
        t.sleep(0.05)
        pdi.press('Enter')
        t.sleep(0.05)
        pdi.write(row['Konto'])
        t.sleep(0.05)
        pdi.press('Enter')


def zamien_faktury(plik):
    df = p.read_excel(plik, header=None)
    df2 = p.DataFrame()
    df2['Kwota'] = df[1]
    df2['Rezerwacja'] = df[0]
    df2['Konto'] = df2.apply(lambda x: slownik(x['Rezerwacja'], dic, dlugi = True), axis=1)
    plik2 = plik.replace('.xlsx', '_zmienione.csv')
    df2.to_csv(plik2, index=False)


def wpisz_przelewy(plik, ile_rozr, ile_tab):
    df = p.read_csv(plik, sep=',', dtype={'Konto': str, 'Rezerwacja': str, 'Kwota': str})
    if ile_rozr == 0:
        ile_rozr = len(df)
    if ile_tab == 0:
        ile_tab = 5
    print('Suma kwot to:', sum([float(x.replace(' ', '')) for x in df['Kwota']]))
    print('5 sekund na zmianę okna')
    t.sleep(5)
    for _, row in df.iterrows():
        pdi.write(row['Kwota'])
        t.sleep(0.05)
        pdi.press('Enter')
        t.sleep(0.05)
        pdi.write(row['Konto'])
        t.sleep(0.05)
        pdi.press('Enter')
    for _ in range(ile_rozr):
        pdi.press('up')
        t.sleep(0.05)
    for _ in range(ile_tab):
        pdi.press('tab')
        t.sleep(0.05)
    for _ in range(ile_rozr):
        rozr()
        t.sleep(0.05)
        pdi.press('down')
        t.sleep(0.05)
    

def bnp_plik(plik):
    plik = plik = os.path.join(plik + ".mt940")
    plik2 = plik.replace(".mt940", "_wynik.txt")
    f = open(plik, 'r')
    dane = f.read()
    f.close()
    opisy = re.findall(BNP_OPIS, dane)
    with open(plik2, 'w') as f:
        for przelew in opisy:
            line2 = przelew[1].replace("^20", "")
            line2 = line2.replace("\n", " ")
            line2 = line2.replace("^21", "")
            line2 = line2.replace("^22", "")
            line2 = line2.replace("^23", "")
            line2 = line2.replace("^24", "")
            line2 = line2.replace("^27", "")
            line2 = line2.replace("^32", "")
            matches = re.findall(PATTERN1, line2)
            matches = [(i[0], i[1].replace(" ", ""), i[2]) for i in matches]
            if len(matches) == 1 or (len(matches) == 2 and matches[0][1] == matches[1][1]):
                nr = matches[0][1]
                if int(nr) in dic:
                    f.write(str(slownik(nr, dic)) + "    <= " + nr + " | " + line2 + "\n")
                else:
                    f.write("Trzeba utworzyć! " + "<= " + nr + " | " + line2 + "\n")
            elif len(matches) > 1:
                f.write("Więcej niż 1! " + "| " + line2 + "\n")
            else:
                f.write("Nie zgadza się ze schematem " + "| " + line2 + "\n")


def bnp_wpis(plik, poz1, poz2):
    plik = plik = os.path.join(plik + ".mt940")
    tab = []
    f = open(plik, 'r')
    dane = f.read()
    f.close()
    opisy = re.findall(BNP_OPIS, dane)
    for przelew in opisy:
        desc = przelew[1].replace("^20", "").replace("\n", " ").replace("^21", "").replace("^22", "").replace("^23", "").replace("^24", "").replace("^25", "").replace("^27","").replace("^32", "")
        matches = re.findall(BNP_PATTERN, desc)
        matches = [(i[0], i[1].replace(" ", ""), i[2]) for i in matches]
        if len(matches) == 1 or (len(matches) == 2 and matches[0][1] == matches[1][1]):
            nr = matches[0][1]
            tab.append([int(nr)])
        elif len(matches) > 1:
            wielokrotne = re.findall(BNP_PATTERN2, desc)
            if len(wielokrotne) < 2:
                tab.append(["N/A"])
            else:
                tab.append([])
                for wiel in wielokrotne:
                    rezerwacja = int(wiel[1].replace(" ", ""))
                    kwota = wiel[4].replace(" ", "")
                    tab[-1].append(rezerwacja)
                    tab[-1].append(kwota)
        else:
            tab.append(["N/A"])
    print('Liczba pozycji: ', len(tab))
    start = int(input("Od którego nr rezerwacji? "))
    skipniecie = [int(x) for x in input("skip: ").split()]
    skip2 = [int(x) for x in input("skip2 (prowizje): ").split()]
    print("10 sekund na zmianę okna") 
    t.sleep(10)
    n = 0
    started = None
    for index, row in enumerate(tab):
        if started != None and poz1 > poz2:
            break
        while poz1 in skip2:
            poz1 += 1
            pdi.press('down')
        if poz1 not in skipniecie: 
            if row[0] == start and not started:
                n = 1
                started = index + 1
            if n == 1:
                print(f'{index - started + 2}: {poz1} z {poz2} -> {row}')
                if len(row) == 1:
                    if row[0] != 'N/A':
                        if row[0] in dic:
                            pdi.keyDown('space')
                            pdi.write(slownik(row[0], dic))
                            rozr()
                        pdi.press('down')
                    else:
                        for _ in range(7):
                            pdi.press('backspace')
                        pdi.write('139-5')
                        rozr()
                        pdi.press('down')
                else:
                    for _ in range(7):
                        pdi.press('backspace')
                    pdi.press('left')
                    for _ in range(10):
                        pdi.press('backspace')
                    lista = [row[i:i+2] for i in range(0, len(row), 2)]
                    for i, minirow in enumerate(lista):
                        pdi.write((minirow[1]))
                        pdi.press('tab')
                        if i == 0:
                            for _ in range(7):
                                pdi.press('backspace')
                        pdi.write(slownik(minirow[0], dic, dlugi = True))
                        if i != len(lista)-1:
                            pdi.press('enter')
                    for i in range(len(lista)-1):
                        pdi.press('up')
                    for i in range(7):
                        pdi.press('tab')
                    for i in range(len(lista)):
                        rozr()   
                        pdi.press('down')
                        if i == 0:
                            pdi.press('tab')
                        if i == len(lista)-1:
                            for _ in range(7):
                                pdi.press('tab')
        else:
            pdi.press('down')
        while poz1 in skip2:
            poz1 += 1
            pdi.press('down')
        if n == 1:
            poz1 += 1

def santander_plik(plik):
    plik = plik = os.path.join(plik + ".eMT")
    plik2 = plik.replace(".eMT", "_wynik.txt")
    f = open(plik, 'r')
    dane = f.read()
    f.close()
    opisy = re.findall(SANTANDER_OPIS, dane)
    with open(plik2, 'w') as f:
        for przelew in opisy:
            line2 = przelew.replace("?20", "")
            line2 = line2.replace("\n", "")
            line2 = line2.replace("?31", "")
            line2 = line2.replace("?32", "")
            matches = re.findall(PATTERN1, line2)
            if len(matches) == 1 or (len(matches) == 2 and matches[0][1].replace(" ", "") == matches[1][1].replace(" ", "")):
                nr = matches[0][1].replace(" ", "")
                if int(nr) in dic:
                    f.write(str(slownik(nr, dic)) + "<= " + nr + " | " + line2 + "\n")
                else:
                    f.write("Trzeba utworzyć! " + "<= " + nr + " | " + line2 + "\n")
            elif len(matches) > 1:
                f.write("Więcej niż 1! " + "| " + line2 + "\n")
            else:
                f.write("Nie zgadza się ze schematem " + "| " + line2 + "\n")
    print('Utworzono plik', plik2)


def santander_wpis(plik, poz1, poz2):
    plik = plik = os.path.join(plik + ".eMT")
    tab = []
    f = open(plik, 'r')
    dane = f.read()
    f.close()
    opisy = re.findall(SANTANDER_OPIS, dane)
    for przelew in opisy:
        desc = przelew.replace("?30", "").replace("\n", "").replace("?31", "")
        matches = re.findall(SANTANDER_PATTERN, desc)
        if len(matches) == 1 or (len(matches) == 2 and matches[0][0].replace(" ", "") == matches[1][0].replace(" ", "")):
            nr = matches[0][0].replace(" ", "")
            tab.append([int(nr)])
        elif len(matches) > 1 or (len(matches) == 2 and matches[0][0] != matches[1][0]):
            wielokrotne = re.findall(SANTANDER_PATTERN2, desc)
            if len(wielokrotne) < 2:
                tab.append(["N/A"])
            else:
                tab.append([])
                for wiel in wielokrotne:
                    rezerwacja = int(wiel[0].replace(" ", ""))
                    kwota = wiel[3].replace(" ", "")
                    tab[-1].append(rezerwacja)
                    tab[-1].append(kwota)
        else:
            tab.append(["N/A"]) 
    print('Liczba pozycji: ', len(tab))
    start = int(input("Od którego nr rezerwacji? "))
    skipniecie = [int(x) for x in input("skip: ").split()]
    skip2 = [int(x) for x in input("skip2 (wplacone w banku): ").split()]
    print("10 sekund na zmianę okna") 
    t.sleep(10)
    ile = poz2 - poz1 + 1
    print(f'{ile=}')
    n = 0
    started = None
    for index, row in enumerate(tab):
        if started and index >= started -1 + ile:
            break
        if poz1 in skip2:
            poz1 += 1
            pdi.press('down')
        if poz1 not in skipniecie: 
            if row[0] == start:
                n = 1
                started = index + 1
            if n == 1:
                if len(row) == 1:
                    if row[0] != 'N/A':
                        if row[0] in dic:
                            pdi.keyDown('space')
                            pdi.write(slownik(row[0], dic))
                            rozr()
                        pdi.press('down')
                    else:
                        for _ in range(7):
                            pdi.press('backspace')
                        pdi.write('139-5')
                        rozr()
                        pdi.press('down')
                else:
                    for _ in range(7):
                        pdi.press('backspace')
                    pdi.press('left')
                    for _ in range(10):
                        pdi.press('backspace')
                    lista = [row[i:i+2] for i in range(0, len(row), 2)]
                    for i, minirow in enumerate(lista):
                        pdi.write((minirow[1]))
                        pdi.press('tab')
                        if i == 0:
                            for _ in range(7):
                                pdi.press('backspace')
                        pdi.write(slownik(minirow[0], dic, dlugi = True))
                        if i != len(lista)-1:
                            pdi.press('enter')
                    for i in range(len(lista)-1):
                        pdi.press('up')
                    for i in range(7):
                        pdi.press('tab')
                    for i in range(len(lista)):
                        rozr()   
                        pdi.press('down')
                        if i == 0:
                            pdi.press('tab')
                        if i == len(lista)-1:
                            for _ in range(7):
                                pdi.press('tab')
        else:
            pdi.press('down')
        if n == 1:
            poz1 += 1

def lista():
    lista1 = input('Nr / nr-y i kwoty: ')
    if lista1 == 'q':
        return
    lista1 = lista1.replace('-', ' ').replace(';', ' ').replace('(', ' ').replace(')', ' ').replace('/', ' ')
    print(lista1)
    if len(lista1.split()) == 1:
        rez = lista1.split()[0]
        if int(rez) in dic:
            r = Tk()
            r.withdraw()
            r.clipboard_clear()
            r.clipboard_append(slownik(rez, dic,dlugi = True))
            r.update()
            print(slownik(rez, dic, dlugi = True))
            print('skopiowano do schowka')
    else:
        lista2 = [lista1.split()[i:i+2] for i in range(0, len(lista1.split()), 2)]
        print(lista2)
        print("5 sekund na zmianę okna") 
        t.sleep(5) 
        for i, row in enumerate(lista2): 
            pdi.write(str(row[1]))
            pdi.press('tab')
            pdi.write(slownik(row[0], dic, dlugi = True))
            if i != len(lista2)-1:
                pdi.press('enter')
        for i in range(len(lista2)-1):
            pdi.press('up')
        for i in range(7):
            pdi.press('tab')
        for i in range(len(lista2)):
            rozr()
            if i != len(lista2)-1:
                pdi.press('down')
            if i == 0:
                pdi.press('tab')
    lista()


def main():
    co = input('''Co chcesz zrobic?:
               
    (a)nulacje 
    (p)rzeksiegowanie
    (z)amienic plik z Przelewy24
    (w)pisac plik z Przelewy24 do PK
    (f)vs
    (bp) bnp plik
    (b) bnp wpis
    (sp) santander plik
    (s) santander wpis
    (n)umery - z listy
    (pp) przeksiewanie prawe
    (pl) przeksiegowanie lewe
    (fvzp) 
    (q)uit
               
''')
    if co == 'z':
        plik = os.path.join('Przelewy24', input('Podaj nazwę pliku: \n'))
        zamien_przelewy(plik)
        if input('Czy chcesz od razu wpisać dane do PK? (t/n) \n') == 't':
            wpisz_przelewy(plik.replace('.csv', '_zmienione.csv'), 0, 5)
    elif co == 'fa':
        plik = os.path.join('Przelewy24', input('Podaj nazwę pliku: \n'))
        zamien_faktury(plik)
    elif co == 'p':
        plik = os.path.join('pk', input('Podaj nazwę pliku: \n'))
        przeksiegowanie(plik)
    elif co == 'w':
        plik = os.path.join('Przelewy24', input('Podaj nazwę pliku: \n'))
        ile_rozr = int(input('Ile rozrachunków? '))
        ile_tab = int(input('Ile tab? '))
        wpisz_przelewy(plik, ile_rozr, ile_tab)
    elif co == 'f':
        plik = input("Podaj nazwę pliku: ")
        plik = os.path.join('FVS', plik)
        plik2 = plik + '.xlsx'
        plik3 = plik + '_rez.xlsx'
        df = p.read_excel(plik2, header=None)
        df.columns = ['fvs', 'rez']
        df2 = df.copy()
        df2['rez'] = df.apply(lambda x: find_reservation(x['rez']), axis=1)
        df2.reset_index(drop=True, inplace=True)
        df2.to_excel(plik3, index=False)
    elif co == 'fd':
        plik = input("Podaj nazwę pliku: ")
        plik = os.path.join('Przelewy24', plik)
        df = p.read_excel(plik, header=None)
        df.columns = ['rez', 'kwota']
        df2 = p.DataFrame()
        df2['Kwota'] = df['kwota']
        df2['Rezerwacja'] = df.apply(lambda x: find_reservation(str(x['rez']), spaces = True), axis=1)
        df2['Konto'] = df2.apply(lambda x: slownik(x['Rezerwacja'], dic, dlugi = True), axis=1)
        df2.reset_index(drop=True, inplace=True)
        plik2 = plik.replace('.xlsx', '_rez.csv')
        df2.to_csv(plik2, index=False)
    elif co == 'bp':
        plik = os.path.join('wb_BNP', input('Podaj nazwę pliku z wyciągiem (bez .mt940): \n'))
        bnp_plik(plik)
        print('Utworzono plik', plik.replace(".mt940", "_wynik.txt"))    
    elif co == 'b':
        plik = input("Podaj nazwę pliku z wyciągiem (bez .mt940): ")
        plik = os.path.join('wb_BNP', plik)
        poz1 = int(input("Pozycja początkowa: "))
        poz2 = int(input("Pozycja końcowa: "))
        bnp_wpis(plik, poz1, poz2)
    elif co == 'sp':
        plik = input("Podaj nazwę pliku z wyciągiem (bez .eMT): ")
        plik = os.path.join('wb_santander', plik)
        santander_plik(plik)
        main()
    elif co == 's':
        plik = input("Podaj nazwę pliku z wyciągiem (bez .eMT): ")
        plik = os.path.join('wb_santander', plik)
        poz1 = int(input("Pozycja początkowa: "))
        poz2 = int(input("Pozycja końcowa: "))
        print(poz1, poz2)
        santander_wpis(plik, poz1, poz2)
    elif co == 'n':
        lista()
    elif co == 'a':
        plik = os.path.join('anulacje', input('Podaj nazwę pliku: \n'))
        anulacje(plik)
    elif co == 'pp':
        pk_prawy(os.path.join('pk', 'grosze.xlsx'))
    elif co == 'pl':
        pk_lewy(os.path.join('pk', 'grosze.xlsx'))
    elif co == 'rozr':
        ile_rozr = int(input('Ile rozrachunków? '))
        t.sleep(5)
        for _ in range(ile_rozr):
            rozr()
            t.sleep(0.05)
            pdi.press('down')
            t.sleep(0.05)
    elif co == 'fvzp':
        fvzp(os.path.join('fvzp', 'fvzp.xlsx'))
    else:
        return
    main()


if __name__ == '__main__':
    main()