import re
import pandas as p
import os

dokument = r'Dokument[\s\S]*?\n\}'
pattern = r'\t*(Zapis)[^\n]*\n[^\n]*\n\t*konto =(201-2-1-)[^\n]*\n[^\n]*\n[^\n]*\n[^\n]*\n[^\n]*\n[^\n]*\n\t*opis[^\n]*(1[012][\d]{5})'

df = p.read_csv('https://raw.githubusercontent.com/kvmilos/OREX/main/kontrahenci.csv')
dic = dict(zip(df['kod'], df['pozycja']))

def zamiana(plik):
    with open(plik, "r") as f:
        dane = f.read()

    faktury = re.findall(dokument, dane)
    for faktura in faktury:
        matches = re.findall(pattern, faktura)
        print(matches)
        for match in matches:
            if int(match[2]) in dic: 
                print(f"Zamieniam {match[1]} na 201-2-1-{str(dic[int(match[2])])}")
                faktura2 = faktura.replace(match[1], f'201-2-1-{str(dic[int(match[2])])}')
                dane = dane.replace(faktura, faktura2)

    plik2 = plik.replace(".txt", "_zmienione.txt")
    with open(plik2, "w") as f:
        f.write(dane)
        print('Utworzono plik', plik2)

def main():
    for plik in os.listdir():
        print(plik)
        if plik.endswith(".txt"):
            zamiana(plik)
    n = input("Naciśnij enter aby zamknąć program")

if __name__ == '__main__':
    main()