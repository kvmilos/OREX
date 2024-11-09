from pdfminer.high_level import extract_text
from os import listdir, rename, makedirs
from re import search

NUM_FAK = r'FS/\d{2}/\d{2}/\d{5}|FSP\d{4}/\d{2}/\d{3}|FVSW\d{4}/\d{2}/\d{3}|FO\d{4}/\d{2}/\d{4}|PRO\d{4}-\d{2}/\d{5}'
NUM_REZ = r'1\d{6}'

def main():
    makedirs('./zmienione', exist_ok=True)
    for file in listdir('.'):
        if file.endswith('.pdf'):
            with open(file, 'rb') as f:
                text = extract_text(f)
                print(text)
            num_fak = search(NUM_FAK, text).group().replace('/', '_')
            num_rez = search(NUM_REZ, text)
            if num_rez:
                num_rez = num_rez.group()
                rename(file, f'./zmienione/{num_fak} rez. {num_rez}.pdf')
            else:
                rename(file, f'./zmienione/{num_fak}.pdf')

if __name__ == '__main__':
    main()