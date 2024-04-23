from pypdf import PdfReader
import os
import re

NUM_FAK = r'FS/\d{2}/\d{2}/\d{5}'
NUM_REZ = r'1\d{6}'

def main():
    for file in os.listdir('.'):
        if file.endswith('.pdf'):
            with open(file, 'rb') as f:
                text = PdfReader(f).pages[0].extract_text()
            num_fak = re.search(NUM_FAK, text).group().replace('/', '_')
            num_rez = re.search(NUM_REZ, text).group()
            os.rename(file, f'./zmienione/{num_fak} rez. {num_rez}.pdf')

if __name__ == '__main__':
    main()