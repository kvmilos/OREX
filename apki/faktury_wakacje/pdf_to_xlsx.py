from os import path, makedirs, rename, listdir
from re import search, finditer
from pdfminer.high_level import extract_text
from pandas import DataFrame, ExcelWriter
from requests import get
from json import loads

def fetch_patterns():
    url = "https://raw.githubusercontent.com/kvmilos/OREX/refs/heads/main/regex_patterns.json"
    response = get(url)
    if response.status_code == 200:
        return loads(response.text)
    else:
        print("An error occurred while fetching patterns.")
        return None

patterns = fetch_patterns()
if patterns is None:
    print("An error occurred while fetching patterns.")
    exit(1)

makedirs('done', exist_ok=True)
faktury = [['numer', 'data_wystawienia', 'termin_platnosci', 'netto', 'kw_vat', 'brutto', 'st_vat', 'rezerwacja', 'plik']]
bledy = [['plik', 'opis']]
DEBUG = False  # Enable debug mode for detailed output

regex_keys = ["numer_faktury", "data_faktury", "termin", "netvat", "brutto", "stawka", "rezerwacja"]
try:
    fak_NUMER = patterns["wakacje_new"]["numer_faktury"]
    fak_DATA_WYSTAWIENIA = patterns["wakacje_new"]["data_faktury"]
    fak_TERMIN_PLATNOSCI = patterns["wakacje_new"]["termin"]
    fak_NETVAT = patterns["wakacje_new"]["netvat"]
    fak_BRUTTO = patterns["wakacje_new"]["brutto"]
    fak_STAWKA = patterns["wakacje_new"]["stawka"]
    fak_REZERWACJA = patterns["wakacje_new"]["rezerwacja"]
except KeyError as e:
    print(f"Missing key in patterns: {e}")
    exit(1)

def convert_to_float(value):
    """Converts a string number with commas or spaces into a float."""
    value = value.replace(' ', '').replace('\xa0', '').replace(',', '.')
    return float(value)

def blad(text, file_name):
    # Check each regex pattern and provide detailed output on failure
    missing_patterns = []
    if not search(fak_NUMER, text):
        missing_patterns.append("fak_NUMER")
    if not search(fak_DATA_WYSTAWIENIA, text):
        missing_patterns.append("fak_DATA_WYSTAWIENIA")
    if not search(fak_TERMIN_PLATNOSCI, text):
        missing_patterns.append("fak_TERMIN_PLATNOSCI")
    if not search(fak_REZERWACJA, text):
        missing_patterns.append("fak_REZERWACJA")
    if not search(fak_NETVAT, text):
        missing_patterns.append("fak_NETVAT")
    if not search(fak_BRUTTO, text):
        missing_patterns.append("fak_BRUTTO")
    if not search(fak_STAWKA, text):
        missing_patterns.append("fak_STAWKA")

    if missing_patterns:
        if DEBUG:
            print(f"[DEBUG] Missing patterns in file '{file_name}': {', '.join(missing_patterns)}")
        return missing_patterns
    return []

def parse_page_data(text, file_name):
    numer = search(fak_NUMER, text).group(1)
    data_wystawienia = search(fak_DATA_WYSTAWIENIA, text).group(1)
    termin_platnosci = search(fak_TERMIN_PLATNOSCI, text).group(1)
    
    netto = convert_to_float(search(fak_NETVAT, text).group(1))
    vat = convert_to_float(search(fak_NETVAT, text).group(2))
    brutto = convert_to_float(list(finditer(fak_BRUTTO, text))[-1].group(1))
    
    stawka = list(finditer(fak_STAWKA, text))[-1].group(1)
    rezerwacja = int(search(fak_REZERWACJA, text).group(1))
    
    return [numer, data_wystawienia, termin_platnosci, vat, netto, brutto, stawka + '%', rezerwacja, file_name]

for file in listdir('.'):
    if not file.endswith('.pdf'):
        continue
    
    try:
        with open(file, 'rb') as f:
            text = extract_text(f)
        
        missing_patterns = blad(text, file)
        if missing_patterns:
            bledy.append([file, ', '.join(missing_patterns)])
            print(f"Błąd w pliku {file}")
        else:
            faktury.append(parse_page_data(text, file))
        
        rename(file, path.join('done', file))
    except Exception as e:
        print(f"Błąd w pliku {file}: {e}")
        if DEBUG:
            import traceback
            traceback.print_exc()

if len(faktury) > 0:
    fak = DataFrame(faktury)
    bl = DataFrame(bledy)
    with ExcelWriter('wakacje_pdfy.xlsx', mode='w') as writer:
        fak.to_excel(writer, sheet_name='faktury', index=False, header=False)
        bl.to_excel(writer, sheet_name='bledy', index=False, header=False)

input('Gotowe! Naciśnij Enter aby zamknąć.')
