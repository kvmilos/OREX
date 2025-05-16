import pandas as pd
import datetime

# Wczytanie pliku wejściowego
input_file = 'dane.xlsx'
df = pd.read_excel(input_file)

# Usuwamy spacje oraz inne białe znaki kolumn from, amount, to jeśli są stringami, jeśli są floatami/intami to zostawiamy
df['from']   = df['from'].apply(lambda x: x.strip() if isinstance(x, str) else x)
df['amount'] = df['amount'].apply(lambda x: x.strip() if isinstance(x, str) else x)
df['to']     = df['to'].apply(lambda x: x.strip() if isinstance(x, str) else x)

# Pozbywamy się \xa0
df['from']   = df['from'].apply(lambda x: x.replace('\xa0', '') if isinstance(x, str) else x)
df['amount'] = df['amount'].apply(lambda x: x.replace('\xa0', '') if isinstance(x, str) else x)
df['to']     = df['to'].apply(lambda x: x.replace('\xa0', '') if isinstance(x, str) else x)

# Zamiana kolumn from, to na string
df['from'] = df['from'].astype(str)
df['to']   = df['to'].astype(str)

# Jeśli kolumna 'amount' jest w formacie tekstowym z przecinkiem jako separatorem dziesiętnym, zamieniamy ją na float
df['amount'] = df['amount'].astype(str).str.replace(',', '.').astype(float)

# Pobranie bieżącej daty w wymaganych formatach
today        = datetime.date.today()
date_str     = today.strftime('%Y-%m-%d')   # np. 2025-02-20
doc_date_str = today.strftime('%d/%m/%y')   # np. 20/02/25

# Lista na wiersze wynikowe
output_rows = []

# Liczniki
row_counter = 0    # do wyznaczenia "Nr wiersza" (co 10000)
doc_counter = 1    # numeracja dokumentu (zmiana co 2 wiersze)

for idx, row in df.iterrows():
    from_val = row['from']
    amount   = row['amount']
    to_val   = row['to']

    opis        = f"Przeks. z rez. {from_val} na rez. {to_val}"
    nr_dokumentu = f"PK_REZ {doc_date_str} {doc_counter:03d}"

    # Pierwszy wiersz – 'from'
    row_counter += 1
    output_rows.append({
        "Nr wiersza": row_counter * 10000,
        "Nazwa instancji dziennika": "PRZEK-REZ",
        "Nazwa szablonu dziennika": "OGÓLNE",
        "Data księgowania": date_str,
        "Data VAT": date_str,
        "Data dokumentu": date_str,
        "Nr dokumentu": nr_dokumentu,
        "Typ konta": "Nabywca",
        "Nr konta": "",
        "Reservation No.": from_val,
        "Opis": opis,
        "Kwota": amount,
        "Kwota Debet": amount,
        "Kwota Kredyt": 0,
        "Typ konta przeciwst.": "Konto K/G",
        "Nr konta przeciwst.": ""
    })

    # Drugi wiersz – 'to'
    row_counter += 1
    output_rows.append({
        "Nr wiersza": row_counter * 10000,
        "Nazwa instancji dziennika": "PRZEK-REZ",
        "Nazwa szablonu dziennika": "OGÓLNE",
        "Data księgowania": date_str,
        "Data VAT": date_str,
        "Data dokumentu": date_str,
        "Nr dokumentu": nr_dokumentu,
        "Typ konta": "Nabywca",
        "Nr konta": "",
        "Reservation No.": to_val,
        "Opis": opis,
        "Kwota": -amount,
        "Kwota Debet": 0,
        "Kwota Kredyt": amount,
        "Typ konta przeciwst.": "Konto K/G",
        "Nr konta przeciwst.": ""
    })

    doc_counter += 1

# Utworzenie i zapis wyników
output_df   = pd.DataFrame(output_rows)
output_file = 'result.xlsx'
output_df.to_excel(output_file, index=False)

print(f"Plik {output_file} został wygenerowany.")