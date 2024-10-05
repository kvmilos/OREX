import pandas as pd

url = 'https://raw.githubusercontent.com/kvmilos/OREX/main/kontrahenci.csv'
df = pd.read_csv(url, dtype=str)
dic = dict(zip('201-2-1-' + df['pozycja'], df['kod']))

def get_reservation(x: str):
    if x in dic:
        return dic[x]
    else:
        return x
    
df = pd.read_excel('egipt.xlsx')
df['rezerwacja'] = df['konto'].apply(get_reservation)
# set type of rezerwacja to str
df['rezerwacja'] = df['rezerwacja'].astype(str)
df['c9'] = df['c9'].astype(float)
print(df.head())

df2 = pd.read_excel('egipt-samo.xlsx')
df2['Reservation'] = df2['Reservation'].astype(str)
df2['Catalogue price'] = df2['Catalogue price'].astype(float)
print(df2.head())

symfonia = df[['rezerwacja', 'c9']]
print(symfonia.head())

samo = df2[['Reservation', 'Catalogue price']]
print(samo.head())

# rename columns 
symfonia.columns = ['rezerwacja', 'cena']
samo.columns = ['rezerwacja', 'cena']

result = pd.DataFrame(columns=['rezerwacja', 'suma_symfonia', 'suma_samo', 'roznica'])
result['rezerwacja'] = pd.concat([symfonia['rezerwacja'], samo['rezerwacja']]).unique()
result['suma_symfonia'] = result['rezerwacja'].apply(lambda x: symfonia[symfonia['rezerwacja'] == x]['cena'].astype(float).sum())
result['suma_samo'] = result['rezerwacja'].apply(lambda x: samo[samo['rezerwacja'] == x]['cena'].astype(float).sum())
result['roznica'] = result['suma_symfonia'] - result['suma_samo']
# sort by rezerwacja
result = result.sort_values(by='rezerwacja')
result.to_excel('result2.xlsx', index=False)