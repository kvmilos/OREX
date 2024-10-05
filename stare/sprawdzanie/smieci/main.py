import pandas as pd

url = 'https://raw.githubusercontent.com/kvmilos/OREX/main/kontrahenci.csv'
df = pd.read_csv(url, dtype=str)
dic = dict(zip('201-2-1-' + df['pozycja'], df['kod']))

def get_reservation(x: str):
    if x in dic:
        return dic[x]
    else:
        return x


samo = pd.read_excel('samo.xlsx')
samo['reservation'] = samo['reservation'].astype(str)
# print(samo.head())



df703 = pd.read_excel('symfonia.xlsx', sheet_name='703')
df704 = pd.read_excel('symfonia.xlsx', sheet_name='704')
df705 = pd.read_excel('symfonia.xlsx', sheet_name='705')
df76398 = pd.read_excel('symfonia.xlsx', sheet_name='763_98')
df7637 = pd.read_excel('symfonia.xlsx', sheet_name='763_7')

df703['sheet'] = '703'
df704['sheet'] = '704'
df705['sheet'] = '705'
df76398['sheet'] = '763_98'
df7637['sheet'] = '763_7'

symfonia = pd.concat([df703, df704, df705, df76398, df7637])
symfonia = symfonia[['date', 'invoice', 'description', 'amount', 'account', 'sheet']]
symfonia['reservation'] = symfonia['account'].apply(get_reservation)
# print(symfonia.head())



result = pd.DataFrame(columns=['reservation', 'sum_symfonia', 'sum_samo', 'difference'])
result['reservation'] = pd.concat([symfonia['reservation'], samo['reservation']]).unique()
result['sum_symfonia'] = result['reservation'].apply(lambda x: symfonia[symfonia['reservation'] == x]['amount'].astype(float).sum())
result['sum_samo'] = result['reservation'].apply(lambda x: samo[samo['reservation'] == x]['amount'].astype(float).sum())
result['difference'] = result['sum_symfonia'] - result['sum_samo']
result.to_excel('result.xlsx', index=False)

