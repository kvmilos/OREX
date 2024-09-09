import os
import pandas as p
# with open ('a.txt', 'r') as f:
#     a = ''
#     for line in f:
#         a += (line.replace('/2023\n', '|'))
#     print(a)
df = p.read_excel(os.path.join('Przelewy24', 'fvzp968.xlsx'), header=None, na_values='NA')
# now save to the dataframe only the odd rows
df = df.iloc[::2]
df[1] = df[1].str.replace(' PLN', '')
print(df)
df.to_excel(os.path.join('Przelewy24', 'fvzp968.xlsx'), index=False, header=False)