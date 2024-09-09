import pandas as pd

dostawcy = pd.read_excel('lul.xlsx', sheet_name='Dostawcy', header=0, dtype={'nip': str, 'nr': str})
symf_odpowi = pd.read_excel('lul.xlsx', sheet_name='symf_odpowi', header=0, dtype={'kod': str, 'nip': str})
symf_pozycje = pd.read_excel('lul.xlsx', sheet_name='symf_pozycje', header=None, dtype=str)

# add a column name
symf_pozycje.columns = ['kod']

# make a dict out of symf_odpowi['kod'] and symf_odpowi['nip']
symf_odpowi_dict = symf_odpowi.set_index('kod').to_dict()['nip']
print(symf_odpowi_dict)

# make a dict out of dostawcy
dostawcy_dict = dostawcy.set_index('nip').to_dict()['nr']
print(dostawcy_dict)

# based on the dict, add a column to symf_pozycje there where its possinle, if not add 'brak'
symf_pozycje['nip'] = symf_pozycje['kod'].map(symf_odpowi_dict)
symf_pozycje['nr'] = symf_pozycje['nip'].map(dostawcy_dict).fillna('brak')

# save the result
symf_pozycje.to_excel('result.xlsx', index=False)