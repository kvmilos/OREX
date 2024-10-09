import pandas as pd
from decimal import Decimal

def read_xlsx(file):
    df = pd.read_excel(file, header=0)
    df = df.map(lambda x: str(x).replace(',', '.').replace('\xa0', ''))
    headers = df.columns.tolist()
    df1 = df.iloc[:, 0].tolist()
    df1 = [Decimal(i) for i in df1]
    df2 = df.iloc[:, 1].tolist()
    df2 = [Decimal(i) for i in df2]
    
    return headers, df1, df2

def match_lists(list1, list2):
    matches = []
    unmatched_list1 = list1.copy()
    unmatched_list2 = list2.copy()
    
    for i in list1:
        if i in unmatched_list2:
            matches.append((i, i))
            unmatched_list1.remove(i)
            unmatched_list2.remove(i)
    
    
    return matches, unmatched_list1, unmatched_list2

def save_to_excel(headers, exact_matches, unmatched_list1, unmatched_list2):
    df_exact = pd.DataFrame(exact_matches, columns=headers)
    unmatched_data = {headers[0]: unmatched_list1 + [None] * (len(unmatched_list2) - len(unmatched_list1)),
                      headers[1]: unmatched_list2 + [None] * (len(unmatched_list1) - len(unmatched_list2))}
    df_unmatched = pd.DataFrame(unmatched_data)

    with pd.ExcelWriter('comparison_output.xlsx') as writer:
        df_exact.to_excel(writer, sheet_name='Exact Matches', index=False)
        df_unmatched.to_excel(writer, sheet_name='Unmatched', index=False)
        
    print('Comparison results saved to comparison_output.xlsx')

def main():
    print('EN: To start the comparison, prepare an Excel file with two columns of data to compare.\nPL: Aby zacząć porównywanie, przygotuj plik Excel z dwiema kolumnami danych do porównania.\n')
    print('EN: The file should be in the same folder as this app and its name should be "compare.xlsx".\nPL: Plik powinien znajdować się w tym samym folderze co ta aplikacja i jego nazwa powinna być "compare.xlsx".\n')
    _ = input('EN: If you are ready, press Enter.\nPL: Jeśli jesteś gotowy/a, naciśnij Enter.')

    file = 'compare.xlsx'
    headers, list1, list2 = read_xlsx(file)

    exact_matches, unmatched_list1, unmatched_list2 = match_lists(list1, list2)

    save_to_excel(headers, exact_matches, unmatched_list1, unmatched_list2)

if __name__ == '__main__':
    main()