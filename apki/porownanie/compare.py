import pandas as pd
from decimal import Decimal, ROUND_HALF_UP

def read_xlsx(file, how):
    df = pd.read_excel(file, header=0)
    headers = df.columns.tolist()

    df1 = df.iloc[:, 0].tolist()
    df2 = df.iloc[:, 1].tolist()

    df1 = [str(i).replace(',', '.').replace('\xa0', '') for i in df1]
    df2 = [str(i).replace(',', '.').replace('\xa0', '') for i in df2]

    if how == '2':
        df1 = list(set(df1))
        df2 = list(set(df2))

    elif how == '1':
        df1 = [Decimal(i).quantize(Decimal('0.01'), rounding=ROUND_HALF_UP) for i in df1]
        df2 = [Decimal(i).quantize(Decimal('0.01'), rounding=ROUND_HALF_UP) for i in df2]
    
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

def main():
    print('To start the comparison, prepare an Excel file with two columns of data to compare.')
    print('The file should be in the same folder as this app and its name should be "compare.xlsx".\n')
    print('Choose which option you want the app to do: \n(1): compare amounts exactly as they are\n(2): compare the numbers removing any duplicates first\n')
    print('If you are ready, enter 1 or 2 and press Enter: ')
    how = input().strip()

    file = 'compare.xlsx'
    headers, list1, list2 = read_xlsx(file, how)

    exact_matches, unmatched_list1, unmatched_list2 = match_lists(list1, list2)
    save_to_excel(headers, exact_matches, unmatched_list1, unmatched_list2)

if __name__ == "__main__":
    main()