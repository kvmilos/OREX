"""
This script processes an Excel file generated from SAMO. Its task is to classify each row into 6 different groups:
1. Unpaid advances - when the advance date is in the past, the total date is in the future, and the paid amount is less than the advance amount
2. 30/30 - when the check-in date is more than 30 days in the future and the payment is greater than 30% of the catalogue price
3. Overdue receivables - when the date total is in the past (including today), the check-in date is in the future, the debt is greater than 0, and the days from confirmation are greater than or equal to 3 
4. Payables in term - when the check-in date is less than or equal to 30 days in the future, the date total is in the future, and the debt is less than 0
5. Receivables in term - not classified in any of the other groups, debt is 0 or they are not overdue
6. Overdue receivables after check-in - when the check-in date is in the past (including today) and the debt is greater than 0
7. Overdue payables after check-in - when the check-in date is in the past and the debt is less than 0
Then, based on the classification, a 'Result' column is added with the appropriate calculations.
Additionally, a 'Grouped agents' column is added with agents assigned based on dictionaries.
"""
import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import os
from decimal import *
import json

OUTPUT_DIRECTORY = './'
INPUT_DIRECTORY = './'
MISSING_RESERVATIONS_DIR = './'
SHEET_NAME = 'all rezervation'

getcontext().prec = 28

def adjust_excel_headers(file_path, sheet_name, new_headers):
    """
    Input excel file has strange layout of headers, this function adj
    """
    df = pd.read_excel(file_path, sheet_name=sheet_name, skiprows=3)
    df.columns = new_headers
    return df

def rename_columns(df, new_column_names):
    return df.rename(columns=new_column_names)

def find_duplicates(df):
    duplicates = df[df.duplicated()].sort_values(by=['Nr rezerwacji'])
    return duplicates

def find_missing_reservations(df, df_prev):
    missing_reservations = df_prev.merge(df, on='Nr rezerwacji', how='left', indicator=True)
    missing_reservations = missing_reservations[missing_reservations['_merge'] == 'left_only']
    missing_reservations = missing_reservations.drop(columns=['_merge'])
    return missing_reservations

def analyse_data(df, df_prev):
    duplicates = find_duplicates(df)

    if not duplicates.empty:
        print('Duplicated rows found')
        duplicates.to_excel("duplicates.xlsx", index=False)
        df = df.drop_duplicates()   
    
    missing_reservations = find_missing_reservations(df, df_prev)
    print('Number of missing reservations:', len(missing_reservations["Nr rezerwacji"]))
    path = MISSING_RESERVATIONS_DIR + date + '_missing_reservations.xlsx'
    missing_reservations.to_excel(path, index=False)

    return df

# def analyse_data(df, df_prev):
    
#     if len(df[df.duplicated()]) > 0:
#         print('Duplicated rows found')
#         duplicates = df[df.duplicated()].sort_values(by=['Nr rezerwacji'])
#         duplicates.to_excel("duplicates.xlsx", index=False)
#         df = df.drop_duplicates()

#     missing_reservations = df_prev.merge(df, on='Nr rezerwacji', how='left', indicator=True)
#     missing_reservations = missing_reservations[missing_reservations['_merge'] == 'left_only']
#     missing_reservations = missing_reservations.drop(columns=['_merge'])
#     path = MISSING_RESERVATIONS_DIR + date + '_missing_reservations.xlsx'
#     missing_reservations.to_excel(path, index=False)

#     return df
    

def classify_rows(df):
    """
    Classifies rows of a DataFrame into specific categories based on business rules.

    Args:
        df (pd.DataFrame): The input DataFrame containing reservation data.

    Returns:
        pd.DataFrame: The DataFrame with an added 'Classification' column.
    """
        
    # Calculate columns for classifications
    today = pd.Timestamp.today().normalize()
    df['Advance date'] = pd.to_datetime(df['Advance date']).dt.normalize()
    df['Advance days'] = (df['Advance date'] - today).dt.days
    df['Check-in date'] = pd.to_datetime(df['Check-in date']).dt.normalize()
    df['Date total'] = df['Check-in date'] - pd.DateOffset(days=21)
    df['Total days'] = (df['Date total'] - today).dt.days
    df['Check-in days'] = (df['Check-in date'] - today).dt.days
    df['Advance amount'] = (df['Catalogue price'] * df['% Advance'] / Decimal('100.00')).map(round_decimal_down) 
    df['Advance amount missing'] = (df['Advance amount'] - df['Paid']).map(round_decimal)
    df['Overpayment of 30/30'] = -((df['Paid'] - Decimal('0.30') * df['Catalogue price']).map(round_decimal))
    df['Overdue receivables amount'] = (df['Catalogue price'] - df['Paid']).map(round_decimal)
    df['Payables in term amount'] = df['Debt']
    df['Overdue receivables after check-in amount'] = (df['Catalogue price'] - df['Paid']).map(round_decimal)
    df['Days from confirmation'] = (today - df['Confirmation date']).dt.days
    df['70% Catalogue price'] = (df['Catalogue price'] * Decimal('0.70')).map(round_decimal)
    df['Confrmation date'] = pd.to_datetime(df['Confirmation date'])
    df['Check-in date'] = pd.to_datetime(df['Check-in date'])
    df['Days to check-in after confirmation'] = (df['Check-in date'] - df['Confirmation date']).dt.days


    # Define conditions and choices for classification
    conditions = [
        # Unpaid advances
        (df['Advance date'] < today) & (today < df['Date total']) & (df['Paid'] < df['Advance amount']),
        # 30/30
        (df['Check-in days'] > 30) & (df['Overpayment of 30/30'] < 0),
        # Overdue receivables
        (df['Date total'] <= today) & (today < df['Check-in date']) & (df['Debt'] > 0) & (df['Days from confirmation'] >= 3),
        # Payables in term
        (df['Check-in days'] <= 30) & (today <= df['Check-in date']) & (df['Debt'] < 0),
        # Overdue receivables after check-in
        (df['Check-in date'] <= today) & (df['Debt'] > 0),
        # Overdue payables after check-in
        (today > df['Check-in date']) & (df['Debt'] < 0),
        # Overdue receivables - 3 days after confirmation
        (df['Date total'] <= today) & (today < df['Check-in date']) & (df['Debt'] > 0) & (df['Days from confirmation'] < 3)
    ]


    choices = ['Unpaid advances', '30/30', 'Overdue receivables', 'Payables in term', 'Overdue receivables after check-in', 'Overdue payables after check-in', 'Overdue receivables - 3 days after confirmation']
    df['Classification'] = np.select(conditions, choices, default='Receivables in term')
    
    return df

def calculate_result(df):
    """
    Adds a 'Result' column to the DataFrame based on the classification.

    Args: df (pd.DataFrame): The input DataFrame containing reservation data with a 'Classification' column.

    Returns: pd.DataFrame: The DataFrame with an added 'Result' column.
    """
    result_conditions = [
        (df['Classification'] == 'Unpaid advances'),
        (df['Classification'] == '30/30'),
        (df['Classification'] == 'Overdue receivables'),
        (df['Classification'] == 'Payables in term'),
        (df['Classification'] == 'Receivables in term'),
        (df['Classification'] == 'Overdue receivables after check-in'),
        (df['Classification'] == 'Overdue payables after check-in'),
        (df['Classification'] == 'Overdue receivables - 3 days after confirmation')
    ]
    result_choices = [
        # Unpaid advances
        df['Advance amount missing'],
        # 30/30
        df['Overpayment of 30/30'],
        # Overdue receivables
        df['Overdue receivables amount'],
        # Payables in term
        df['Payables in term amount'],
        # Receivables in term
        df['Debt'],
        # Overdue receivables after check-in
        df['Overdue receivables after check-in amount'],
        # Overdue payables after check-in
        df['Payables in term amount'],
        # Overdue receivables - 3 days after confirmation
        df['Overdue receivables amount']
    ]
    df['Result'] = np.select(result_conditions, result_choices, default=None)
    
    return df

def add_receivables_to_30(df):
    rows_to_duplicate = df[df['Classification'] == '30/30'].copy()
    rows_to_duplicate['Classification'] = 'Receivables in term'
    rows_to_duplicate['Result'] = rows_to_duplicate['70% Catalogue price']
    df = pd.concat([df, rows_to_duplicate], ignore_index=True)

    return df

def add_type_of_arrears(df):
    """
    Adds column for collections department
    """
    df['Type of arrears'] = np.where(df['Classification'] == 'Unpaid advances', 'Zaliczka', np.where(df['Classification'].isin(['Overdue receivables', 'Overdue receivables after check-in', 'Overdue receivables - 3 days after confirmation']), 'Dopłata', None))
    df.loc[(df['Type of arrears'] == 'Zaliczka') & (df['Confirmed'] == 'CANCELLED'), 'Type of arrears'] = 'Dopłata'

    return df

def add_collections_status(df):
    """
    Adds column for collections department. It is done according to the rules from table from the department. 
    Only Last minute rule was changed to 1 day after check-in date.
    """
    # Zaliczka blokada
    df.loc[
        (df['Grouped agents'] == 'AGENT') & 
        (df['Type of arrears'] == 'Zaliczka') & 
        (df['Days to check-in after confirmation'] > 14) & 
        (df['Ageing'] >= 16), 'Collections status'] = 'Blokada'
    # Dopłata blokada
    df.loc[
        (df['Grouped agents'] == 'AGENT') & 
        (df['Type of arrears'] == 'Dopłata') & 
        (df['Days to check-in after confirmation'] > 14) & 
        (df['Check-in days'] <= 8), 'Collections status'] = 'Blokada'
    # Dopłata Kaczmarski
    df.loc[
        (df['Grouped agents'] == 'AGENT') & 
        (df['Type of arrears'] == 'Dopłata') & 
        (df['Days to check-in after confirmation'] > 14) & 
        (df['Ageing'] > 22) & (df['Result'] >= 500), 'Collections status'] = 'Kaczmarski'
    # Dopłata prawnik
    df.loc[
        (df['Grouped agents'] == 'AGENT') & 
        (df['Type of arrears'] == 'Dopłata') & 
        (df['Days to check-in after confirmation'] > 14) & 
        (df['Ageing'] > 22) & 
        (df['Result'] < 500), 'Collections status'] = 'Prawnik' 
    # Lasty dopłata blokada
    df.loc[
        (df['Grouped agents'] == 'AGENT') & 
        (df['Type of arrears'] == 'Dopłata') & 
        (df['Days to check-in after confirmation'] <= 14) & 
        ((today - df['Confirmation date']).dt.total_seconds() > 24 * 3600), 'Collections status'] = 'Blokada'

    return df


def group_agents(df, assign_group_NIP, assign_group_agents):
    def map_agent(row):
        NIP_number = row['NIP']
        agent_name = row['Agent']
        
        if NIP_number in assign_group_NIP:
            return assign_group_NIP[NIP_number]
        elif agent_name in assign_group_agents:
            return assign_group_agents[agent_name]
        else:
            return 'AGENT'

    df['Grouped agents'] = df.apply(map_agent, axis=1)
    return df


def ageing(df):
    """
    Adds 'Ageing' column to the DataFrame based on the classification. - It is the number of overdue days. Depending on the classification, it is calculated differently.
    For 'Unpaid advances' it is the number of days since the advance date.
    For 'Overdue receivables' it is the number of days since the total date.
    For '30/30' it is the number of days until the check-in date.
    """
    today = pd.Timestamp.today().normalize()
    df['Ageing advance'] = (today - df['Advance date']).dt.days
    df['Ageing total'] = (today - df['Date total']).dt.days

    overdue_receivables_conditions = ['Overdue receivables', 'Overdue receivables after check-in', 'Overdue payables', 'Overdue payables after check-in', 'Overdue receivables - 3 days after confirmation']

    df['Ageing'] = np.where(df['Classification'] == 'Unpaid advances', df['Ageing advance'],
                             np.where(df['Classification'].isin(overdue_receivables_conditions) , df['Ageing total'],
                                         np.where(df['Classification'] == '30/30', df['Check-in days'], None)))

    df = df.drop(columns=['Ageing advance', 'Ageing total'], axis=1)

    return df


def format_output_file(df):
    date_columns = ['Advance date', 'Check-in date', 'Date total', 'Дата окончания']
    for col in date_columns:
        df[col] = pd.to_datetime(df[col]).dt.date

    columns_to_drop = ['Advance amount missing', 'Overpayment of 30/30', 'Overdue receivables amount', 'Payables in term amount', 'col1', 'col2', 'col3', 'col4', 'Days from confirmation', '$', 'К оплате', 'Дата окончания', 'Программа', 'Overdue receivables after check-in amount', 'Тур']#, 'Check-in days']
    
    # df.drop(columns=columns_to_drop, axis=1, inplace=True)

    df['Result'] = df['Result'].apply(lambda x: float("{:.2f}".format(float(x))))
    df['Debt'] = df['Debt'].apply(lambda x: float("{:.2f}".format(float(x))))
    df['Paid'] = df['Paid'].apply(lambda x: float("{:.2f}".format(float(x))))
    df['Catalogue price'] = df['Catalogue price'].apply(lambda x: float("{:.2f}".format(float(x))))
    df['Advance amount'] = df['Advance amount'].apply(lambda x: float("{:.2f}".format(float(x))))
    df['% Advance'] = df['% Advance'].apply(lambda x: float("{:.2f}".format(float(x))))

    df.drop(columns=columns_to_drop, axis=1, inplace=True)

    df = df[['№ Reservation', 'Type of arrears', 'Result', 'Collections status', 'Reservation status', 'Classification', 'Check-in date', 'Check-in days', 'Total days', 'Advance days', 'Confirmation date', 'Date total','Catalogue price', 'Paid', 'Debt', 'Commission','% Advance', 'Advance amount', 'Advance date', 'Confirmed', 'Agent', 'Grouped agents', 'Email', 'Mobile phones','NIP', 'Ageing', 'Expected Payment Date (samo)']]

    return df

def convert_to_decimal(value):
    try:
        return Decimal(str(value))
    except Exception:
        return print(f"Error: {value}")
    
def round_decimal(value):
    return Decimal(str(value)).quantize(Decimal('0.01'), rounding=ROUND_HALF_UP)

def round_decimal_down(value):
    return Decimal(str(value)).quantize(Decimal('0.01'), rounding=ROUND_DOWN)

def save_to_excel(df1, df2, output_path):
    with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
        df1.to_excel(writer, sheet_name='Details', index=False)
        workbook = writer.book
        worksheet = writer.sheets['Details']

        for i, col in enumerate(df1.columns):
            max_len = max(df1[col].astype(str).map(len).max(), len(col)) + 2  
            worksheet.set_column(i, i, max_len)

        number_format = workbook.add_format({'num_format': '# ### ##0.00', 'align': 'right'})


        worksheet.set_column('E:E', 15, number_format)
        worksheet.set_column('F:F', 15, number_format) 
        worksheet.set_column('G:G', 15, number_format) 
        worksheet.set_column('O:O', 15, number_format) 

        (max_row, max_col) = df1.shape
        column_settings = [{'header': column} for column in df1.columns]

        worksheet.add_table(0, 0, max_row, max_col - 1, {
            'columns': column_settings,
            'name': 'MyTable',  
            'style': 'Table Style Medium 9',
            'autofilter': True,  
        })

        header_format = workbook.add_format({
            'align': 'left',
            'bold': True,   
            'text_wrap': True  
        })

        for col_num, col_name in enumerate(df1.columns):
            worksheet.write(0, col_num, col_name, header_format)

# def update_email(df, update_email):
#     df['Email'] = df['NIP'].map(update_email).fillna(df['Email'])
#     return df

def update_email(df, email_dict):
    df['NIP'] = df['NIP'].astype(str)
    df['Email'] = df['NIP'].map(email_dict).fillna(df['Email'])
    return df

def load_and_prepare_data(file_path, usecols_list, dtype_dict, new_column_names_dict):
    df = pd.read_excel(file_path, sheet_name=SHEET_NAME, skiprows=3)
    df = adjust_excel_headers(file_path, SHEET_NAME, usecols_list)
    df = rename_columns(df, new_column_names_dict)
    columns_to_convert = ['Catalogue price', 'Paid', 'Debt', '% Advance']
    df[columns_to_convert] = df[columns_to_convert].map(convert_to_decimal)
    return df

# def classify_and_calculate(df,):


def process_excel_file(file_path, df_prev):
    """
    Opens an Excel file, processes it, and saves the output to a new file.

    Args: file_path (str): The path of the input Excel file.
        df_prev (pd.DataFrame): The DataFrame from the previous report.
    """
    sheet_name = 'all rezervation'
    file_name = os.path.basename(file_path)

    output_file_name = f"output_{file_name}"
    output_path = os.path.join(OUTPUT_DIRECTORY, output_file_name)

    dtype_dict = {
        'Название': str, 
        'ИНН': str  
    }
    usecols_list = ['№ заявки','Подтверждено','Дата подтверждения/неподтверждения','$','Дата предполагаемой оплаты','Дата предполагаемой частичной оплаты','% частичной оплаты','По каталогу','К оплате','Оплачено','Долг','Название','Тур', 'Дата/время создания','Дата начала','Дата окончания','Email','Программа','Мобильные телефоны','ИНН','Сумма']
    # usecols_list = ['Client Group', 'Type of debt',	'Overdue days',	'Date to pay', 'Reporting date', 'Total ', 'Advance', ' % of payment', '30 Days to fly', 'Overpayment', 'SUM', '№ Reservation', 'Подтверждено', 'Дата подтверждения/неподтверждения',	'$', 'Дата предполагаемой оплаты', 'Дата предполагаемой частичной оплаты', '% частичной оплаты', 'По каталогу',	'К оплате',	'Оплачено',	'Долг',	'Agent', 'Тур Название', 'Дата начала',	'Дата окончания', 'Email', 'Программа',	'Мобильные телефоны', 'ИНН', 'Комиссия']

    new_column_names = {
        '№ заявки':'№ Reservation',
        'Подтверждено': 'Confirmed',
        'Дата подтверждения/неподтверждения': 'Confirmation date',
        'Дата предполагаемой частичной оплаты': 'Advance date',
        '% частичной оплаты': '% Advance',
        'По каталогу':'Catalogue price',
        'Оплачено':'Paid',
        'Долг':'Debt',
        'Дата начала':'Check-in date', 
        'Название': 'Agent',
        'ИНН': 'NIP',
        'Сумма': 'Commission',
        'Мобильные телефоны': 'Mobile phones',
        'Статус': 'Reservation status',
        'Дата/время создания': 'Receiving of reservation',
        'Туp': 'Tour',
        'Дата предполагаемой оплаты': 'Expected Payment Date (samo)',

    }

    assign_group_NIP = {
    '9570778385': 'WAKACJE.Pl',
    '5252399530': 'FOSTERTRAVEL.PL',
    '5423401648': '24HOLIDAY',
    '7010891205': 'FLY.PL',
    '6452536937': 'HOLIDAYCAFE.PL',
    '1182078094': 'JET TRAVEL',
    '6112635474': 'TRAVEL SHOPS',
    '8971652554': 'TRAVELPLANET.PL',
    '9542736185': 'ANEX',
    # '100379260100003': 'ANEX'
    }

    assign_group_agents = {
        'GTS Poland': 'BLOCK',
        'Anex Tourism Worldwide DMCC': 'BLOCK',
        'POLAND OREX': 'ANEX',
        'Sun & Fun (Blue-style)': 'BLOCK',
        'JoinUP': 'BLOCK',
        'ANEX SERVICES ANTALYA AIRPORT RUS': 'TICKET ONLY',
        'Anex Tour Egypt': 'TICKET ONLY',
        'Anex Tour Spain': 'TICKET ONLY',
        'Anexservices Turizm Org.Tas.Tic.A.S.':	'TICKET ONLY',
        'ANEX TOUR_Ersin': 'TICKET ONLY',
        'Anex Tour Greece': 'TICKET ONLY'
    }

    dict_update_email = {
        '5771973246': 'biuro@denartravel.pl', #denar travel
        '5681438696': 'biuro@denartravel.pl', #denar travel
        '6292490088': 'szymon@globnet.com.pl', #due travel
        '7511517972': 'ksiegowosc@traveliada.pl', #traveliada
        '1180067086': 'ksiegowosc@portspin.pl', #portspin
        '6222839025': 'ostrow@planettour.pl', #els-glogow
        '9512559100': 'rozliczenia@sunwaytravel.pl', #sunway travel
        '8392716417': 'rozliczenia@sunwaytravel.pl', #sunway travel
        '6793203378': 'rozliczenia@nakanikuly.pl; awaryjnie można pisać do: viktoria@nakanikuly.com.ua, dima@nakanikuly.com.ua', #nakanikuly
        '6840004718': 'wieslaw.orlinski@radtur.pl', #radtur
        '9731016581': 'rozliczenia@odkrywca.net', #odkrywca
        '8272281187': 'mkslawinska@gmail.com', #slawia
        '6412171257': 'rezerwacje@beckermanntravel.pl, jolanta.beckermann@wp.pl', #f.h.u global jr beckermann
        '7743231048': 'karolina.zawodniak@urlopy.pl', #centrum usług turystycznych wyszogrodzka
        '6452447808': 'tomasz.kosz@perfect-holidays.pl oraz kamil.szot@perfect-holidays.pl',  #perfect holidays
        '5791808645': 'wakacje@muranotravel.pl oraz wycieczki_murano@o2.pl', #murano travel
        '8351615642': 'm.dahi@aishatravel.pl oraz faktury@aishatravel.pl', #soraya
        '6521167415': 'rozliczenia@latajmyrazem.pl', #latajmy razem
        '6641901065': 'biuro@travel-market.pl (tylko)', #travel market
        '8971796081': 'lukasz.zajac@orlica.pl', #podroz orlica
        '6341888171': 'biuro@olivierstravels.pl; mikolow@olivierstravels.pl', #oliviers travels
        '8222275948': 'ewcia.orzechowska@gmail.com', #ewa
        '9691079114': 'rezerwacje@kolibri.com.pl', #kolibri
        '6971509293': 'ksiegowosc@eurotop.leszno.pl', #eurotop
        '7891573450': 'bok@wymarzonewakacje.eu', #wymarzone wakacje
        '8381880739': 'biuro@lentravel.pl', #PM SP. Z O.O.
        '8722093280': 'damian@chodaktravel.pl', #chodak travel
        '7831766932': 'platnosci.turystyka@oskar.com.pl',#BIURO PODRÓŻY OSKAR SP.Z.O.O.
        '8371553826': 'm.raczkowski@protravel.pl; piekary@protravel.pl', #PROTRAVEL MICHAL RACZKOWSKI
        '9290095840': 'szpak@bajecznewakacje.pl' #SZPAK TRAVEL - ZIELONA GORA
    }

    dict_update_phone = {
        '6550008889': '413781733 (płatności), ewentualnie 604188746', #l-tour
        '6641901065': '501246878 (płatności obie lokalizacje)', #travel market
        '9591748403': '505775291', #AFRICA TRAVEL
        '7861597598': '798320676', #Podróży swiat
        '8971796081': '606424968', #podroz orlica
        '8262189296':'Edyta 605575585', #dream travel
        '7891573450': '791691291' #wymarzone wakacje

    }

    columns = ['№ заявки', 'Подтверждено', 'Дата подтверждения/неподтверждения', '$', 'Дата предполагаемой оплаты', 'Дата предполагаемой частичной оплаты', '% частичной оплаты', 'По каталогу', 'К оплате', 'Оплачено', 'Долг', 'Название', 'Тур', 'Дата/время создания','Дата начала', 'Дата окончания', 'Email', 'Программа', 'Мобильные телефоны', 'ИНН', 'Сумма', 'Статус','col1', 'col2', 'col3', 'col4']
    df = adjust_excel_headers(file_path, sheet_name, columns)
    df = rename_columns(df, new_column_names)
    df['Confirmed'] = df['Confirmed'].replace({'Да': 'YES', 'Нет': 'NO'})
    df['Reservation status'] = df['Reservation status'].replace({'Не оплачено': 'Not paid', 'Неоплаченный штраф': 'Not paid penalty', 'Опл. штраф': 'Paid penalty', 'Оплачено': 'Paid', 'Отменено': 'Canceled', 'Част. оплата': 'Partly paid', 'Част. оплата штрафа': 'Partly paid penalty'})
    columns_to_convert = ['Catalogue price', 'Paid', 'Debt', '% Advance']
    df[columns_to_convert] = df[columns_to_convert].map(convert_to_decimal)
    df = classify_rows(df)
    df = calculate_result(df)
    df = group_agents(df, assign_group_NIP, assign_group_agents)
    df = add_receivables_to_30(df)
    df = ageing(df)
    df = add_type_of_arrears(df)
    df = add_collections_status(df)
    today = pd.to_datetime(datetime.now().date())
    df['Check-in date'] = pd.to_datetime(df['Check-in date'], errors='coerce')
    df_after_check_in = df[(df['Check-in date'] <= today) & (df['Debt'] > 0)]
    df = update_email(df, dict_update_email)
    df = format_output_file(df)
    df.rename(columns={'№ Reservation': 'Nr rezerwacji', 'Type of arrears': 'Typ zaległości', 'Result': 'Suma', 'Collection status': 'Status windykacji'}, inplace=True)
    df = analyse_data(df, df_prev)

    save_to_excel(df, df_after_check_in, output_path)

    return output_path

if __name__ == "__main__":
    import sys
    global date
    # if len(sys.argv) > 1: sys.argv[1]

    today = datetime.today()
    today_str = today.strftime("%m%d")

    if today.weekday() == 0:
        last_day = today - timedelta(days=3)
    else:
        last_day = today - timedelta(days=1)

    last_day_str = last_day.strftime("%m%d")
    # last_day_str = '1108'
    # today_str = '1101'
    prev_file = OUTPUT_DIRECTORY + 'output_' + last_day_str + ' DEBT COLLECTION REPORT.xlsx'
    print(prev_file)
    # prev_file = 'C:\\analyze\\calculate-classification\\arkusze_wyjsciowe\\output_1031 DEBT COLLECTION REPORT.xlsx' # Dni wolne
    df_prev = pd.read_excel(prev_file, sheet_name='Details')
    file_path = INPUT_DIRECTORY + today_str +' DEBT COLLECTION REPORT.xlsx'
    # file_path = 'C:\\analyze\\calculate-classification\\arkusze_wyjsciowe\\output_1101 DEBT COLLECTION REPORT.xlsx' # Dni wolne
    date = file_path.split('\\')[-1].split()[0]
    process_excel_file(file_path, df_prev)
