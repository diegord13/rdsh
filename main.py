import gspread
import pandas as pd
from oauth2client.service_account import ServiceAccountCredentials
import numpy as np
from itertools import chain
import re

def takeExcel():
    scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
    # add credentials to the account
    creds = ServiceAccountCredentials.from_json_keyfile_name('indigo-bazaar-352022-d876f1cad079.json', scope)
    # authorize the clientsheet
    client = gspread.authorize(creds)
    sheet = client.open('Copia de Copy of 2022 Apr COGS')

    sheet_instance = sheet.get_worksheet(2)

    records_data = sheet_instance.get_all_values()[5:]


    df = pd.DataFrame(records_data)
    df = df.rename(columns={0: '0',
                            1:'Date',
                            2:'Transaction_Type',
                            3:'Num',
                            4:'Name',
                            5:'Memo/Description',
                            6:'Account',
                            7:'Class',
                            8:'Amount',
                            9:'Balance'
                            })

    df = df.drop(['0', 'Class', 'Balance'], axis=1)
    df = df[(df.Transaction_Type == 'Journal Entry') | (df.Transaction_Type == 'Bill') | (df.Transaction_Type == 'Vendor Credit')]

    #df_filtered['Date'] = pd.to_datetime(df_filtered['Date'], format='%d/%m/%Y', errors='coerce')
    #df_filtered.dropna(subset=["Date"], inplace=True)
    #df = df_filtered["Memo/Description"].str.split('\n').explode('Memo/Description')

    def repeater(s):
        return list(chain.from_iterable(s.str.split('\n')))


    lens = df['Memo/Description'].str.split('\n').map(len)

    # create new dataframe, repeating or chaining as appropriate
    res = pd.DataFrame({'Date': np.repeat(df['Date'], lens),
                        'Transaction_Type': np.repeat(df['Transaction_Type'], lens),
                        'Num': np.repeat(df['Num'],lens),
                        'Name': np.repeat(df['Name'],lens),
                        'Memo/Description': repeater(df['Memo/Description']),
                        'Account': np.repeat(df['Account'],lens),
                        'Amount': np.repeat(df['Amount'],lens)
                        })
    res['Amount 1'] = res['Memo/Description'].str.extract("([\d,.]+)[^\d,.]*$")
    res['Title match'] = res['Memo/Description'].str.extract("(?i)\b")
    print(res.shape)
    res.to_csv('sheet_name', header=True, index=False)


if __name__ == '__main__':
    takeExcel()


