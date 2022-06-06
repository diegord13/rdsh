import gspread
import pandas as pd
from oauth2client.service_account import ServiceAccountCredentials
import numpy as np
from itertools import chain
from gspread_dataframe import get_as_dataframe, set_with_dataframe
import re

def creds():
    scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
    # add credentials to the account
    creds = ServiceAccountCredentials.from_json_keyfile_name('indigo-bazaar-352022-d876f1cad079.json', scope)
    # authorize the clientsheet
    client = gspread.authorize(creds)
    sheet = client.open('Copia de Copy of 2022 Apr COGS')
    return sheet

def takeExcel(page,skiprow):
    sheet_instance = creds().get_worksheet(page)
    records_data = sheet_instance.get_all_values()[skiprow:]
    df = pd.DataFrame(records_data)
    return df

df = pd.DataFrame(takeExcel(2,5))
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


#res['Title match'] = res['Memo/Description'].str.extract("(?i)")
df_names_titles = pd.DataFrame(takeExcel(1,0))

res["title"] = " "

s1 = df_names_titles[0]
s2 = res['Memo/Description']


s1.replace("", np.NaN, inplace=True)
s1=s1.dropna()


for i in s1.index:
    if s1[i] == np.NaN:
        continue
    #print(str(s1[i]))
    for j in s2.index:
        if str(s1[i]) in (str(s2[j])):
            res.at[j, "title"] = str(s1[i])
            print(str(s1[i]))
            print(i)
print(res)


# res['Title match'] = res['Memo/Description'].str.extract("(?i)\b")
print(res.shape)

worksheet_title = 'upload'
try:
    worksheet = creds().worksheet(worksheet_title)
except gspread.WorksheetNotFound:
    worksheet = creds().add_worksheet(title=worksheet_title, rows=1000, cols=1000)

# Write a test DataFrame to the worksheet

set_with_dataframe(worksheet, res)
res.to_csv('sheet_name', header=True, index=False)


# if __name__ == '__main__':
#     takeExcel()


