import gspread
import pandas as pd
from oauth2client.service_account import ServiceAccountCredentials

<<<<<<< HEAD
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
    # view the data
    df_filtered = df[(df.Transaction_Type =='Journal Entry') | (df.Transaction_Type =='Bill') | (df.Transaction_Type =='Vendor Credit')]
    df_filtered = df_filtered.drop(['0', 'Class','Balance'], axis=1)


    print(df_filtered)

    df_filtered.to_csv('sheet_name', header=True, index=False)



if __name__ == '__main__':
    takeExcel()
=======
def print_hi(name):
    # Use a breakpoint in the code line below to debug your script.
    print(f'Hi, {name}')  # Press Ctrl+F8 to toggle the breakpoint.

    scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']

    # add credentials to the account
    creds = ServiceAccountCredentials.from_json_keyfile_name('indigo-bazaar-352022-11f2f6a01dd7.json', scope)

    # authorize the clientsheet
    client = gspread.authorize(creds)
    sheet = client.open('Plan de trabajo avanciencias')

# get the first sheet of the Spreadsheet

    sheet_instance = sheet.get_worksheet(0)
    sheet_instance.col_count
    ## >> 26

    # get the value at the specific cell
    print(sheet_instance.cell(col=3, row=2))
    records_data = sheet_instance.get_all_records()
    df = pd.DataFrame(records_data)
    # view the data
    print(df)

# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    print_hi('PyCharm')

# See PyCharm help at https://www.jetbrains.com/help/pycharm/
>>>>>>> origin/main
