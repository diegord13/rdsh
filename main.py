import gspread
import pandas as pd
from oauth2client.service_account import ServiceAccountCredentials

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
