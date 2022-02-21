import pandas as pd
from openpyxl import load_workbook

def deletes_user(url, i, sheet):
    book = load_workbook(url)
    ws = book[sheet]
    ws.delete_rows(i,1)
    book.save(url)

def read_user(url, sheet):
    read = pd.read_excel(url, sheet_name=sheet, index_col=False)
    return read

def delete_user():
    while True:
        all_user = read_user('./account.xlsx', 'User')
        print('Select which user you want to delete')
        print(all_user)
        choose_delete = int(input(': '))
        if choose_delete <= len(all_user):
            deletes_user('./account.xlsx', choose_delete+1, 'User')
            deletes_user('./bank.xlsx', choose_delete+1, 'Bank account')
            deletes_user('./bank.xlsx', choose_delete+1, 'Credit')
            deletes_user('./bank.xlsx', choose_delete+1, 'Transfer history')
        else:
            break
        print('Would you like to continue? 1.Yes 2.No')
        choose_exit = int(input(':'))
        if choose_exit == 2:
            break