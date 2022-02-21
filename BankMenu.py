from delete_user import delete_user
import setting_currency 
import perevod
import проект
import pandas as pd
def read_user(url, sheet):
    read = pd.read_excel(url, sheet_name=sheet, index_col=False)
    return read
def worker_menu():
    while True:
        print('\t\t\t\t\tMenu')
        print('1.Show client list\n2.Find client\n3.Show client name that has taken the maximum amount of credit')
        print('4.Show the name of the client who took the minimum credit amount\n5.Set the new amount to be converted\n6.Perform the transfer')
        print('7.Setting up a credit \n8.Delete user\n9.Exit')
        choose_worker_menu = int(input(':'))
        if choose_worker_menu == 1:
            users = read_user('./account.xlsx', 'User')
            print(users)
        if choose_worker_menu == 2:
            проект.searching()
        elif choose_worker_menu == 3:
            credits = read_user('./bank.xlsx', 'Credit')
            if credits['Credit'][0] != None:
                for i in range(len(credits['email'])): 
                    if max(credits['Credit']) == credits['Credit'][i]: 
                        print('\t\t\t\tMax credit user:')
                        print(credits['email'][i],':', max(credits['Credit']))
        elif choose_worker_menu == 4:
            credits = read_user('./bank.xlsx', 'Credit')
            if credits['Credit'][0] != None:
                for i in range(len(credits['email'])): 
                    if min(credits['Credit']) == credits['Credit'][i]: 
                        print('\t\t\t\tMin credit user:')
                        print(credits['email'][i],':', min(credits['Credit']))
        elif choose_worker_menu == 5:
            setting_currency.setting_convert()
        elif choose_worker_menu == 6:
            perevod.perevods()
        elif choose_worker_menu == 7:
            setting_currency.credit_setting()
        elif choose_worker_menu == 8:
            delete_user()
        elif choose_worker_menu == 9:
            print('The program is over, we look forward to your return!')
            return 6