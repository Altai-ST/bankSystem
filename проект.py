import pandas as pd


def name_search(name, fullname):
    for i in range(len(fullname['Name'])):
        if (name == fullname['Name'][i]):
            print('Name:',fullname['Name'][i], 'email:', fullname['Email'][i])

def read_user(url, sheet):
    read = pd.read_excel(url, sheet_name=sheet, index_col=False)
    return read

def searching():
    while True:
        print('\t\t\t\t\tSearch')
        fullname=str(input('Enter your full name: '))
        print('1.Search\n2.Cancel')
        choose_search = int(input(':'))
        if choose_search == 1:
            list_names = read_user('./account.xlsx', 'User')
            name_search(fullname, list_names)
        elif choose_search == 2:
            break