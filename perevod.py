import pandas as pd
import string
import random
from openpyxl import load_workbook
import datetime


book = load_workbook('./perevod_data.xlsx')
def add_perevod(url, perevod_data):
    with pd.ExcelWriter(url, engine = 'openpyxl') as writer:
        writer.book = book
        writer.sheets = dict((ws.title, ws) for ws in book.worksheets)
        perevod_data.to_excel(writer, 'Perevod', index=False)
        writer.save()

def read_perevod(url, sheet):
    read = pd.read_excel(url, sheet_name=sheet, index_col=False)
    return read

def random_code():
    chars = string.ascii_uppercase + string.ascii_lowercase + string.digits
    size = random.randint(8, 12)
    return ''.join(random.choice(chars) for x in range(size))

def perevods():
    while True:
        print('\t\t\t\tPlease choose the direction to transfer money')
        print('\t1.International\n\t2.Within Kyrgyzstan')
        
        choose_perevod = int(input(':'))
        if choose_perevod == 1:
            print('\t\t\t\tAttention! If sent by this transfer, a fee of 5% of the transfer amount is charged.')
            country = input('Please enter the country for the translation: ')
            city = input('Please enter a city: ')
            whom = input('Please enter the name of the person you want to send the transfer to: ')
            print('1.Send transfer\n2.Cancel')
            choose_send_perevod = int(input(': '))

            if choose_send_perevod == 1:
                code = random_code()
                date = datetime.datetime.now()
                saved_perevod = read_perevod('./perevod_data.xlsx', 'Perevod')
                df = pd.DataFrame(saved_perevod)
                update_perevod = df.append({'Country':country, 'City': city, 'NAME': whom, 'Date':date, 'Code':code}, ignore_index=True)
                add_perevod('./perevod_data.xlsx', update_perevod)
                print('Your money transfer',whom,'in country',country,'and in city', city, 'successfully completed!')
                print('The code to get the money:', code)
            elif choose_send_perevod == 2:
                break

        if choose_perevod == 2:
            print('\t\t\t\tAttention! If sent by this transfer, a commission of 1.5% of the transfer amount is charged.')
            city_kg = input('Please enter a city: ')
            whom_kg = input('Please enter the name of the person you want to send the transfer to: ')
            print('1.Send transfer\n2.Cancel')
            choose_send_perevod_kg = int(input(': '))
            if choose_send_perevod_kg == 1:
                code = random_code()
                date = datetime.datetime.now()
                saved_perevod = read_perevod('./perevod_data.xlsx', 'Perevod')
                df = pd.DataFrame(saved_perevod)
                update_perevod = df.append({'Country':'Kyrgyzstan', 'City': city_kg, 'NAME': whom_kg, 'Date':date, 'Code':code}, ignore_index=True)
                add_perevod('./perevod_data.xlsx', update_perevod)
                print('Your money transfer ',whom_kg,' in city ', city_kg, ' successfully completed!')
                print('The code to get the money:', code)
            elif choose_send_perevod_kg ==2:
                break


