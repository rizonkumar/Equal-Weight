import numpy as np
import pandas as pd
import requests 
import xlsxwriter
import math

stocks = pd.read_csv('sp_500_stocks.csv')
stocks

from secrets import IEX_CLOUD_API_TOKEN

symbol='AAPL'
api_url = f'https://sandbox.iexapis.com/stable/stock/{symbol}/quote?token={IEX_CLOUD_API_TOKEN}'
data = requests.get(api_url).json()
data


data['latestPrice']
data['marketCap']

my_columns = ['Ticker', 'Price','Market Capitalization', 'Number Of Shares to Buy']
final_dataframe = pd.DataFrame(columns = my_columns)
final_dataframe


final_dataframe = final_dataframe.append(pd.Series(['AAPL', data['latestPrice'], 
                                                    data['marketCap'], 'N/A'],
                                                    index = my_columns), ignore_index = True)

#print(final_dataframe)
final_dataframe

final_dataframe = pd.DataFrame(columns = my_columns)
for symbol in stocks['Ticker'][:5]:     #not to slow so first 5 stock 
    api_url = f'https://sandbox.iexapis.com/stable/stock/{symbol}/quote?token={IEX_CLOUD_API_TOKEN}'
    data = requests.get(api_url).json()
    final_dataframe = final_dataframe.append(pd.Series([symbol,data['latestPrice'],data['marketCap'],'N/A'],index = my_columns), 
                                            ignore_index = True)

#print(final_dataframe)
#final_dataframe

#Batch API Call to improve performance

def chunks(lst, n):
    """Yield successive n-sized chunks from lst."""
    for i in range(0, len(lst), n):
        yield lst[i:i + n]

symbol_groups = list(chunks(stocks['Ticker'], 100))
symbol_strings = []
for i in range(0, len(symbol_groups)):
    symbol_strings.append(','.join(symbol_groups[i]))


final_dataframe = pd.DataFrame(columns = my_columns)

for symbol_string in symbol_strings:

    batch_api_call_url = f'https://sandbox.iexapis.com/stable/stock/market/batch/?types=quote&symbols={symbol_string}&token={IEX_CLOUD_API_TOKEN}'
    data = requests.get(batch_api_call_url).json()
    for symbol in symbol_string.split(','):
        final_dataframe = final_dataframe.append(
                                        pd.Series([symbol,data[symbol]['quote']['latestPrice'],data[symbol]['quote']['marketCap'],'N/A'],index = my_columns),ignore_index = True)
        
    
print(final_dataframe)
final_dataframe

#Calculation number of shares
stock_size = input("Enter the value of your stock:")

try:
    val = float(stock_size)
except ValueError:
    print("That is not a number!! \n Try again:")
    stock_size = input("Enter the value of your stock:")

position_size = float(stock_size) / len(final_dataframe.index)
print(position_size)
for i in range(0, len(final_dataframe['Ticker'])-1):
    final_dataframe.loc[i, 'Number Of Shares to Buy'] = math.floor(position_size / final_dataframe['Price'][i])
print(final_dataframe)



#Formating our excel Sheet
writer = pd.ExcelWriter('trades.xlsx', engine='xlsxwriter')
final_dataframe.to_excel(writer, sheet_name='Trades', index = False)

background_color = 'DE354C'
font_color = '#ffffff'

string_format = writer.book.add_format(
        {
            'font_color': font_color,
            'bg_color': background_color,
            'border': 1
        }
    )

dollar_format = writer.book.add_format(
        {
            'num_format':'$0.00',
            'font_color': font_color,
            'bg_color': background_color,
            'border': 1
        }
    )

integer_format = writer.book.add_format(
        {
            'num_format':'0',
            'font_color': font_color,
            'bg_color': background_color,
            'border': 1
        }
    )


writer.sheets['Trades'].write('A1', 'Ticker', string_format)
writer.sheets['Trades'].write('B1', 'Price', string_format)
writer.sheets['Trades'].write('C1', 'Market Capitalization', string_format)
writer.sheets['Trades'].write('D1', 'Number Of Shares to Buy', string_format)



column_formats = { 
                    'A': ['Ticker', string_format],
                    'B': ['Price', dollar_format],
                    'C': ['Market Capitalization', dollar_format],
                    'D': ['Number of Shares to Buy', integer_format]
                    }

for column in column_formats.keys():
    writer.sheets['Trades'].set_column(f'{column}:{column}', 20, column_formats[column][1])
    writer.sheets['Trades'].write(f'{column}1', column_formats[column][0], string_format)  
writer.save()
 