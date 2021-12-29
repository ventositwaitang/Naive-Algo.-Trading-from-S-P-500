import numpy as np
import pandas as pd
import math
import requests
import xlsxwriter
from scipy import stats

stocks = pd.read_csv('sp_500_stocks.csv')
from secrets import IEX_CLOUD_API_TOKEN 

symbol = 'AAPL'
api_url = f'https://sandbox.iexapis.com/stable/stock/{symbol}/stats?token={IEX_CLOUD_API_TOKEN}'
data = requests.get(api_url).json()
#print(data['year1ChangePercent'])

def chunks(lst, n):
    """Yield successive n-sized chunks from lst."""
    for i in range(0, len(lst), n):
        yield lst[i:i + n]   
        
symbol_chunks = list(chunks(stocks['Ticker'], 100))
symbol_strings = []
for i in range(len(symbol_chunks)):
    symbol_strings.append(','.join(symbol_chunks[i]))
#    print(symbol_strings[i])

col = ['Ticker', 'Price', 'One-Year Price Return', 'Number of Shares to Buy']

df = pd.DataFrame(columns=col)
for symbol_string in symbol_strings:
    batch_api_url = f'https://sandbox.iexapis.com/stable/stock/market/batch/?types=quote,stats&symbols={symbol_string}&token={IEX_CLOUD_API_TOKEN}'
    # ↑一次過攞曬symbol_string，唔好再逐個symbol for個api返黎
    data = requests.get(batch_api_url).json()
    for symbol in symbol_string.split(','):
        df = df.append(pd.Series([symbol,data[symbol]['quote']['latestPrice'],data[symbol]['stats']['year1ChangePercent'],'N/A'],index=col),ignore_index=True)

df.sort_values('One-Year Price Return', ascending = False, inplace = True)
df = df[:50]
df.reset_index(drop = True, inplace = True)# pandas要經常性 inplace = True

def portfolio_input():
    global portfolio_size
    portfolio_size = input('Enter the size of your portfolio:')
    #如果入錯點算？ try except
    try:
        float(portfolio_size)
    except ValueError:
        print('That is not a number! \nPlease try again:')
        portfolio_size = input('Enter the size of your portfolio:')

portfolio_input()
print(portfolio_size)

position_size = float(portfolio_size)/len(df.index)
for i in range(len(df)):
    df.loc[i,'Number of Shares to Buy'] = math.floor(position_size/df['Price'][i])
print(df)

hqm_col = ['Ticker', 'Price', 'Number of Shares to Buy', 
               'One-Year Price Return', 
               'One-Year Return Percentile', 
               'Six-Month Price Return', 
               'Six-Month Return Percentile',
               'Three-Month Price Return',
               'Three-Month Return Percentile', 
               'One-Month Price Return',
               'One-Month Return Percentile',
               'HQM Score'
          ]
hqm_df = pd.DataFrame(columns= hqm_col)

for symbol_string in symbol_strings:
    batch_api_url = f'https://sandbox.iexapis.com/stable/stock/market/batch/?types=quote,stats&symbols={symbol_string}&token={IEX_CLOUD_API_TOKEN}'
    # ↑一次過攞曬symbol_string，唔好再逐個symbol for個api返黎
    data = requests.get(batch_api_url).json()
    for symbol in symbol_string.split(','):
        hqm_df = hqm_df.append(pd.Series([symbol,data[symbol]['quote']['latestPrice'],'N/A',
                                          data[symbol]['stats']['year1ChangePercent'],
                                          'N/A',
                                          data[symbol]['stats']['month6ChangePercent'],
                                          'N/A',
                                          data[symbol]['stats']['month3ChangePercent'],
                                          'N/A',
                                          data[symbol]['stats']['month1ChangePercent'],
                                          'N/A',
                                          'N/A'
                                         ],index=hqm_col),ignore_index=True)

time_periods = ['One-Year', 'Six-Month', 'Three-Month', 'One-Month']
hqm_df = hqm_df.fillna(value = np.nan) # ★任何計算都要numpy將N/A填充為np.nan可計算: .fillna(np.nan)

for row in hqm_df.index:
    for time_period in time_periods:
        hqm_df.loc[row, f'{time_period} Return Percentile'] = stats.percentileofscore(hqm_df[f'{time_period} Price Return'], hqm_df.loc[row, f'{time_period} Price Return'])/100
        # percentileofscore(col,score)指成列數≤score的比例，如重複：搵平均值of(不包重複數比例+包重複數比例)/2

from statistics import mean

for row in hqm_df.index:
    momentum_percentiles = []
    for time_period in time_periods:
        momentum_percentiles.append(hqm_df.loc[row,f'{time_period} Return Percentile'])
    hqm_df.loc[row,'HQM Score'] = mean(momentum_percentiles)
    
print(hqm_df)

hqm_df.sort_values('HQM Score', ascending = False, inplace = True)
hqm_df.reset_index(drop = True, inplace = True)
hqm_df  = hqm_df[:50]
print(hqm_df)

portfolio_input()

position_size = float(portfolio_size)/len(hqm_df.index)
for row in hqm_df.index:
    hqm_df.loc[row,'Number of Shares to Buy'] = math.floor(position_size/hqm_df.loc[row,'Price'])
ptint(hqm_df)

writer = pd.ExcelWriter('momentum_strategy.xlsx', engine='xlsxwriter')
hqm_df.to_excel(writer, sheet_name='Momentum Strategy', index = False)

##########
background_color = '#0a0a23'
font_color = '#ffffff'

string_template = writer.book.add_format(
        {
            'font_color': font_color,
            'bg_color': background_color,
            'border': 1
        }
    )

dollar_template = writer.book.add_format(
        {
            'num_format':'$0.00',
            'font_color': font_color,
            'bg_color': background_color,
            'border': 1
        }
    )

integer_template = writer.book.add_format(
        {
            'num_format':'0',
            'font_color': font_color,
            'bg_color': background_color,
            'border': 1
        }
    )

percent_template = writer.book.add_format(
        {
            'num_format':'0.0%',
            'font_color': font_color,
            'bg_color': background_color,
            'border': 1
        }
    )
###############



column_formats = {'A': ['Ticker',string_template],
               'B': ['Price',dollar_template],
               'C': ['Number of Shares to Buy',integer_template] ,
               'D': ['One-Year Price Return', percent_template],
               'E': ['One-Year Return Percentile', percent_template],
               'F': ['Six-Month Price Return', percent_template],
               'G': ['Six-Month Return Percentile',percent_template],
               'H': ['Three-Month Price Return',percent_template],
               'I': ['Three-Month Return Percentile', percent_template],
               'J': ['One-Month Price Return',percent_template],
               'K': ['One-Month Return Percentile',percent_template],
               'L': ['HQM Score',percent_template]
              }
for column in column_formats.keys():
    writer.sheets['Momentum Strategy'].set_column(f'{column}:{column}', 22, column_formats[column][1])#formating cells
    writer.sheets['Momentum Strategy'].write(f'{column}1', 25, column_formats[column][1])# formating headers

writer.save() #記住啊

