from bs4 import BeautifulSoup
import numpy as np
import pandas as pd
from selenium import webdriver

russell_3000_csv = pd.read_csv('Russell 3000 - 2019.csv')

tickers = russell_3000_csv['Ticker']
tickers = tickers.tolist()

columns = [
    'ticker',
    'PREV_CLOSE-value',
    'OPEN-value',
    'BID-value',
    'ASK-value',
    'DAYS_RANGE-value',
    'FIFTY_TWO_WK_RANGE-value',
    'TD_VOLUME-value',
    'AVERAGE_VOLUME_3MONTH-value',
    'MARKET_CAP-value',
    'BETA_5Y-value',
    'PE_RATIO-value',
    'EPS_RATIO-value',
    'EARNINGS_DATE-value',
    'DIVIDEND_AND_YIELD-value',
    'EX_DIVIDEND_DATE-value',
    'ONE_YEAR_TARGET_PRICE-value'
]

non_numeric_columns = [
    'ticker',
    'BID-value',
    'ASK-value',
    'DAYS_RANGE-value',
    'FIFTY_TWO_WK_RANGE-value',
    'MARKET_CAP-value',
    'EARNINGS_DATE-value',
    'DIVIDEND_AND_YIELD-value',
    'EX_DIVIDEND_DATE-value'
]

dataframe = pd.DataFrame(columns=columns)

for ticker in tickers:
    url = f'https://finance.yahoo.com/quote/{ticker}/'

    driver = webdriver.Chrome()
    driver.get(url)
    html = driver.execute_script('return document.body.innerHTML;')
    soup = BeautifulSoup(html, 'lxml')
    driver.close()

    row = []

    for num, each in enumerate(columns):
        if num == 0:
            row.append(ticker)
        else:
            element = soup.find('td', {'data-test': each})
            if element is None:
                row.append(np.nan)
            else:
                row.append(element.text)

    df_row = dict(zip(columns, row))

    dataframe = dataframe.append(df_row, ignore_index=True)

dataframe['FIFTY_TWO_WK_RANGE-value'] = dataframe['FIFTY_TWO_WK_RANGE-value'].fillna(value=' - ')
split_data = dataframe['FIFTY_TWO_WK_RANGE-value'].str.split(' - ', n=1, expand=True)
dataframe['52 Week Low'] = split_data[0]
dataframe['52 Week High'] = split_data[1]

for each in range(0, len(dataframe.columns)):
    if dataframe.iloc[:,each].name in non_numeric_columns:
        continue
    dataframe.iloc[:,each] = pd.to_numeric(dataframe.iloc[:,each].str.replace(',', ''), errors='coerce')

dataframe.to_excel('invest.xlsx')
