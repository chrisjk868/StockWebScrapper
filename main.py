import openpyxl as xl
import yfinance as yf
import datetime as dt


def getTickers():
    ws = xl.load_workbook('HK_2021.xlsx', data_only=True).active
    processTickers = dict()
    for row in range(2, ws.max_row):
        current_cell = "{}{}".format('B', row)
        processTickers[ws[current_cell].value + ' ' + current_cell] = 0
    return processTickers


def getStockPrices():
    tickerSymbols = [*getTickers()]
    for i in range(0, len(tickerSymbols)):
        tickerSymbols[i] = tickerSymbols[i].split(' ')[0]
    stocksInfo = yf.Tickers(tickerSymbols)
    print(stocksInfo)

def main():
    getStockPrices()


if __name__ == '__main__':
    main()

