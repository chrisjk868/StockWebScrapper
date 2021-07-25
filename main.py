import openpyxl as xl
import yfinance as yf
import pandas as pd


def getTickers(ws, col):
    processTickers = dict()
    print(ws.max_row)
    for row in range(2, ws.max_row):
        current_cell = "{}{}".format(col, row)
        processTickers[ws[current_cell].value] = ['C' + current_cell[1:], 0]
    return processTickers


def getStockPrices(stocks):
    pd.set_option('display.max_columns', None)
    tickerSymbols = [*stocks]
    stocksInfo = yf.Tickers(tickerSymbols)
    currentPrice = stocksInfo.history(period='1d')['Close'][0:3].to_numpy()[0]
    return currentPrice


def updateStockPrices(stockIndex, close):
    i = 0
    for stock in stockIndex.keys():
        stockIndex[stock][1] = close[i]
        i += 1
    return stockIndex


def writeExcel(ws, stockIndex):
    for stock in stockIndex:
        ws[stockIndex[stock][0]] = stockIndex[stock][1]


def main():
    wb = xl.load_workbook('HK_2021.xlsx', data_only=True)
    ws = wb.active
    stockDict = getTickers(ws, 'B')
    closingPrices = getStockPrices(stockDict)
    updatedIndex = updateStockPrices(stockDict, closingPrices)
    print(updatedIndex)
    writeExcel(ws, updatedIndex)
    wb.save('HK_2021.xlsx')


if __name__ == '__main__':
    main()

