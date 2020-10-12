# -- coding: utf-8 --
import time
import pyRofex
import xlwings as xw
from datetime import datetime

wb = xw.Book('main.xlsm')
shtLogin = wb.sheets('login')
sht = wb.sheets('main')

if not shtLogin.range('B2').value or not shtLogin.range('B3').value or not shtLogin.range('B4').value:
    sht.range('C1').value = 'Off'
    shtLogin.range('C6').value = 'Missing credentials.'
    exit()

firstRow = 5

sht.range('C1').value = 'Waiting'

pyRofex.initialize(user=shtLogin.range('B2').value,
                   password=shtLogin.range('B3').value,
                   account=shtLogin.range('B4').value,
                   environment=pyRofex.Environment.REMARKET)
shtLogin.range('B2').value = ""
shtLogin.range('B3').value = ""
shtLogin.range('B4').value = ""

instruments = pyRofex.get_all_instruments()
detailed = pyRofex.get_detailed_instruments()


def get_available_symbols():

    all_instr = pyRofex.get_all_instruments()
    instr = all_instr['instruments']
    avail_instruments = []
    for x in instr:
        avail_instruments.append(x["instrumentId"]["symbol"])

    return avail_instruments


def get_symbols_from_excel(a):
    # One param, row index from start counting on column A. Until it find a empty cell.

    avail_instruments = get_available_symbols()

    i = a
    symbols = []
    sht = xw.Book('main.xlsm').sheets('main')
    while sht.range((i, 1)).value:

        if str(sht.range((i, 1)).value) in avail_instruments:
            symbols.append(str(sht.range((i, 1)).value))

        i = i + 1

    return symbols


def market_data_handler(message):
    print(message)
    symbols = get_symbols_from_excel(firstRow)
    print(symbols)

    if message['instrumentId']['symbol'] == xw.Book('main.xlsm').sheets('main').range(((firstRow + int(symbols.index(message['instrumentId']['symbol']))), 1)).value:
        i = (firstRow + int(symbols.index(message['instrumentId']['symbol'])))
        main = xw.Book('main.xlsm').sheets('main')
        main.range((i, 2)).value = message['marketData']['LA']['price']
        main.range((i, 11)).value = datetime.fromtimestamp(message['marketData']['LA']['date'] / 1000).strftime('%H:%M')
        main.range((i, 4)).value = message['marketData']['BI'][0]['size']
        main.range((i, 5)).value = message['marketData']['BI'][0]['price']
        main.range((i, 6)).value = message['marketData']['OF'][0]['price']
        main.range((i, 7)).value = message['marketData']['OF'][0]['size']
        main.range((i, 9)).value = message['marketData']['LO']
        main.range((i, 10)).value = message['marketData']['HI']
        main.range((i, 8)).value = message['marketData']['CL']['price']


def error_handler(message):
    print("Error Message Received: {0}".format(message))


def exception_handler(e):
    print("Exception Occurred: {0}".format(e.message))


pyRofex.init_websocket_connection(market_data_handler=market_data_handler,
                                  error_handler=error_handler,
                                  exception_handler=exception_handler)

instrument = get_symbols_from_excel(firstRow)  # Instruments list to subscribe
entries = [pyRofex.MarketDataEntry.BIDS,
           pyRofex.MarketDataEntry.OFFERS,
           pyRofex.MarketDataEntry.LAST,
           pyRofex.MarketDataEntry.LOW_PRICE,
           pyRofex.MarketDataEntry.HIGH_PRICE,
           pyRofex.MarketDataEntry.CLOSING_PRICE]

pyRofex.market_data_subscription(tickers=instrument,
                                 entries=entries)

wba = xw.Book('main.xlsm')
sht = wba.sheets('main')
sht.range('C1').value = 'Running'
time.sleep(sht.range('B1').value*60)
sht.range('C1').value = 'Off'
pyRofex.close_websocket_connection()
