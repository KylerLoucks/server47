import time, calendar
from datetime import datetime
import traceback # stack trace
from td.client import TDClient
import openpyxl # excel library
from data import CONSUMER_KEY, REDIRECT_URI, JSON_PATH # Import from data.py
from secretunstageddata import TD_ACCOUNT # TODO SET ENVIRONMENT VARIABLES




#pip install td-ameritrade-python-api
#https://www.youtube.com/watch?v=8N1IxYXs4e8

#======================================
#           TD AMERITRADE
#======================================
td_client = TDClient(client_id= CONSUMER_KEY, redirect_uri= REDIRECT_URI, credentials_path= JSON_PATH)
td_client.login() # Authenticate. This needs to be done every 90 days.


transactions = td_client.get_transactions(account=TD_ACCOUNT, transaction_type='BUY_ONLY') # get the transactions that are BUY ORDERS ONLY

accountData = td_client.get_accounts(account=TD_ACCOUNT, fields=['orders']) # get the account data
positionsData = td_client.get_accounts(account=TD_ACCOUNT, fields=["positions"])
# account value of TD Ameritrade account
accountValue = accountData['securitiesAccount']['initialBalances']['accountValue'] # array list with dics within dics


'''
Get a list of all the SYMBOLs of the positions that are held
'''
def getOwnedPositionSymbols():
    ownedPositions = []
    i = 0
    while i < len(positionsData["securitiesAccount"]["positions"]):
        if (str(positionsData["securitiesAccount"]["positions"][i]["instrument"]["symbol"]) != "MMDA1" ): # exclude MMDA1 (money stored on account, but not invested)
            ownedPositions.append(positionsData["securitiesAccount"]["positions"][i]['instrument']['symbol'])
            #print(str(positionsData["securitiesAccount"]["positions"][i]['instrument']['symbol']))
        i+=1
    return ownedPositions



'''
Convert a date string in the format "MM/DD/YYYY" to UTC Epoch
Variables: date - the date string
Epoch time in miliseconds for the final trading day of 2020 (GMT Thursday, December 31, 2020 6:00:00 AM)
'''
def convertUTCToEpochMiliseconds(date):
    time_var = " 06:00:00"
    dt_obj = datetime.strptime((date + time_var), "%m/%d/%Y %H:%M:%S") # parse string to datetime object
    epoch = calendar.timegm(dt_obj.utctimetuple())
    return epoch * 1000 # multiply result by 1000 to account for miliseconds

def getPriceHistory(symbol, startdate, enddate):
    startdate = str(convertUTCToEpochMiliseconds(startdate))
    enddate = str(convertUTCToEpochMiliseconds(enddate))
    price_history = td_client.get_price_history(symbol=symbol, start_date=startdate, end_date=enddate)

    print(price_history)



def getCurrentPositionMarketValue(symbol):
    i = 0
    while i < len(positionsData['securitiesAccount']['positions']):
        if symbol == positionsData["securitiesAccount"]['positions'][i]['instrument']['symbol']: # if the string entered matches the position index symbol string
            return positionsData["securitiesAccount"]['positions'][i]['marketValue'] # return the current market value
        i+=1

#======================================
#           EXCEL WORK BOOK
#======================================
# | FIELDS | #
workBookPath = "./excelworkbooktest/TD Ameritrade stonks.xlsx"
excelWorkBook = openpyxl.load_workbook(workBookPath)

ws_transactions = excelWorkBook['Transactions']     # Transactions worksheet in TD Ameritrade Stonks.xlxs
ws_contributed = excelWorkBook['$Contributed$']     # $Contributed$ worksheet in TD Ameritrade Stonks.xlxs
ws_portfolio = excelWorkBook['2021 Portfolio']      #  Portfolio worksheet in TD Ameritrade Stonks.xlxs
ws_position_data = excelWorkBook['Position Data']   #  Position Data worksheet in TD Ameritrade Stonks.xlxs


'''
Update the 'Position Data' excel worksheet with all owned symbol data (bid price, last price, etc...)
'''
def updateStockData():
    quotes = td_client.get_quotes(getOwnedPositionSymbols())
    values = list(quotes.values()) ## list of all dictionary key values (bidPrice, etc.)
    index = 1
    for stonk in values:
        index+=1      
        #print(str(stonk['symbol']) + " Bid Price changed from: " + str(ws_position_data.cell(row=index, column =2).value) + " to: " + str(stonk['bidPrice']))
        print(str(stonk['symbol']) + " Last Price changed from: " + str(ws_position_data.cell(row=index, column =4).value) + " to: " + str(stonk['lastPrice']))
        ws_position_data.cell(row=index, column=1, value=str(stonk['symbol'])).number_format = '$#,##0.00' # update symbol cells
        ws_position_data.cell(row=index, column=2, value=float(stonk['bidPrice'])).number_format = '$#,##0.00' # update bidPrice cells
        ws_position_data.cell(row=index, column=3, value=float(stonk['askPrice'])).number_format = '$#,##0.00' # update askPrice cells
        ws_position_data.cell(row=index, column=4, value=float(stonk['lastPrice'])).number_format = '$#,##0.00' # update lastPrice cells
        ws_position_data.cell(row=index, column=5, value=float(stonk['openPrice'])).number_format = '$#,##0.00' # update openPrice cells
        ws_position_data.cell(row=index, column=6, value=float(stonk['highPrice'])).number_format = '$#,##0.00' # update highPrice cells
        ws_position_data.cell(row=index, column=7, value=float(stonk['lowPrice'])).number_format = '$#,##0.00' # update lowPrice cells
        ws_position_data.cell(row=index, column=8, value=float(stonk['closePrice'])).number_format = '$#,##0.00' # update closePrice cells

        excelWorkBook.save(workBookPath)


'''
Update total account value 
'''
def updateAccountValue():
    currentYear = datetime.now()
    day = currentYear.date()
    year = int(day.strftime("%Y"))
    try:
        print("Updated " + str(ws_portfolio) + " 'Account Value' from: " + str(ws_portfolio.cell(row=2, column=8).value) + " to: $" + str(accountValue))
        ws_portfolio.cell(row=2, column=8, value=accountValue) # update total account value cell

        if (year == 2021):
            print("Updated " + str(ws_contributed) + " 'Account Value' from: " + str(ws_contributed.cell(row=2, column=4).value) + " to: $" + str(accountValue))
            ws_contributed.cell(row=2, column=4, value=accountValue) # update YTD account value cell
        if (year == 2022):
            print("Updated " + str(ws_contributed) + " 'Account Value' from: " + str(ws_contributed.cell(row=2, column=8).value) + " to: $" + str(accountValue))
            ws_contributed.cell(row=2, column=8, value=accountValue) # update YTD account value cell

        # save the excel file
        excelWorkBook.save(workBookPath)
    except:
        #print(e)
        traceback.print_exc() # stacktrace
        input("There was an error updating the total account value. Make sure the spreadsheet isn't already opened. Press Enter to continue...")




'''
Convert the date format of example: 2021-08-13T16:20:10+0000 -> 08/13/2021
'''
def convertAnnoyingDateFormat(date):
    oldDateFormat = datetime.strptime(str(date), '%Y-%m-%dT%H:%M:%S%z')
    parsedDate = datetime.strftime(oldDateFormat, '%m/%d/%Y')
    return parsedDate
        
'''
Update the 'Transactions' sheet with data from every transaction in the 'transactions' dictionary
'''
def updateTransactions():
    try:
        print("Updated the " + str(ws_transactions) + " with the following transaction data: ")
        i = 0
        while i < len(transactions): # while i < the size of the transactions list
            id = int(transactions[i]['orderId']) # grab the ID from the list
            symbol = str(transactions[i]['transactionItem']['instrument']['symbol'])
            date = str((transactions[i]['transactionDate']))
            date = convertAnnoyingDateFormat(date) # convert the annoying tedious ameritradious date format to a readable format.
            amount_paid = float(transactions[i]['netAmount'])
            share_price = float(transactions[i]['transactionItem']['price'])
            shares = int(transactions[i]['transactionItem']['amount'])

            if (amount_paid < 0): # If the transaction was a "Buy"
                amount_paid = -1 * amount_paid # Convert the negative 'transaction' to a positive float.
            else:
                shares = -1 * shares # subtract the amount of 'shares' you own

            index = i+2
            ws_transactions.cell(row=index, column=1, value=id) # update ID cells
            ws_transactions.cell(row=index, column=2, value=symbol) # update symbol cells
            ws_transactions.cell(row=index, column=3, value=date) # update date cells
            ws_transactions.cell(row=index, column=4, value=share_price).number_format = '$#,##0.00' # update share_price cells
            ws_transactions.cell(row=index, column=5, value=shares) # update shares cells
            ws_transactions.cell(row=index, column=6, value=amount_paid).number_format = '$#,##0.00' # update amount cells
            
            print("ID: " + str(id) + " SYMBOL: " + str(symbol) + " DATE: " + str(date) + " SHARE PRICE: " + str(share_price) + " SHARES: " + str(shares) + " AMOUNT PAID: " + str(amount_paid))
            i+=1
        # save the excel file
        excelWorkBook.save(workBookPath)
    except:
        traceback.print_exc() # stacktrace
        input("There was an error updating the transactions. Make sure the spreadsheet isn't already opened. Press Enter to continue...")




def recursiveUpdate():
    updateStockData()
    print("Updated Transactions! \n")
    time.sleep(10)
    recursiveUpdate() #recursion

print("Do you want to run continuously? (Press 1), otherwise, press (Enter)...")
nonStopRun = input().lower()
if nonStopRun == "1":
    recursiveUpdate()
else:
    #print(transactions)
    updateTransactions()
    updateAccountValue()
    updateStockData()

    #print(getOwnedPositionSymbols())

    input("Updated! Press Enter to close...")


