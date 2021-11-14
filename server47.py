from data import CONSUMER_KEY, REDIRECT_URI, JSON_PATH # Import from data.py
from secretunstageddata import TD_ACCOUNT
from td.client import TDClient
import openpyxl # excel library
from datetime import datetime
import time
import traceback # stack trace


#pip install td-ameritrade-python-api
#https://www.youtube.com/watch?v=8N1IxYXs4e8

#======================================
#           TD AMERITRADE
#======================================
td_client = TDClient(client_id= CONSUMER_KEY, redirect_uri= REDIRECT_URI, credentials_path= JSON_PATH)
td_client.login()




transactions = td_client.get_transactions(account=TD_ACCOUNT, transaction_type='BUY_ONLY') # get the transactions that are BUY ORDERS ONLY

accountData = td_client.get_accounts(account=TD_ACCOUNT, fields=['orders']) # get the account data
# account value of TD Ameritrade account
accountValue = accountData['securitiesAccount']['initialBalances']['accountValue'] # array list with dics within dics




#======================================
#           EXCEL WORK BOOK
#======================================
# | FIELDS | #
workBookPath = "D:\Python Scripts\Server47\excelworkbooktest\TD Ameritrade stonks.xlsx"
excelWorkBook = openpyxl.load_workbook(workBookPath)
transactionsWS = excelWorkBook['Transactions']# Transactions worksheet in TD Ameritrade Stonks.xlxs
contributedWS = excelWorkBook['$Contributed$']# $Contributed$ worksheet in TD Ameritrade Stonks.xlxs
ws_portfolio = excelWorkBook['2021 Portfolio']#  Portfolio worksheet in TD Ameritrade Stonks.xlxs


def updateAccountValue():
    try:
        ws_portfolio.cell(row=2, column=8, value=accountValue) # update total account value cell
        contributedWS.cell(row=2, column=4, value=accountValue) # update total account value cell
        # save the excel file
        excelWorkBook.save(workBookPath)
        print("Updated " + str(ws_portfolio) + " " + str(ws_portfolio.cell(row=2, column=8).value) + " to: " + "$" + str(accountValue))
    except:
        traceback.print_exc() # stacktrace
        input("There was an error updating the total account value. Make sure the spreadsheet isn't already opened. Press Enter to continue...")
        


def convertAnnoyingDateFormToSomethingUnderstandableByAHumanBeing(date):
    oldDateForm = datetime.strptime(str(date), '%Y-%m-%dT%H:%M:%S%z')
    parsedDate = datetime.strftime(oldDateForm, '%m/%d/%Y')
    return parsedDate


def updateTransactions():
    try:
        print("Updated the " + str(transactionsWS) + " with the following: ")
        i = 0
        while i < len(transactions): # while the size of the transactions list is < i
            id = int(transactions[i]['transactionId']) # grab the ID from the list
            symbol = str(transactions[i]['transactionItem']['instrument']['symbol'])
            date = str((transactions[i]['transactionDate']))
            date = convertAnnoyingDateFormToSomethingUnderstandableByAHumanBeing(date) # convert the annoying tedious ameritradious date form to a readable form.
            amount_paid = -1 * float(transactions[i]['netAmount']) # Convert the negative transaction to a positive float of a string.
            share_price = float(transactions[i]['transactionItem']['price'])
            shares = int(transactions[i]['transactionItem']['amount'])

            index = i+2
            transactionsWS.cell(row=index, column=1, value=id) # update ID column cells
            transactionsWS.cell(row=index, column=2, value=symbol) # update symbol column cells
            transactionsWS.cell(row=index, column=3, value=date) # update date cells
            transactionsWS.cell(row=index, column=4, value=share_price).number_format = '$#,##0.00' # update share_price cells
            transactionsWS.cell(row=index, column=5, value=shares) # update shares cells
            transactionsWS.cell(row=index, column=6, value=amount_paid).number_format = '$#,##0.00' # update amount column cells
            
            print("ID: " + str(id) + " SYMBOL: " + str(symbol) + " DATE: " + str(date) + " SHARES: " + str(shares) + " AMOUNT PAID: " + str(amount_paid) + " SHARE PRICE: " + str(share_price))
            i+=1
        # save the excel file
        excelWorkBook.save(workBookPath)
    except:
        traceback.print_exc() # stacktrace
        input("There was an error updating the transactions. Make sure the spreadsheet isn't already opened. Press Enter to continue...")
        


def sleep():
    updateTransactions()
    updateAccountValue()
    print("Updated Transactions! \n")
    time.sleep(10)
    sleep()

print("Do you want to run continuously? (Press 1), otherwise, press (Enter)...")
nonStopRun = input().lower()
if nonStopRun == "1":
    sleep()
if nonStopRun != "1":
    updateTransactions()
    updateAccountValue()
    input("Updated! Press Enter to close...")

