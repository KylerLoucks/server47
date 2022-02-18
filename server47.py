import time
from datetime import datetime
import traceback # stack trace
from td.client import TDClient
import openpyxl as excel # excel library
from data import TD_ACCOUNT, CONSUMER_KEY, REDIRECT_URI, JSON_PATH # Import from data.py




#pip install td-ameritrade-python-api
#https://www.youtube.com/watch?v=8N1IxYXs4e8

#=====================================================================================================================================================================
#           TD AMERITRADE
#=====================================================================================================================================================================
TD_CLIENT = TDClient(client_id= CONSUMER_KEY, redirect_uri= REDIRECT_URI, credentials_path= JSON_PATH)
TD_CLIENT.login() # Authenticate. This needs to be done every 90 days.

# get the transactions that are BUY ORDERS ONLY
TRANSACTIONS_DICT = TD_CLIENT.get_transactions(account=TD_ACCOUNT, transaction_type='BUY_ONLY') 

# get the transactions that were dividends
DIVIDENDS_DICT = TD_CLIENT.get_transactions(account=TD_ACCOUNT, transaction_type='DIVIDEND') 

# get the account data
ACCOUNT_DATA_DICT = TD_CLIENT.get_accounts(account=TD_ACCOUNT, fields=['orders'])

# position data
POSITION_DATA_DICT = TD_CLIENT.get_accounts(account=TD_ACCOUNT, fields=["positions"])

# account value of TD Ameritrade account
ACCOUNT_VALUE = ACCOUNT_DATA_DICT['securitiesAccount']['initialBalances']['accountValue']



def get_owned_position_symbols() -> list:
    '''
    Get a list of all the SYMBOLs of the positions that are held
    '''
    owned_positions = []
    for i in range(0, len(POSITION_DATA_DICT["securitiesAccount"]["positions"])):
        if (str(POSITION_DATA_DICT["securitiesAccount"]["positions"][i]["instrument"]["symbol"]) != "MMDA1" ): # exclude MMDA1 (money stored on account, but not invested)
            owned_positions.append(POSITION_DATA_DICT["securitiesAccount"]["positions"][i]['instrument']['symbol'])
            #print(str(POSITION_DATA_DICT["securitiesAccount"]["positions"][i]['instrument']['symbol']))
    return owned_positions



# def convert_utc_to_epoch_miliseconds(date: str) -> int:
#     '''
#     Convert a date string in the format "MM/DD/YYYY" to UTC Epoch
#     Variables: date - the date string
#     Epoch time in miliseconds for the final trading day of 2020 (GMT Thursday, December 31, 2020 6:00:00 AM)
#     '''
    
#     time_var = " 06:00:00"
#     dt_obj = datetime.strptime((date + time_var), "%m/%d/%Y %H:%M:%S") # parse string to datetime object
#     epoch = calendar.timegm(dt_obj.utctimetuple())
#     return epoch * 1000 # multiply result by 1000 to account for miliseconds

# def getPriceHistory(symbol: str, startdate: str, enddate: str) -> None:
#     startdate = str(convert_utc_to_epoch_miliseconds(startdate))
#     enddate = str(convert_utc_to_epoch_miliseconds(enddate))
#     price_history = td_client.get_price_history(symbol=symbol, start_date=startdate, end_date=enddate)

#     print(price_history)


# def get_current_position_market_value(symbol: str) -> None:
#     for i in range(0, len(positions_data['securitiesAccount']['positions'])):
#         if symbol == positions_data["securitiesAccount"]['positions'][i]['instrument']['symbol']: # if the string entered matches the position index symbol string
#             return positions_data["securitiesAccount"]['positions'][i]['marketValue'] # return the current market value

#=====================================================================================================================================================================
#           EXCEL WORK BOOK
#=====================================================================================================================================================================
# | FIELDS | #
WORK_BOOK_PATH = "./excelworkbooktest/TD Ameritrade stonks.xlsx"
EXCEL_WORK_BOOK = excel.load_workbook(WORK_BOOK_PATH)

ws_transactions = EXCEL_WORK_BOOK['Transactions']     # Transactions worksheet in TD Ameritrade Stonks.xlxs
ws_contributed = EXCEL_WORK_BOOK['$Contributed$']     # $Contributed$ worksheet in TD Ameritrade Stonks.xlxs
ws_portfolio = EXCEL_WORK_BOOK['Portfolio']           # Portfolio worksheet in TD Ameritrade Stonks.xlxs
ws_position_data = EXCEL_WORK_BOOK['Position Data']   # Position Data worksheet in TD Ameritrade Stonks.xlxs
ws_dividends = EXCEL_WORK_BOOK['Dividends']


'''
Update the 'Position Data' excel worksheet with all owned symbol data (bid price, last price, etc...)
'''
def update_stock_data() -> None:
    quotes = TD_CLIENT.get_quotes(get_owned_position_symbols())
    values = list(quotes.values())      # list of all dictionary key values (bidPrice, etc.)
    index = 1
    for stonk in values:
        index+=1
        
        old_price = float(ws_position_data.cell(row=index, column=4).value)
        new_price = float(stonk['lastPrice'])
        delta = ((new_price - old_price) / old_price)
        percent = "{:.2%}".format(delta) # format 2 decimal places example: 0.0345 = 3.45
        print(f"{stonk['symbol']} Last Price changed from: ${(ws_position_data.cell(row=index, column=4).value)} to: ${stonk['lastPrice']} | change: {percent}")
        ws_position_data.cell(row=index, column=1, value=str(stonk['symbol'])).number_format = '$#,##0.00'          # update symbol cells
        ws_position_data.cell(row=index, column=2, value=float(stonk['bidPrice'])).number_format = '$#,##0.00'      # update bidPrice cells
        ws_position_data.cell(row=index, column=3, value=float(stonk['askPrice'])).number_format = '$#,##0.00'      # update askPrice cells
        ws_position_data.cell(row=index, column=4, value=float(stonk['lastPrice'])).number_format = '$#,##0.00'     # update lastPrice cells
        ws_position_data.cell(row=index, column=5, value=float(stonk['openPrice'])).number_format = '$#,##0.00'     # update openPrice cells
        ws_position_data.cell(row=index, column=6, value=float(stonk['highPrice'])).number_format = '$#,##0.00'     # update highPrice cells
        ws_position_data.cell(row=index, column=7, value=float(stonk['lowPrice'])).number_format = '$#,##0.00'      # update lowPrice cells
        ws_position_data.cell(row=index, column=8, value=float(stonk['closePrice'])).number_format = '$#,##0.00'    # update closePrice cells

    EXCEL_WORK_BOOK.save(WORK_BOOK_PATH)


'''
Update the 'Dividends' excel worksheet with all dividend transactions
'''
def update_dividend_data() -> None:
    try:
        print(f"Updated the {ws_dividends} with the following transaction data: ")

        for i in range(0, len(DIVIDENDS_DICT)):               # while i < the size of the transactions list
            id = int(DIVIDENDS_DICT[i]['transactionId'])            # grab the ID from the list
            symbol = str(DIVIDENDS_DICT[i]['transactionItem']['instrument']['symbol'])
            date = convert_annoying_date_format(str((DIVIDENDS_DICT[i]['transactionDate']))) # convert the annoying tedious ameritradious date format to a readable format.
            amount_recieved = float(DIVIDENDS_DICT[i]['netAmount'])
            
            index = i+2
            ws_dividends.cell(row=index, column=1, value=id)         # update ID cells
            ws_dividends.cell(row=index, column=2, value=symbol)     # update symbol cells
            ws_dividends.cell(row=index, column=3, value=date)       # update date cells
            ws_dividends.cell(row=index, column=4, value=amount_recieved).number_format = '$#,##0.00' # amount recieved for the dividend
            
            print(f"ID: {id} SYMBOL: {symbol} DATE: {date} DIV AMOUNT: ${amount_recieved}")
        # save the excel file
        EXCEL_WORK_BOOK.save(WORK_BOOK_PATH)
    except:
        traceback.print_exc() # stacktrace
        input("There was an error updating the transactions. Make sure the spreadsheet isn't already opened. Press Enter to continue...")

def update_portfolio() -> None:
    # Iterate through the columns and rows to look for the cell that has "Symbol" as the value
    for column in range(1, 50):
        for row in range(1, 50):
            if (ws_portfolio.cell(row=row, column=column).value == "Symbol"): # If the cell has 'Symbol' as the value
                for symbol in get_owned_position_symbols(): # For every symbol that is owned
                    row=row+1 # iterate to the row just below the "Symbol" column
                    ws_portfolio.cell(row=row, column=column, value=symbol) # Update the cells below the "Symbol" cell
                break # Break out of the loop as we've finished updating
    print(f"Updated {ws_portfolio} 'Symbols' to: {get_owned_position_symbols()}")
                

def update_account_value() -> None:
    """
    Update total Account Value
    """
    currentYear = datetime.now()
    year = int(currentYear.date().strftime("%Y"))
    try:

        # Iterate through the columns and rows
        for column in range(1, 50):
            for row in range(1, 50):
                if (ws_portfolio.cell(row=row, column=column).value == "Account Value"): # If the cell has 'Account Value' as the value
                    print(f"Updated {ws_portfolio} 'Account Value' from: {ws_portfolio.cell(row=row+1, column=column).value} to: ${ACCOUNT_VALUE}")
                    ws_portfolio.cell(row=row+1, column=column, value=ACCOUNT_VALUE) # Update the cell just below "Account Value" cell
                
        if (year == 2021):
            print(f"Updated {ws_contributed} 'Account Value' from: {ws_contributed.cell(row=2, column=4).value} to: ${ACCOUNT_VALUE}")
            ws_contributed.cell(row=2, column=4, value=ACCOUNT_VALUE)    # update YTD account value cell for 2021
        if (year == 2022):
            print(f"Updated {ws_contributed} 'Account Value' from: {ws_contributed.cell(row=2, column=9).value} to: ${ACCOUNT_VALUE}")
            ws_contributed.cell(row=2, column=9, value=ACCOUNT_VALUE)    # update YTD account value cell for 2022

        # save the excel file
        EXCEL_WORK_BOOK.save(WORK_BOOK_PATH)
    except:
        #print(e)
        traceback.print_exc() # stacktrace
        input("There was an error updating the total account value. Make sure the spreadsheet isn't already opened. Press Enter to continue...")





def convert_annoying_date_format(date: str) -> str:
    '''
    Convert the date format of example: 2021-08-13T16:20:10+0000 -> 08/13/2021
    '''
    oldDateFormat = datetime.strptime(date, "%Y-%m-%dT%H:%M:%S%z")
    parsedDate = datetime.strftime(oldDateFormat, "%m/%d/%Y")
    return parsedDate
        

def update_transactions():
    '''
    Update the 'Transactions' sheet with data from every transaction in the 'transactions' dictionary
    '''
    try:
        print(f"Updated the {ws_transactions} with the following transaction data: ")
        
        # for i = 0; i < the size of the transactions list; ++i
        for i in range(0, len(TRANSACTIONS_DICT)):    
            id = int(TRANSACTIONS_DICT[i]['orderId'])            # grab the ID from the list
            symbol = str(TRANSACTIONS_DICT[i]['transactionItem']['instrument']['symbol'])
            date = convert_annoying_date_format(str((TRANSACTIONS_DICT[i]['transactionDate']))) # convert the annoying tedious ameritradious date format to a readable format.
            amount_paid = float(TRANSACTIONS_DICT[i]['netAmount'])
            share_price = float(TRANSACTIONS_DICT[i]['transactionItem']['price'])
            shares = int(TRANSACTIONS_DICT[i]['transactionItem']['amount'])

            if (amount_paid < 0): # If the transaction was a "Buy"
                amount_paid = -1 * amount_paid # Convert the negative 'transaction' to a positive float.
            else:
                shares = -1 * shares # subtract the amount of 'shares' you own

            index = i+2
            ws_transactions.cell(row=index, column=1, value=id)         # update ID cells
            ws_transactions.cell(row=index, column=2, value=symbol)     # update symbol cells
            ws_transactions.cell(row=index, column=3, value=date)       # update date cells
            ws_transactions.cell(row=index, column=4, value=share_price).number_format = '$#,##0.00' # update share_price cells
            ws_transactions.cell(row=index, column=5, value=shares)     # update shares cells
            ws_transactions.cell(row=index, column=6, value=amount_paid).number_format = '$#,##0.00' # update amount cells
            
            print(f"ID: {id} SYMBOL: {symbol} DATE: {date} SHARE PRICE: {share_price} SHARES: {shares} AMOUNT PAID: {amount_paid}")
        # save the excel file
        EXCEL_WORK_BOOK.save(WORK_BOOK_PATH)
    except:
        traceback.print_exc() # stacktrace
        input("There was an error updating the transactions. Make sure the spreadsheet isn't already opened. Press Enter to continue...")




def recursive_update():
    update_stock_data()
    print("Updated Transactions! \n")
    time.sleep(10)
    recursive_update() #recursion



def main():
    print("Do you want to run continuously? (Press 1), otherwise, press (Enter)...")
    non_stop_run = input().lower()
    if non_stop_run == "1":
        recursive_update()
    else:
        update_transactions()
        update_dividend_data()
        update_account_value()
        update_stock_data()
        
        print("Do you want to update the symbols you own? (y/n), otherwise, press (Enter)...")
        update_portfolio_symbols = input().lower()
        if update_portfolio_symbols == "Y":
            update_portfolio()
            input("Updated! Press Enter to close...")
        
        input("Updated! Press Enter to close...")

if __name__ == "__main__":
    main()


