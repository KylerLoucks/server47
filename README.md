# server47
Updates "TD Ameritrade Stonks.xlsx" file with API call to TD Ameritrade developer API account transactions

1.
Create a file named 'secretunstageddata.py' and include the following lines in the file:
TD_ACCOUNT = <your_account_id> # replace <your_account_id> with your 9 digit TD Ameritrade account ID. (E.g. 123456789)


2.
Authenticating:
On the first execution of the application and after every 90 days, you'll be asked to authenticate. A link will be provided to do so in the console.
Once you navigate to the link, you'll be asked to sign in to TD Ameritrade.
After signing in, you be directed to a dead page. Copy the entire link of that page and paste it in the application prompt.
You should now be authenticated.
