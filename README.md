# server47

Takes data from a specified TD Ameritrade account and places it in an Excel spreadsheet for easier porfolio access and data analysis by using Python and the TDAmeritrade API.

Uses a base spreadsheet 'base.xlsx', inputs data from a specified TDAmeritrade account, and saves it as a spreadsheet called 'portfolio.xlsx' 

### Step 1 - Create environment variable containing your TD account ID
Navigate to Windows Start Menu > Search for environment variables
![image](https://user-images.githubusercontent.com/60449948/153258085-347e2969-28af-49b5-be83-74e671656277.png)


Navigate to Advanced > Environment Variables
![Screenshot 2022-02-09 094717](https://user-images.githubusercontent.com/60449948/153260484-8c516418-f881-406e-a02d-87226475491b.png)



Create a new User Variable
![Screenshot 2022-02-09 095449](https://user-images.githubusercontent.com/60449948/153261239-4de7f63e-75ed-488f-a371-7e393c6d80e0.png)



Variable Name: TD_ACCOUNT 
Variable Value: <your_9-digit-account_id> (E.g. 123456789)
![image](https://user-images.githubusercontent.com/60449948/153259075-15873f65-b243-4c63-bb1d-866456fb5eb2.png)


### Step 2 - Start server47.py and authenticate by logging into your account
On the first execution of the application, you'll be asked to authenticate. A link will be provided to do so in the console.
Once you navigate to the link, you'll be asked to sign-in to TD Ameritrade. 
After signing in, you be directed to a dead page. Copy the entire link of that page and paste it in the application prompt.
You should now be authenticated, td_state.json will be generated containing your access key
