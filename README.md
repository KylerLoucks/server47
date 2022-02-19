# server47
Updates "TD Ameritrade Stonks.xlsx" excel spreadsheet with position data pulled from an API call to TD Ameritrade for a visual portfolio of your account

1.
Navigate to Windows Start Menu > Search for environment variables
![image](https://user-images.githubusercontent.com/60449948/153258085-347e2969-28af-49b5-be83-74e671656277.png)


Navigate to Advanced > Environment Variables
![Screenshot 2022-02-09 094717](https://user-images.githubusercontent.com/60449948/153260484-8c516418-f881-406e-a02d-87226475491b.png)



Create a new User Variable
![Screenshot 2022-02-09 095449](https://user-images.githubusercontent.com/60449948/153261239-4de7f63e-75ed-488f-a371-7e393c6d80e0.png)



Variable Name: TD_ACCOUNT 
Variable Value: <your_account_id> (E.g. 123456789)
![image](https://user-images.githubusercontent.com/60449948/153259075-15873f65-b243-4c63-bb1d-866456fb5eb2.png)


2.
Authenticating:
On the first execution of the application, you'll be asked to authenticate. A link will be provided to do so in the console.
Once you navigate to the link, you'll be asked to sign in to TD Ameritrade. 
After signing in, you be directed to a dead page. Copy the entire link of that page and paste it in the application prompt.
You should now be authenticated, td_state.json will be generated containing your access key
