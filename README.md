1. Download the project
2. Run 

```
npm install
```

3. Go to Office 365 portal [www.microsoft365.com] and login with admin account
4. Download "Active Users" CSV file > Rename it as Users.csv > Put it on src/raw_data folder
5. Go to Exchange > Download Mailboxes CSV file > Rename it as Exchange.csv > Put it on src/raw_data folder
6. Go to Proofpoint website and copy the table from End Users and Functional Accounts and Paste in the Example file: Proofpoint_example.xlsx

> [!IMPORTANT]
> The colums must be consistant with the information in the head row. Check it carefully

7. Save the Proofpoint_example.xlsx as Proofpoint.xlsx on src/raw_data folder
8. run 

```
npm run start
```