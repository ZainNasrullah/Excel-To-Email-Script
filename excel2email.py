import openpyxl
import win32com.client as win32


emailmode = ""
while (emailmode.upper() != "YES" and emailmode.upper() != "NO"):
    print("Send e-mail mode? (answer yes or no)")
    emailmode = input()

#ask for name of the excel file
print("What is the name of the excel file that you are using (format: example.xlsx)?)
name = input()
#name = "test.xlsx"


print("This program will append '@ge.com' to the SSO column and send an e-mail to that address.")
emailappend = "@hotmail.com" #value to append is set here

#Load workbook with given name (must be in same folder) and worksheet
wb = openpyxl.load_workbook(name)
ws = wb.active

MaxRow = ws.max_row #calculates the max row
MaxColumn = ws.max_column #calculates the max columns

UserUsage = "" #Create an empty string to save the message body
for rows in range (2, MaxRow+1,1): #iterate across all the rows starting at row 2

    recipient = ws.cell(row=rows,column=1).value + emailappend
    for columns in range(1,MaxColumn+1,1): #iterates through every column
        UserUsage+=(str(ws.cell(row=1,column=columns).value) + ": " + str(ws.cell(row=rows,column=columns).value) + "\n")

    print (UserUsage)
    print(recipient)
    print("")

    o = win32.Dispatch("outlook.application")
    mail = o.CreateItem(0)
    mail.To = recipient
    mail.Subject = "Notification"
    mail.body = str(UserUsage)
    if (emailmode.upper() == "YES"):
        mail.send
    UserUsage=""

