import openpyxl
import win32com.client as win32
import sys

emailmode = ""

#ask for name of the excel file
print("What is the name of the excel file that you are using (format: example.xlsx)?")
name = input()
#name = "test.xlsx"

try:
    wb = openpyxl.load_workbook(name)
    ws = wb.active
except:
        print("Could not load your excel workbook. Enter any key to exit.")
        start = input()
        sys.exit(130)

print("Which row is the header information (Division, phone#, user, employee id, etc.) in?")
HeaderRow = int(input())

print("What is the first row in which there is user data?")
FirstRow = int(input())

print("Which column(enter an integer value ex. A is 1, B is 2, etc.) is the SSO in?")
SSO = int(input())

print("This program will append '@ge.com' to the SSO column and send an e-mail to that address.")
emailappend = "@hotmail.com" #value to append is set here

print("We have detected " + str(MaxRow) + " rows and " + str(MaxColumn) + " columns in total.")

print("To start, type anything and hit enter")
start = input()

#Load workbook with given name (must be in same folder) and worksheet


MaxRow = ws.max_row #calculates the max row
MaxColumn = ws.max_column #calculates the max columns

UserUsage = "" #Create an empty string to save the message body
for rows in range (FirstRow, MaxRow+1,1): #iterate across all the rows starting at row 2

    recipient = ws.cell(row=rows,column=SSO).value + emailappend #appends the e-mail address

    for columns in range(1,MaxColumn+1,1): #iterates through every column
        UserUsage+=(str(ws.cell(row=HeaderRow,column=columns).value) + ": " + str(ws.cell(row=rows,column=columns).value) + "\n")

    # print(recipient)
    #print (UserUsage)
    #print("")

    outlook = win32.Dispatch("outlook.application")
    mail = outlook.CreateItem(0)
    mail.To = recipient
    mail.Subject = "Notification"
    mail.body = str(UserUsage)
    if (rows == FirstRow):
        print("The first e-mail will look like the following:\n" + str(UserUsage))
        print("Are you sure you want to send (type yes or it will quit)? This message will not repeat for all subsequent entries.")
        emailmode = input()
        if (emailmode.upper() != "YES"):
            sys.exit(130)
    mail.send
    UserUsage=""

print("Complete. Enter anything to exit.")
start = input()
sys.exit(130)


