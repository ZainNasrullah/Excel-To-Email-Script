import openpyxl
import win32com.client as win32

#ask for name of the excel file
#print("What is the name of the excel file that you are using (format: example.xlsx)?)
#name = input()
name = "test.xlsx"

#Load workbook with given name (must be in same folder) and worksheet
wb = openpyxl.load_workbook(name)
ws = wb.active

MaxRow = ws.max_row #calculates the max row
MaxColumn = ws.max_column #calculates the max columns

UserUsage = "" #Create an empty string to save the message body
for rows in range (2, MaxRow+1,1): #iterate across all the rows starting at row 2

    recipient = ws.cell(row=rows,column=1).value + "@hotmail.com"
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
    #mail.send
    UserUsage=""

