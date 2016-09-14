import openpyxl
import win32com.client as win32

wb = openpyxl.load_workbook("test.xlsx")
print (wb.get_sheet_names())

o = win32.Dispatch("outlook.application")
mail = o.CreateItem(0)
mail.To = "supppyy@hotmail.com"
mail.Subject = "test1"
mail.body = "body of mail"
mail.send

