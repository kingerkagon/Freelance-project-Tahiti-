import win32com.client
import os

# Sort with excel
excel = win32com.client.Dispatch("Excel.Application")

# keeps window closed and ignores errors
excel.Visible = False
excel.DisplayAlerts = False

# open the file that was saved
workbook = excel.Workbooks.Open(os.path.abspath('Fichier analcli/fichier_test.xlsx'))
print('excel sheet opened')
worksheet = workbook.Worksheets('Sheet')
new_last_row = worksheet.UsedRange.Rows.Count


workbook.Save()
excel.Application.Quit()


