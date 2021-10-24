import win32com.client
import os
import time

start = time.time()

#-----------------------------------------------------------------------------------------------------------------------
input2 = 'fichier_test.xlsx'
# Opens a session of Excel
excel2 = win32com.client.Dispatch("Excel.Application")
# keeps window closed and ignores errors
excel2.Visible = True
excel2.DisplayAlerts = False
# open the file that was saved
workbook_test = excel2.Workbooks.Open(os.path.abspath(input2))
worksheet_test = workbook_test.Worksheets('Sheet')
#-----------------------------------------------------------------------------------------------------------------------
input3 = 'Fichier PPN LIBRE PG.xlsx'
# Opens a session of Excel
excel3 = win32com.client.Dispatch("Excel.Application")
# keeps window closed and ignores errors
excel3.Visible = False
excel3.DisplayAlerts = False
# open the file that was saved
workbook_PPN = excel3.Workbooks.Open(os.path.abspath(input3))
worksheet_PPN = workbook_PPN.Worksheets('ulvalstf')
#----------------------------------------------------------------------------------------------------------------------
worksheet_test.Columns("A:V").Sort(Key1=worksheet_test.Range('G1'), Order1=1, Orientation=1)
last_row_PPN = worksheet_PPN.UsedRange.Rows.Count
print(last_row_PPN)
for iteration_PPN in range(2,last_row_PPN):
    try:
        Code_to_look_for_PPN = worksheet_PPN.Cells(iteration_PPN, 2).Value
        Value_to_find_PPN = worksheet_test.Columns("G").Find(Code_to_look_for_PPN, LookAt=1)
        My_row_PPN = Value_to_find_PPN.Row
        if worksheet_test.Cells(My_row_PPN,22).Value == worksheet_PPN.Cells(iteration_PPN,12).Value:
            print('rien à changer -- ligne : ' + str(iteration_PPN) + ' / '+str(last_row_PPN))
            pass
        else:
            print('A modifier -- ligne : '+ str(iteration_PPN) + ' / '+str(last_row_PPN))
            print(str(worksheet_test.Cells(My_row_PPN,22).Value) + " =/= " + str(worksheet_PPN.Cells(iteration_PPN,12).Value))
            print('MyRow est : '+str(My_row_PPN))
            My_last_row_PPN = My_row_PPN
            while worksheet_test.Cells(My_last_row_PPN,7).Value==worksheet_test.Cells(My_last_row_PPN + 1,7).Value :
                My_last_row_PPN = My_last_row_PPN + 1
            print("Ma valeur à copier : "+str(worksheet_PPN.Cells(iteration_PPN,12)))
            worksheet_PPN.Cells(iteration_PPN,12).Copy()
            print("copié : "+str(worksheet_PPN.Cells(iteration_PPN,12).Value))
            worksheet_test.Range("V" + str(My_row_PPN) + ":V" + str(My_last_row_PPN)).PasteSpecial(Paste=-4163)
            print("collé à : "+ str("V"+str(My_row_PPN)+":V"+str(My_last_row_PPN)))
            print("Valeur : "+ str(worksheet_test.Cells(My_last_row_PPN,7).Value)+" Début : "+str(My_row_PPN)+ " // Fin : " + str(My_last_row_PPN))
            My_row_PPN = My_last_row_PPN

    except AttributeError:
          continue
          print('Pas trouvé')


workbook_test.Save()
stop = time.time()
excel2.Application.Quit()
excel3.Application.Quit()

print(stop - start)