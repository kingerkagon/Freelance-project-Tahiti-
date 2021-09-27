import csv
import openpyxl
import win32com.client
import time
import os

start = time. time()
input_file = 'Fichier analcli/analcliannuel2020_convertit.txt'

output_file = 'Fichier analcli/fichier_test.xlsx'
print('Conversion du fichier' + str(input_file))

#Creates the files and workbook
num_row=0 #to skip first rows
wb = openpyxl.Workbook()
ws = wb.worksheets[0]
with open(input_file, encoding='utf8', errors='ignore') as data:
    reader = csv.reader(data, delimiter=';')
    for row in reader:
        if num_row<3:
            pass # skips first rows
        else:
            ws.append(row) #adds the rows from .txt file into excel worksheet
        num_row += 1
    print('fichier ouvert')

    # .xlsx file saved
    ws.insert_cols(1)
    wb.save(output_file)
    print(str(output_file) + ' : file is saved ')

    # Sort with excel
    excel = win32com.client.Dispatch("Excel.Application")
    # keeps window closed and ignores errors
    excel.Visible = False
    excel.DisplayAlerts = False

    # open the file that was saved
    workbook = excel.Workbooks.Open(os.path.abspath('Fichier analcli/fichier_test.xlsx'))
    print('excel sheet opened')
    worksheet = workbook.Worksheets('Sheet')

    #Sorts the lines and deletes the files with TOTAL in them
    last_row = worksheet.UsedRange.Rows.Count
    worksheet.Columns("A:N").Sort(Key1=worksheet.Range('H1'), Order1=1, Orientation=1)

    #Loops through all used cells to find TOTAL cells and delete them
    for all_cells in reversed(range(1,last_row)):

        #Checks for 'TOTAL DEPARTEMENT' rows
        if worksheet.Cells(all_cells, 8).Value == None:
            worksheet.Cells(all_cells, 8).EntireRow.Delete()
            if worksheet.Cells(all_cells-1, 8).Value != None:
                worksheet.Cells(all_cells, 8).EntireRow.Delete()
                print('supprimé lignes '+str(all_cells))
        #Checks for other TOTAL rows
        if worksheet.Cells(all_cells, 8).Value == 'TOTAL DEPARTEMENT':
            worksheet.Cells(all_cells, 8).EntireRow.Delete()
            if worksheet.Cells(all_cells-1, 8).Value != 'TOTAL DEPARTEMENT':
                worksheet.Cells(all_cells, 8).EntireRow.Delete()
                print('supprimé lignes '+str(all_cells))
                break
        else:
            print(all_cells)

    #recounts the last row that changed
    new_last_row = worksheet.UsedRange.Rows.Count
    #Fills first row with 2020
    worksheet.Range("A1:A"+str(new_last_row)).value= '2020'
    #Create new first row
    worksheet.Range("A1:A1").EntireRow.Insert()
    worksheet.cells(1, 1).value = 'Années'
    worksheet.cells(1, 2).value = 'REPRES'
    worksheet.cells(1, 3).value = 'Code'
    worksheet.cells(1, 4).value = 'Nom du client'
    worksheet.cells(1, 5).value = 'Catég. client'
    worksheet.cells(1, 6).value = 'DEPART'
    worksheet.cells(1, 7).value = 'REFERENCE'
    worksheet.cells(1, 8).value = 'DESIGNATION'
    worksheet.cells(1, 9).value = 'REGROUP'
    worksheet.cells(1, 10).value = 'QTE'
    worksheet.cells(1, 11).value = 'C.A. NET '
    worksheet.cells(1, 12).value = 'Remise'
    worksheet.cells(1, 13).value = 'MARGE'
    worksheet.cells(1, 14).value = '%MRG'
    worksheet.cells(1, 15).value = 'REP 2021 CLIENT'
    worksheet.cells(1, 16).value = 'Groupe de clients'
    worksheet.cells(1, 17).value = 'Marque'
    worksheet.cells(1, 18).value = 'Unité de besoin'
    worksheet.cells(1, 19).value = 'Fournisseur'
    worksheet.cells(1, 20).value = 'Catégorie'
    worksheet.cells(1, 21).value = 'Sous-catégorie'
    worksheet.cells(1, 22).value = 'Tarification'
    #Autofits column sizes
    worksheet.Columns.AutoFit()
    #Saves finished workbook
    workbook.Save()
    print('fichier sauvegardé')
    excel.Application.Quit()
    stop = time.time()
    print('Temps de complétion : '+str(stop-start)+ ' secondes')

