import csv
import openpyxl
import win32com.client
import time
import os
import numpy
import json
from Stacker import stacker

class Excel_compiler:
    def Excel_file_open(file_to_search,sheet_to_search):
        # Opens a session of Excel
        excel = win32com.client.Dispatch("Excel.Application")

        # keeps window closed and ignores errors
        excel.Visible = False
        excel.DisplayAlerts = False

        # open the file that was saved
        workbook_rep_client = excel.Workbooks.Open(os.path.abspath(file_to_search))
        worksheet_rep_client = workbook_rep_client.Worksheets(sheet_to_search)
        print("File opened : " + str(file_to_search))
        return [worksheet_rep_client,worksheet_rep_client]

    def Row_finder(code_to_look_for):
        # Find the value and first row associated to the code
        Value_to_find = worksheet.Columns("C").Find(code_to_look_for, LookAt=1)
        First_row_to_find = Value_to_find.Row
        # While loop top find all the rows that are the same
        Last_row_to_find = First_row_to_find
        while worksheet.Cells(Last_row_to_find, 3).Value == worksheet.Cells(Last_row_to_find + 1, 3).Value:
            Last_row_to_find = Last_row_to_find + 1
        return [First_row_to_find , Last_row_to_find]

    def Column_filler(code_column_to_search,code_column_to_add,
                      code_to_look_for,first_row_to_change, last_row_to_change,
                      column_to_fill,worksheet_in_work):
        """ Open and get Rep_client informations, this function is used to fill the 15th column of the output file """

        # Find the code from the list of all the codes
        Code_to_find = worksheet_in_work.Columns(code_column_to_search).Find(code_to_look_for, LookAt=1)

        # Different strategy if multiple cells or not
        try:
            # int is to make sure the format is right
            Code_to_add = int(worksheet_in_work.Cells(Code_to_find.Row, code_column_to_add).Value)
        # In case the value isn't a number
        except ValueError:
            Code_to_add = worksheet_in_work.Cells(Code_to_find.Row, code_column_to_add).Value

        # Sets the cells to the right value and throws error if the code isnt found
        if Code_to_find == None:
            print("Valeur inconnu")

        else:
            worksheet.Range(str(column_to_fill) + str(first_row_to_change) + ":" + str(column_to_fill) + str(
                last_row_to_change)).Value = Code_to_add

    def Unique_code_lister(self):
        # Loop to create a list of all the codes to find
        list_of_all_codes = []
        for test in range(1, new_last_row):
            if "Code" in str(worksheet.Cells(test, 3).Value):
                continue
            else:
                print("Calcul de tous les codes : "+ str(test)+" / "+str(new_last_row))
                list_of_all_codes.append(int(worksheet.Cells(test, 3).Value))
        unique_list = numpy.unique(list_of_all_codes)
        print(unique_list)

        return unique_list

#-----------------------------------------------------------------------------------------------------------------------
# Opens the json file
f = open('my_chosen_files.json')
my_json_file_data = json.load(f)

# Opens a session of Excel
start = time. time()

#for loop to go through all 4 analcli files
for file_number in range(0,4):

    input_file = my_json_file_data['filename'][file_number]
    print("Ceci est mon input file : "+str(input_file))

    if file_number == 0:
        output_file = 'Bible resultat.xlsx'
        print('Conversion du fichier : ' + str(input_file))
    else :
        output_file = 'fichier_supplementaire'+str(file_number)+'.xlsx'

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

    #-----------------------------------------------------------------------------------------------------------------------
    """ Opens the main file to format the columns and sort accordingly to the prototype """
    # Sort with excel
    excel = win32com.client.Dispatch("Excel.Application")
    # keeps window closed and ignores errors
    #excel.Visible = False
    excel.DisplayAlerts = False

    # open the file that was saved
    workbook = excel.Workbooks.Open(os.path.abspath(output_file))
    print('excel sheet'+str(output_file)+'opened')
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
            print("Suppression des lignes TOTAL : "+str(all_cells) + " / " + str(last_row))

    #Sorts the lines by code
    worksheet.Columns("A:N").Sort(Key1=worksheet.Range('C1'), Order1=1, Orientation=1)

    #Changes cells formula to have to right "Code" format
    column_to_change = 1
    #Condition is for the first character to be a '0'
    while str(worksheet.Cells(column_to_change, 3).Formula)[0] == '0':
        code_formula = worksheet.Cells(column_to_change, 3).Formula
        worksheet.Cells(column_to_change, 3).Formula = '=' + str(code_formula) + '*1'
        print("Code à changer : "+str(column_to_change)+" / "+str(last_row))
        column_to_change += 1

    #Changes to the right value to be consistent after 1000
    worksheet.Range("C1:C"+str(column_to_change)).Copy()
    worksheet.Range("C1:C"+str(column_to_change)).PasteSpecial(Paste=-4163)

    #recounts the last row that changed
    new_last_row = worksheet.UsedRange.Rows.Count

    #Changes first row for each files
    if file_number == 0:
        worksheet.Range("A1:A"+str(new_last_row)).Value= '2020'
    if file_number == 1:
        worksheet.Range("A1:A" + str(new_last_row)).Value = '2021'
    if file_number == 2:
        worksheet.Range("A1:A" + str(new_last_row)).Value = 'Mois 2020'
    if file_number == 3:
        worksheet.Range("A1:A" + str(new_last_row)).Value = 'Mois 2021'

    #Replaces codes for their departement name
    worksheet.Columns("F").Replace('1','HBE', MatchCase=False,SearchFormat=False, ReplaceFormat=False)
    worksheet.Columns("F").Replace('6','LEM', MatchCase=False,SearchFormat=False, ReplaceFormat=False)
    worksheet.Columns("F").Replace('7','ATPA', MatchCase=False,SearchFormat=False, ReplaceFormat=False)
    worksheet.Columns("F").Replace('8','FARDIS', MatchCase=False,SearchFormat=False, ReplaceFormat=False)
    worksheet.Columns("F").Replace('3','VINS ALCOOLS', MatchCase=False,SearchFormat=False, ReplaceFormat=False)
    worksheet.Columns("F").Replace('9','BRADERIE', MatchCase=False,SearchFormat=False, ReplaceFormat=False)

    #Changes columns format to be a number
    worksheet.Columns("J").Replace('.',',', MatchCase=False, SearchFormat=False, ReplaceFormat=False)
    worksheet.Columns("N").Replace('.',',',MatchCase=False,SearchFormat=False, ReplaceFormat=False)
    #worksheet.Columns("J").Numberformat = "Standard"
    worksheet.Columns("J").Style = "Normal"

    #-----------------------------------------------------------------------------------------------------------------------
    #Calls the method to get the list of unique codes
    #Call this method only once
    #Calculates a list of unique codes from the method
    my_unique_list = Excel_compiler.Unique_code_lister('self')
    max_number_code = len(my_unique_list)
    print("La liste à été calculé")

    #-----------------------------------------------------------------------------------------------------------------------
    #Opens excel instance with the wanted REP file from the Open method
    File1=Excel_compiler.Excel_file_open(my_json_file_data['filename'][4], 'Feuil1')
    workbook_rep_client = File1[0]
    worksheet_rep_client = File1[1]
    print("workbook opened")

    #Opens excel instance with the wanted Categorie file from the Open method
    File2=Excel_compiler.Excel_file_open(my_json_file_data['filename'][5], "ulistclrep")
    workbook_categorie = File2[0]
    worksheet_categorie = File2[1]
    print("workbook categorie opened")

    #Opens excel instance with the wanted Classification file from the Open method
    File3=Excel_compiler.Excel_file_open(my_json_file_data['filename'][7], 'Table des marques')
    workbook_marques = File3[0]
    worksheet_marques = File3[1]
    #Delete the departement column
    worksheet_marques.Columns("J").Delete()
    print("workbook classification opened")

    #-----------------------------------------------------------------------------------------------------------------------
    #Loops through all the codes in the output file from REP client
    for iteration_of_code in range(0, max_number_code - 1):
        print("Remplissage des column REP et catégorie client : "+str(iteration_of_code)+" / "+str(max_number_code))
        try:
            #Method to find the range of code cells to set from the output file
            Range_to_change = Excel_compiler.Row_finder(my_unique_list[iteration_of_code])
            #Method to set thoses codes as the right value
            Excel_compiler.Column_filler("A", 2 , my_unique_list[iteration_of_code], Range_to_change[0], Range_to_change[1], "O",worksheet_rep_client)
            Excel_compiler.Column_filler("B", 9 , my_unique_list[iteration_of_code], Range_to_change[0], Range_to_change[1], "P", worksheet_categorie)

        except TypeError:
             print("Pas trouvé")
             continue
    #-----------------------------------------------------------------------------------------------------------------------
    new_last_row = worksheet.UsedRange.Rows.Count
    worksheet.Columns("A:P").Sort(Key1=worksheet.Range('G1'), Order1=1, Orientation=1)

    first_row_produit = 1
    last_row_produit = first_row_produit
    for i in range(1,new_last_row):
        try:
            Code_to_look_for = worksheet.Cells(i,7).Value
            if Code_to_look_for == worksheet.Cells(i+1,7).Value :
                print("Remplissage des colonnes Q à V : " + str(i) + " / " + str(new_last_row))
                last_row_produit += 1
                pass
            else:
                Value_to_find = worksheet_marques.Columns("E").Find(Code_to_look_for, LookAt=1)
                My_row = Value_to_find.Row
                worksheet_marques.Range("H" + str(My_row) + ":M" + str(My_row)).Copy()
                worksheet.Range("Q" + str(first_row_produit) + ":V" + str(last_row_produit)).PasteSpecial(Paste=-4163)
                print("Remplissage des colonnes Q à V : " + str(i) + " / " + str(new_last_row))
                last_row_produit += 1
                first_row_produit = last_row_produit

        except AttributeError:
             pass
             print("Pas trouvé")
             last_row_produit += 1
        except:
            pass
            print("Pas trouvé")
            last_row_produit += 1


    #-----------------------------------------------------------------------------------------------------------------------
    input4 = my_json_file_data['filename'][6]
    # Opens a session of Excel
    excel4 = win32com.client.Dispatch("Excel.Application")
    # keeps window closed and ignores errors
    excel4.Visible = False
    excel4.DisplayAlerts = False
    # open the file that was saved
    workbook_PPN = excel4.Workbooks.Open(os.path.abspath(input4))
    worksheet_PPN = workbook_PPN.Worksheets('ulvalstf')

    #For the last columns, sorting by ref is needed to speed up the process
    worksheet.Columns("A:V").Sort(Key1=worksheet.Range('G2'), Order1=1, Orientation=1)
    last_row_PPN = worksheet_PPN.UsedRange.Rows.Count
    for iteration_PPN in range(2,last_row_PPN):
        try:
            Code_to_look_for_PPN = worksheet_PPN.Cells(iteration_PPN, 2).Value
            Value_to_find_PPN = worksheet.Columns("G").Find(Code_to_look_for_PPN, LookAt=1)
            My_row_PPN = Value_to_find_PPN.Row
            if worksheet.Cells(My_row_PPN,22).Value == worksheet_PPN.Cells(iteration_PPN,12).Value:
                print('rien à changer -- ligne : ' + str(iteration_PPN) + ' / '+str(last_row_PPN))
                pass
            else:
                print('A modifier -- ligne : '+ str(iteration_PPN) + ' / '+str(last_row_PPN))
                My_last_row_PPN = My_row_PPN
                while worksheet.Cells(My_last_row_PPN,7).Value==worksheet.Cells(My_last_row_PPN + 1,7).Value :
                    My_last_row_PPN = My_last_row_PPN + 1
                worksheet_PPN.Cells(iteration_PPN,12).Copy()
                worksheet.Range("V" + str(My_row_PPN) + ":V" + str(My_last_row_PPN)).PasteSpecial(Paste=-4163)
                My_row_PPN = My_last_row_PPN

        except AttributeError:
              continue
              print('Pas trouvé')

    #Sorts the lines and deletes the files with TOTAL in them
    last_row = worksheet.UsedRange.Rows.Count

    for row_destructor_iterator in reversed(range(1,last_row)):
        if "DEPART" in worksheet.Cells(row_destructor_iterator,6).Value:
            worksheet.Rows(row_destructor_iterator).Delete()
            print("Ligne : "+str(row_destructor_iterator) + " est détruit.")
        else:
            break

    #Create new first row
    worksheet.Range("A1:P1").EntireRow.Insert()
    worksheet.Cells(1, 1).Value = 'Années'
    worksheet.Cells(1, 2).Value = 'REPRES'
    worksheet.Cells(1, 3).Value = 'Code'
    worksheet.Cells(1, 4).Value = 'Nom du client'
    worksheet.Cells(1, 5).Value = 'Catég. client'
    worksheet.Cells(1, 6).Value = 'DEPART'
    worksheet.Cells(1, 7).Value = 'REFERENCE'
    worksheet.Cells(1, 8).Value = 'DESIGNATION'
    worksheet.Cells(1, 9).Value = 'REGROUP'
    worksheet.Cells(1, 10).Value = 'QTE'
    worksheet.Cells(1, 11).Value = 'C.A. NET '
    worksheet.Cells(1, 12).Value = 'Remise'
    worksheet.Cells(1, 13).Value = 'MARGE'
    worksheet.Cells(1, 14).Value = '%MRG'
    worksheet.Cells(1, 15).Value = 'REP 2021 CLIENT'
    worksheet.Cells(1, 16).Value = 'Groupe de clients'
    worksheet.Cells(1, 17).Value = 'Marque'
    worksheet.Cells(1, 18).Value = 'Unité de besoin'
    worksheet.Cells(1, 19).Value = 'Fournisseur'
    worksheet.Cells(1, 20).Value = 'Sous-catégorie'
    worksheet.Cells(1, 21).Value = 'Catégorie'
    worksheet.Cells(1, 22).Value = 'Tarification'
    #-----------------------------------------------------------------------------------------------------------------------
    #Autofits column sizes
    worksheet.Columns.AutoFit()

    #Saves finished workbook
    workbook.Save()
    print('fichier sauvegardé')
    excel.Application.Quit()
    stop = time.time()
    print('Temps de complétion : '+str(int((stop-start)/60))+ ' minutes')

stacker.file_stacker('self')

final_text ="""
███████╗██╗███╗░░██╗  
██╔════╝██║████╗░██║  
█████╗░░██║██╔██╗██║  
██╔══╝░░██║██║╚████║  
██║░░░░░██║██║░╚███║  
╚═╝░░░░░╚═╝╚═╝░░╚══╝  

██████╗░██╗░░░██╗  
██╔══██╗██║░░░██║  
██║░░██║██║░░░██║  
██║░░██║██║░░░██║  
██████╔╝╚██████╔╝  
╚═════╝░░╚═════╝░  

████████╗██████╗░░█████╗░██╗████████╗███████╗███╗░░░███╗███████╗███╗░░██╗████████╗
╚══██╔══╝██╔══██╗██╔══██╗██║╚══██╔══╝██╔════╝████╗░████║██╔════╝████╗░██║╚══██╔══╝
░░░██║░░░██████╔╝███████║██║░░░██║░░░█████╗░░██╔████╔██║█████╗░░██╔██╗██║░░░██║░░░
░░░██║░░░██╔══██╗██╔══██║██║░░░██║░░░██╔══╝░░██║╚██╔╝██║██╔══╝░░██║╚████║░░░██║░░░
░░░██║░░░██║░░██║██║░░██║██║░░░██║░░░███████╗██║░╚═╝░██║███████╗██║░╚███║░░░██║░░░
░░░╚═╝░░░╚═╝░░╚═╝╚═╝░░╚═╝╚═╝░░░╚═╝░░░╚══════╝╚═╝░░░░░╚═╝╚══════╝╚═╝░░╚══╝░░░╚═╝░░░"""

print(final_text)

