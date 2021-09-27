from tkinter.filedialog import askopenfilename
import win32com.client
from pathlib import Path
import openpyxl
import time
import csv
import os

file_list = ['0','1','2','3','4','5','6','7']  # list needs to be populated first

class transformer():

    #Method to chose which files to transform
    def setFile(labeltochange, file_number):
        file_path = askopenfilename()  # show an "Open" dialog box and return the path to the selected file
        path = Path(file_path)
        labeltochange['text'] = path.name
        file_list[file_number] = file_path
        return (file_path, file_number)

    #Method to convert .xlsx files as stated in the documentation
    def convert_files(self):
        #Time testing start
        start = time.time()

        for i in (0,1,2,3):
            input_file = file_list[i]
            output_file = 'Fichier analcli/fichier' + str(i) + '.xlsx'
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
                wb.save(output_file)
                print(str(output_file) + ' : file is saved ')

                # Sort with excel
                excel = win32com.client.Dispatch("Excel.Application")
                # keeps window closed and ignores errors
                excel.Visible = False
                excel.DisplayAlerts = False

                # open the file that was saved
                workbook = excel.Workbooks.Open(os.path.abspath('Fichier analcli/fichier' + str(i) + '.xlsx'))
                print('excel sheet opened')
                worksheet = workbook.Worksheets('Sheet')

                #Sorts the lines and deletes the files with TOTAL in them
                annulation = 0
                print(annulation)
                last_row = worksheet.UsedRange.Rows.Count

                worksheet.Range('G3:G93154').Sort(Key1=worksheet.Range('G1'), Order1=1, Orientation=1)

                for all_cells in range(1 ,last_row):
                    if worksheet.Range(worksheet.Cells(all_cells, 7),worksheet.Cells(all_cells, 7)).Value == 'TOTAL DEPARTEMENT':
                        annulation += 1
                    else:
                        print(all_cells)
                print('nombre d annulation : '+ str(annulation))


                workbook.Save()
                print('fichier sauvegardé')

                excel.Application.Quit()

            # Time testing
            end = time.time()
            print('Elapsed time : '+ str(end - start)+ ' Secondes')









