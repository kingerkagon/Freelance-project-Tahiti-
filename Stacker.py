import win32com.client
import os
import time

start = time.time()
class stacker:
    def file_stacker(self):

        #---------------------------------------------------------------------------------------------------------------
        input1 = 'Bible résultat.xlsx'
        # Opens a session of Excel
        excel_2020 = win32com.client.Dispatch("Excel.Application")
        # keeps window closed and ignores errors
        excel_2020.Visible = True
        excel_2020.DisplayAlerts = False
        # open the file that was saved
        workbook_2020 = excel_2020.Workbooks.Open(os.path.abspath(input1))
        worksheet_2020 = workbook_2020.Worksheets('Sheet')
        #---------------------------------------------------------------------------------------------------------------
        input2 = 'fichier_supplementaire1.xlsx'
        # Opens a session of Excel
        excel_2021 = win32com.client.Dispatch("Excel.Application")
        # keeps window closed and ignores errors
        excel_2021.Visible = True
        excel_2021.DisplayAlerts = False
        # open the file that was saved
        workbook_2021 = excel_2021.Workbooks.Open(os.path.abspath(input2))
        worksheet_2021 = workbook_2021.Worksheets('Sheet')
        #---------------------------------------------------------------------------------------------------------------
        input3 = 'fichier_supplementaire2.xlsx'
        # Opens a session of Excel
        excel_2020_mois = win32com.client.Dispatch("Excel.Application")
        # keeps window closed and ignores errors
        excel_2020_mois.Visible = True
        excel_2020_mois.DisplayAlerts = False
        # open the file that was saved
        workbook_2020_mois = excel_2020_mois.Workbooks.Open(os.path.abspath(input3))
        worksheet_2020_mois = workbook_2020_mois.Worksheets('Sheet')
        #---------------------------------------------------------------------------------------------------------------
        input4 = 'fichier_supplementaire3.xlsx'
        # Opens a session of Excel
        excel_2021_mois = win32com.client.Dispatch("Excel.Application")
        # keeps window closed and ignores errors
        excel_2021_mois.Visible = True
        excel_2021_mois.DisplayAlerts = False
        # open the file that was saved
        workbook_2021_mois = excel_2021_mois.Workbooks.Open(os.path.abspath(input4))
        worksheet_2021_mois = workbook_2021_mois.Worksheets('Sheet')
        #---------------------------------------------------------------------------------------------------------------
        last_row_2020 = worksheet_2020.UsedRange.Rows.Count
        last_row_2021 = worksheet_2021.UsedRange.Rows.Count
        last_row_2020_mois = worksheet_2020_mois.UsedRange.Rows.Count
        last_row_2021_mois = worksheet_2021_mois.UsedRange.Rows.Count

        #Paste from first file
        worksheet_2021.Range("A2:V" + str(last_row_2021)).Copy()
        worksheet_2020.Range("A" + str(last_row_2020) + ":V" + str(last_row_2020)).PasteSpecial(Paste=-4163)

        #Paste from second file, number of rows are added
        worksheet_2020_mois.Range("A2:V" + str(last_row_2020_mois)).Copy()
        worksheet_2020.Range("A" + str(last_row_2020+last_row_2021) + ":V" + str(last_row_2020+last_row_2021)).PasteSpecial(Paste=-4163)

        #Paste from second file, number of rows are added
        worksheet_2021_mois.Range("A2:V" + str(last_row_2021_mois)).Copy()
        worksheet_2020.Range("A" + str(last_row_2020+last_row_2021+last_row_2020_mois) + ":V" + str(last_row_2020+last_row_2021+last_row_2020_mois)).PasteSpecial(Paste=-4163)

        workbook_2020.Save()
        excel_2020.Application.Quit()
        excel_2021.Application.Quit()
        excel_2020_mois.Application.Quit()
        excel_2021_mois.Application.Quit()

        stop = time.time()

        print(stop - start)

#stacker.file_stacker('self')