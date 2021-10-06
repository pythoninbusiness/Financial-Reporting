"""
Purpose: This script will go through a root folder and find all folders inside the root folder. Then the program loops through each 
folder to gather all the excel files and prints out the first sheet into a PDF (stored in our download folder). 

The program finally combines all the PDF's into one pdf with bookmarks.
"""

import win32com.client
import glob
import os, re


root_folder = r"N:\Accounting\CLNS\SEC Reporting\2021 Q3\Support\2020 Q3 Disc Ops Comparative"
download_folder = r".\PDF Folder"
os.chdir(root_folder)

number_of_excel_sheets_to_print = 1

sub_folders = [path_string for path_string in os.listdir() if os.path.isdir(path_string)][:2]

try: 
    excel_object = win32com.client.Dispatch('Excel.Application')
    workbook_counter = 1

    for folder in sub_folders:

        os.chdir(folder)
        print(f"Changed to {folder}\n")

        workbooks = glob.glob("*.xls*")

        for workbook in workbooks:
            full_workbook_path = os.path.join(os.getcwd(), workbook)
            workbook_name_string = re.sub(r"\.xls.", "", workbook) 

            print(f"Opening {workbook}", end="\n")        

            wb = excel_object.Workbooks.Open(full_workbook_path, UpdateLinks=0, ReadOnly=True)
            last_saved_date = wb.BuiltinDocumentProperties["Last save time"].value.strftime("%m-%d-%Y %H:%M")

            # We ignore all sheets that are hidden (i.e., Visible == 0); We can further add logic to only extract names with \d\-
            worksheet_names = [sheet.Name for sheet in list(wb.Worksheets) if sheet.Visible != 0]
            worksheet_names = worksheet_names[:number_of_excel_sheets_to_print]

            # variables to organize our newly printed files
            worksheet_counter_map = "abcdefghijklmnopqrstuvwxyz"
            worksheet_counter = 0

            for worksheet_name in worksheet_names:
                ws = wb.Worksheets[worksheet_name]
                worksheet_reference = worksheet_counter_map[worksheet_counter]
                print_name = f"{workbook_counter}{worksheet_reference}. {workbook_name_string}.pdf"

                print(f"\tPrinting {worksheet_name}".replace(".pdf", ""))
                ws.PageSetup.Zoom = False
                ws.PageSetup.FitToPagesTall = 1
                ws.PageSetup.FitToPagesWide = 1
                ws.PageSetup.LeftHeader = f"&B&12&Kff0000Workbook: {workbook_name_string}, Sheet: {ws.Name}, Last Modified: {last_saved_date}"
                ws.PageSetup.ScaleWithDocHeaderFooter = False
                ws.PrintOut(PrToFileName=os.path.join(download_folder, print_name))

                worksheet_counter += 1

            print(f"Closing {workbook}\n")
            wb.Close(SaveChanges=False)
            workbook_counter += 1

        os.chdir("..")
        print(f"\n\nChanged back to {os.getcwd()}\n")

except Exception as e:
    print(f"\n**Error - {e}")

finally:
    excel_object.Quit()