import pandas as pd
import win32com.client
import glob
import re, os
import sqlite3


class Number():
    def __init__(self):
        self.reference = ""
        self.value = 0
        self.possible_matches = []


    def searchForMatch(self):
        pass


def extractExportTab(full_workbook_path, excel_com_object, engine_object):
    """function responsible for opening excel file and extracting export tab into a DF"""

    export_tab_name = "0. Export"
    wb = excel_com_object.Workbooks.Open(full_workbook_path, UpdateLinks=0, ReadOnly=1)
    last_saved_date = wb.BuiltinDocumentProperties["Last save time"].value.strftime("%m-%d-%Y %H:%M")

    if export_tab_name in [sheet.Name for sheet in list(wb.Worksheets) if sheet.Visible != 0]:
        export_sheet = wb.Worksheets[export_tab_name]
        export_content = export_sheet.UsedRange()
        column_labels = export_content[0]
        export_data = export_content[1:]
        
        export_df = pd.DataFrame(export_data, columns=column_labels)
        export_df = export_df.dropna(how="all")
        export_df["WorkpaperName"] = full_workbook_path
        export_df["LastSaved"] = last_saved_date
        loadDatatoDB(export_df, engine_object)
        wb.Close(SaveChanges=False)
    else:
        print(f"!! {full_workbook_path} has not export tab.\n")
        raise Exception("Export tab not found")


def loadDatatoDB(export_dataframe, engine_obj):
    """Store the export dataframe from a workbook into our database"""

    table_name = "support"
    export_dataframe.to_sql(table_name, index=False, con=engine_obj, if_exists="append")


def compileSearchData():
    """
    Run this to create a database of all the numbers from our support workpapers. 
    Should loop through each workpaper in the support folder, extract all numbers in the "Export" tab and 
    add to a database.
    
    Database Structure:
        WorkpaperName: str
        Name: str
        Value: float
        To: str
        LastUpdated: str
    """

    engine_object = sqlite3.connect("support.db")
    support_directory = r"C:\Users\RuoyuChen\OneDrive - DigitalBridge\Desktop\WIP\Self Tie Out\workpapers"
    excluded_dirs = [
        "x_ss", 
    ]
    exclude_files = [
        "_Foxtrot CLNC Latam NRF Disc Ops Checklist-4Q20 Comparative.xlsx"
    ] 
    os.chdir(support_directory)
    print("changing to support directory...")
    directories = [item for item in os.listdir() if os.path.isdir(item) and item not in excluded_dirs]

    try:
        excel_object = win32com.client.Dispatch('Excel.Application')
        workbook_counter = 1
        directory_counter = 1
        total_directories = len(directories)

        if len(directories) == 0:
            workbooks = glob.glob("*.xls*")
            workbooks = [workbook for workbook in workbooks if workbook not in exclude_files]

            for workbook in workbooks:
                full_workbook_path = os.path.join(os.getcwd(), workbook)
                print(f"Opening {workbook}", end="... ")
                try:
                    extractExportTab(full_workbook_path, excel_object, engine_object)
                    workbook_counter += 1
                except Exception as e:
                    print(e)
                    workbook_counter += 1
                    continue
        else:
            for directory in directories:
                os.chdir(directory)
                print(f"Changed to {directory}\n")

                workbooks = glob.glob("*.xls*")
                workbooks = [workbook for workbook in workbooks if workbook not in exclude_files]

                for workbook in workbooks:
                    full_workbook_path = os.path.join(os.getcwd(), workbook)
                    print(f"Opening {workbook}", end="... ")
                    try:
                        extractExportTab(full_workbook_path, excel_object, engine_object)
                        workbook_counter += 1
                    except Exception as e:
                        print(e)
                        workbook_counter += 1
                        continue

                os.chdir("..")
                print(f"\nReturned to {os.getcwd()} - {directory_counter}/{total_directories}")
                directory_counter += 1 

    except Exception as e:
        print(e)

    finally:
        engine_object.close()
        excel_object.Quit()

