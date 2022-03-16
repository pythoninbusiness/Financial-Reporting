import pandas as pd
import win32com.client
import glob
import re, os
import sqlite3
import config


def tag_numbers(html_text):
    html_text = re.sub("&#160;", " ", html_text)

    # paragraph numbers
    html_text = re.sub(r"(\$[\d,\.]+) (million|billion|thousand)*", "<span class='tie-out'>" + r"\g<0>" + "</span>", html_text)

    # paragraph units/shares
    html_text = re.sub(r"([\d,\.]+ shares)", "<span class='tie-out'>" + r"\g<0>" + "</span>", html_text)
    html_text = re.sub(r"([\d,\.]+ [\w+]* shares)", "<span class='tie-out'>" + r"\g<0>" + "</span>", html_text)
    html_text = re.sub(r"([\d,\.]+ units)", "<span class='tie-out'>" + r"\g<0>" + "</span>", html_text)
    html_text = re.sub(r"([\d,\.]+ [\w+]* units)", "<span class='tie-out'>" + r"\g<0>" + "</span>", html_text)

    # paragraph percentages
    html_text = re.sub(r'([\d\.\,]+\%)[^"<]', " <span class='tie-out'>" + r"\g<1>" + "</span>", html_text)
    return html_text


def remove_tags_from_numbers_in_tables(soup):
    tables = soup.find_all("table")
    for table in tables:
        tagged_table_numbers = table.find_all("span", {"class": "tie-out"})
        for number in tagged_table_numbers:
            number["class"].remove("tie-out")
            number["class"].remove("reference")


def tag_tables(soup):
    tables = soup.find_all("table")
    for table in tables:
        table['class'] = table.get('class', []) + ['tie-out', "reference"]


def clean_number_string(num_string):
    percentage = True if "%" in num_string else False
    clean_num = re.findall(r"([\d\.\,]+)", num_string)
    if len(clean_num) == 1:
        clean_num = clean_num[0]
        if percentage:
            clean_num = float(clean_num) / 100
        else:
            clean_num = re.sub(r",", "", clean_num)
            if len(clean_num) > 0 and clean_num != ".":
                clean_num = float(clean_num)

        return clean_num
    else:
        print(f"error with number - {num_string}")
        return False


def retrieve_support_dataframe():
    engine = sqlite3.connect(config.DB_NAME)
    support_dataframe = pd.read_sql(f"SELECT * FROM {config.SUPPORT_TABLE_NAME}", con=engine)
    engine.close()
    return support_dataframe


def search_support_wps(number, support_dataframe):
    clean_number = clean_number_string(number)
    references = support_dataframe.query("Value == @clean_number")["WorkpaperName"].tolist()
    found = True if len(references) > 0 else False
    return found, references


def extract_export_tab(full_workbook_path, excel_com_object, engine_object):
    """function responsible for opening excel file and extracting export tab into a DF"""

    export_tab_name = "0. Export"
    wb = excel_com_object.Workbooks.Open(full_workbook_path, UpdateLinks=0, ReadOnly=1)
    last_saved_date = wb.BuiltinDocumentProperties["Last save time"].value.strftime("%m-%d-%Y %H:%M")

    cursor = engine_object.cursor()
    cursor.execute(f"DROP TABLE {config.SUPPORT_TABLE_NAME}")
    engine_object.commit()

    if export_tab_name in [sheet.Name for sheet in list(wb.Worksheets) if sheet.Visible != 0]:
        export_sheet = wb.Worksheets[export_tab_name]
        export_content = export_sheet.UsedRange()
        column_labels = export_content[0]
        export_data = export_content[1:]
        
        export_df = pd.DataFrame(export_data, columns=column_labels)
        export_df = export_df.dropna(how="all")
        if "WorkpaperName" not in column_labels:
            export_df["WorkpaperName"] = full_workbook_path.split("\\")[-1]
        export_df["LastSaved"] = last_saved_date
        load_data_to_DB(export_df, engine_object)
        wb.Close(SaveChanges=False)
    else:
        print(f"!! {full_workbook_path} does not have an export tab.\n")
        raise Exception("Export tab not found")


def load_data_to_DB(export_dataframe, engine_obj):
    """Store the export dataframe from a workbook into our database"""

    export_dataframe.to_sql(config.SUPPORT_TABLE_NAME, index=False, con=engine_obj, if_exists="append")


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

    engine_object = sqlite3.connect(config.DB_NAME)
    engine_object.execute(f"DROP TABLE {config.SUPPORT_TABLE_NAME}")
    support_directory = config.SUPPORT_DIRECTORY
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
                    extract_export_tab(full_workbook_path, excel_object, engine_object)
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
                        extract_export_tab(full_workbook_path, excel_object, engine_object)
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

