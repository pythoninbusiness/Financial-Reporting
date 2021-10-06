"""
# Purpose: To create a python script that will later be run by Windows Task Scheduler (or cron) to monitor a root folder for any changes in the files within. 
The reason being changes made to support files may not be reflected in the live document (i.e., Workiva Workdesk). 
The script will send an email with a list of all the files that have been modified since the last time the program was run. 
"""

from collections import namedtuple
import win32com.client as win32
import datetime as dt
import pandas as pd
import sqlite3
import os


# SET UP
folder_where_data_will_be_stored = r"C:\Users\John\Check for Changes"
root_directory = r"C:\Users\John\path we want to monitor"

os.chdir(folder_where_data_will_be_stored)

todays_date = dt.datetime.today().strftime('%#m-%d-%y')

engine = sqlite3.connect("last_modified.db")
previously_modified_report_table_name = "previously_modified"
grandfather_table_name = "previously_modified_df"

file_and_modified_date = namedtuple("file_and_modified_date", "filename lastmodified")
list_of_tuples = list()


# LOOP THROUGH FILES IN ROOT DIRECTORY
for root_path, directories, files in os.walk(root_directory):
    for name in files:
        full_path_to_file = os.path.join(root_path, name)
        last_modified_date = dt.datetime.fromtimestamp( os.path.getmtime(full_path_to_file) ).strftime(r"%m/%d/%Y %H:%M") 

        list_of_tuples.append(
            file_and_modified_date(
                filename=name,
                lastmodified=last_modified_date
            )
        )

last_modified_dataframe = pd.DataFrame(list_of_tuples)


# LOAD AND READ FROM DATABASE
previously_noted_modified_date_df = pd.read_sql(f"Select * from {previously_modified_report_table_name}", engine)

# Move the last result to the grandfather table in case we want to analyze the previous results
previously_noted_modified_date_df.to_sql(grandfather_table_name, con=engine, index=False, if_exists="replace")
last_modified_dataframe.to_sql(previously_modified_report_table_name, con=engine, index=False, if_exists="replace")
engine.close()


# COMBINE
merged = last_modified_dataframe.merge(previously_noted_modified_date_df, on="filename", how="left")
files_recently_updated = merged.query(" `lastmodified_x` != `lastmodified_y` ")
files_recently_updated.columns = ["File", "Last Modified", "Previously Modified"]


# EMAIL TO SELF
outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = "youremail@yourdomain.com"
mail.Subject = f"Last Modified Report - {todays_date}"

mail.HtmlBody = (f"""
<style>
table {{border:none; border-collapse:collapse;}}
thead tr {{text-align: left !important;}}
tr {{text-align: left !important; text-indent: 10px;}}
th, td {{padding: 5px;border-style:none;}}
tr td:nth-of-type(2) {{color: #4e8762 !important;}}
tr td:nth-of-type(3) {{color: #bd7272;}}
thead {{border-bottom: 2px solid black;}}
</style>

Files that have been recently modified:
<br><br>
{files_recently_updated.to_html(index=False)}
<br><br>
""")

# mail.Send()
mail.Display()
