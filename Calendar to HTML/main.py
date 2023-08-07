import pandas as pd
import datetime as dt
import os, re


file_name = "Calendar to list.xlsx"
date_format = "%A, %B %d"

df = pd.read_excel(file_name)
date_columns = df.columns[1:]
html_outline = "<div><ul>"

for date in date_columns:
    
    html_outline += f"<li><h1>{date.strftime(date_format)}</h1>\n<ul>\n"
    
    for task in df[date].dropna():
        subtasks = re.split(r"\n+", task)
        for subtask in subtasks:
            html_outline += f"<li>{subtask}</li>\n"
            
    html_outline += "</ul></li>\n"
html_outline += "</ul></div>"

with open("tasks_outline.html", "w") as f:
    f.write(html_outline)

print("HTML outline generated successfully!")