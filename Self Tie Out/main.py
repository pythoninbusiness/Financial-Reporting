from bs4 import BeautifulSoup
import pandas as pd
import jinja2
import os, re


env = jinja2.Environment(loader=jinja2.FileSystemLoader("template"))
template = env.get_template("template.html")

file_path = r"C:\Users\RuoyuChen\OneDrive - DigitalBridge\Desktop\WIP\Self Tie Out\DBRG 2021 Q4 10-K.html"
with open(file_path, 'r', encoding='utf-8') as file_:
    html_text = file_.read()

html_text = re.sub("&#160;", " ", html_text)
html_text = re.sub(r"(\$[\d,\.]+) (million|billion|thousand)*", "<span class='tie-out'>" + r"\g<0>" + "</span>", html_text)

html_text = re.sub(r"([\d,\.]+ shares)", "<span class='tie-out'>" + r"\g<0>" + "</span>", html_text)
html_text = re.sub(r"([\d,\.]+ [\w+]* shares)", "<span class='tie-out'>" + r"\g<0>" + "</span>", html_text)
html_text = re.sub(r"([\d,\.]+ units)", "<span class='tie-out'>" + r"\g<0>" + "</span>", html_text)

html_text = re.sub(r"([\d\.\,]+\%)", " <span class='tie-out'>" + r"\g<1>" + "</span>", html_text)
soup = BeautifulSoup(html_text, 'lxml')

tie_out_numbers = soup.find_all("span", {"class": "tie-out"})

tie_out_numbers[0]["data-tooltip"] = "test\ntest<br>test"
tie_out_numbers[0]["class"] += ["reference"]
tie_out_numbers[1]["data-tooltip"] = "this is a longer reference"
tie_out_numbers[1]["class"] += ["reference"]


context = dict(
    html_text= str(soup)
)

output = template.render(**context)

with open("Tied Out Report.html", 'w', encoding="utf-8") as file_:
    file_.write(output)