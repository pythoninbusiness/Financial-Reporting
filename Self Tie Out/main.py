from bs4 import BeautifulSoup
import jinja2
import helpers

setup = {
    "export_name": "Tied Out Report.html", 
    "file_path": r"C:\Users\RuoyuChen\OneDrive - DigitalBridge\Desktop\Programs\FR Tools\Self Tie Out\DBRG 2022 Q1 10-Q (3.16.22 157 PM).htm"

    # "export_name": "DBRG 2021 Q4 10-K (Current Period).html", 
    # "file_path": r"C:\Users\RuoyuChen\OneDrive - DigitalBridge\Desktop\Programs\FR Tools\Self Tie Out\DBRG 2021 Q4 10-K.htm"
}

env = jinja2.Environment(loader=jinja2.FileSystemLoader("template"))
template = env.get_template("template.html")

file_path = setup["file_path"]
document_name = file_path.split("\\")[-1]
with open(file_path, 'r', encoding='utf-8') as file_:
    html_text = file_.read()

html_text = helpers.tag_numbers(html_text)
soup = BeautifulSoup(html_text, 'lxml')

tie_out_numbers = soup.find_all("span", {"class": "tie-out"})
tables = soup.find_all("table")

support_df = helpers.retrieve_support_dataframe()

for number in tie_out_numbers:
    found, wp_reference = helpers.search_support_wps(number.text, support_df)
    if found:
        number["data-tooltip"] = wp_reference
        number["class"] += ["reference"]
        if len(wp_reference) > 1:
            number["data-tooltip"] = ";".join(wp_reference)
            number["class"] += ["multiple-reference"]
    else:
        number["data-tooltip"] = "not found"
        number["class"] += ["reference", "ref-not-found"]

helpers.remove_tags_from_numbers_in_tables(soup)
# helpers.tag_tables(soup) # HTML uses tables everywhere, so need to refine

output = template.render(html_text=str(soup), document_name=document_name)
with open(setup["export_name"], 'w', encoding="utf-8") as file_:
    file_.write(output)