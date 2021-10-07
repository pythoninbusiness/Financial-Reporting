"""
Purpose: To take a an HTML financial document (i.e., 10-Q) and tie out all numbers found in paragraphs. For example, $8 million or 5,676,392 shares. 
The script uses regular expressions, so depending on how the Company writes their disclosures and reports, one may need to update the regular expression to catch all. 

The output file is an HTML file which can then be printed into a PDF and further tied out.
"""


import re

file_path = r".\Example 10-Q.htm"

with open(file_path, 'r') as file_:
    html_text = file_.read()

html_text = re.sub("&#160;", " ", html_text)
html_text = re.sub(r"(\$[\d,\.]*) (million|billion|thousand)*", "<span class='tie-out'>" + r"\g<0>" + "</span>", html_text)
html_text = re.sub(r"([\d,\.]+ shares)", "<span class='tie-out'>" + r"\g<0>" + "</span>", html_text)
html_text = re.sub(r"\s([\d\.\,]+\%)", " <span class='tie-out'>" + r"\g<1>" + "</span>", html_text)

output = f"""
<html>
    <head>
        <style>
        .tie-out {{
        border: 2px solid red;
        padding: 1px;
        }}
	.hide-tie-out {{
            border: none !important;
            position: unset;
        }}
        </style>
    </head>

    <body>
    {html_text}

     <script>
        const tieOutValues = document.querySelectorAll('.tie-out');

        tieOutValues.forEach(el => el.addEventListener('click', event => {{
            el.classList.toggle("hide-tie-out");
        }}));
	
	function tieOutSelection() {{
            var newSpan = document.createElement('span');
            newSpan.setAttribute('class', 'tie-out');

            selection_object = window.getSelection()
            selection_text = selection_object.toString()
            newSpan.textContent = selection_text;

            var range = selection_object.getRangeAt(0);
            range.deleteContents();
            range.insertNode(newSpan);
        }}

        document.addEventListener('keydown', function (event) {{
            if (event.key === 't') {{
                tieOutSelection();
            }}
        }})
    </script>
    </body>
</html>
"""

with open("Tied Out Report.html", 'w', encoding="utf-8") as file_:
    file_.write(output)
