# Load excel 
# 추후 기능 추가로는 해당 요소에 실제 input file이 들어올것임
wb = openpyxl.load_workbook('test.xlsx')
sheet = wb.active

# initalize empty table string
table = ""

# Iterate through the rows and cells of the sheet
for row in sheet.iter_rows():
    table += "<tr>"
    for cell in row:
        table += f"<td>{cell.value or ''}</td>"
    table += "</tr>"
    
# Create the full HTML page
html = f'''
<!DOCTYPE html>
<html>
    <head>
        <title>Excel</title>
    </head>
    <body>
        <table>
            {table}
        </table>
    </body>
</html>
'''

# Save the HTML to a file
with open("test.html", "w", encoding="utf-8") as f:
    f.write(html)
    
print("file saved as html")