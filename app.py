# step 1 : UPLOADS 폴더 만들기
# step 2 : HTML TEMPLATE 만들어서 url 만들기
# step 3 : FORM 에서 POST 이벤트 발생 시 함수 실행
# step 4 : request.file로 파일을 지정한 정로에 저장하기
# step 5 : openpyxl로 해당 폴더를 가지고와서 값 변환하기
# step 6 : 변환된 값을 table_html 글로벌 전역 변수로 선언하기
# step 7 : 해당 값을 가진 html을 폴더에 저장하기
# step 8 : 다운로드 / 을 통한 요청이 왔을 시 해당 파일을 보내주기


# 업로드 폴더 만들기
from flask import Flask, request, send_file, render_template_string
import openpyxl
import os

app = Flask(__name__)

UPLOAD_FOLDER = "uploads"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

# HTML 폼 템플릿 (보여질 화면)
HTML_TEMPLATE = '''
<!DOCTYPE html>
<html>
    <head>
        <title>Excel Viewer</title>
    </head>
    <body>
        <h2>Upload an Excel File</h2>
        <form action="/" method="post" enctype="multipart/form-data">
            <input type="file" name="file">
            <input type="submit" value="Upload">
        </form>
        <a href="/download" class="download-button">다운로드 버튼</a>
    </body>
</html>
'''

# 값을 넣을 출력값 변수 설정
table_html = ""

def workbook(file_path):
    wb = openpyxl.load_workbook(file_path)
    sheet = wb.active
    table = ""
    
    for row in sheet.iter_rows():
        table += "<tr>"
        for cell in row:
            table += f"<td>{cell.value or ''}</td>"
        table += "</tr>"
    return table

@app.route("/", methods=["GET", "POST"])
def upload_file():
    global table_html
    
    if request.method == "POST":
        if "file" not in request.files:
            return "No file uploaded"
        
        file = request.files["file"]
        file_path = os.path.join(UPLOAD_FOLDER, file.filename)
             
        file.save(file_path) # 해당 경로 파일 저장
        
        table_html = workbook(file_path) # 값을 가지고 와서 함수 실행    
    
    return render_template_string(HTML_TEMPLATE, table = table_html)

@app.route("/download")
def download():
    html_file_path = os.path.join(UPLOAD_FOLDER, "table.html")
    
    downHtml = f'''
        <!DOCTYPE html>
        <html>
            <head>
                <title>Excel</title>
            </head>
            <body>
                <table>
                    {table_html}
                </table>
            </body>
        </html>
    '''
    
    
    with open(html_file_path, "w", encoding="utf-8") as f:
        f.write(downHtml)
        
    return send_file(html_file_path, as_attachment = True)

if __name__ == "__main__":
    app.run(debug=True, port=5000)



# 내용 정리
# 생성한 html 태그에서 /을 이용해 Get, Post 요청을 보낼 수 있다.
# 해당 요소를 받는 함수에서 해당 값을 읽고 실행시킨다.




