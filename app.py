from flask import Flask, request, render_template, send_file
import PyPDF2
import pandas as pd
from io import StringIO, BytesIO

app = Flask(__name__)

def convert_pdf_to_excel(pdf_file):
    reader = PyPDF2.PdfReader(pdf_file)
    num_pages = len(reader.pages)
    df_list = []

    for page_num in range(num_pages):
        page = reader.pages[page_num]
        text = page.extract_text()
        
        # พิมพ์ข้อความที่ดึงมาจากแต่ละหน้าเพื่อตรวจสอบ
        print(f"Text from page {page_num + 1}:\n{text}\n{'-'*40}")

        if text.strip():
            try:
                # แยกข้อความตามบรรทัด
                lines = text.split('\n')
                rows = []
                for line in lines:
                    # แยกแต่ละบรรทัดตามช่องว่าง (space)
                    row = line.split()
                    rows.append(row)
                
                # สร้าง DataFrame จากรายการที่จัดเรียง
                data = pd.DataFrame(rows)
                print(f"DataFrame from page {page_num + 1}:\n{data}\n{'-'*40}")
                df_list.append(data)
            except Exception as e:
                print(f"Error processing page {page_num}: {e}")

    # ตรวจสอบว่า df_list มี DataFrame หรือไม่
    if df_list:
        df = pd.concat(df_list, ignore_index=True)
        print(f"Combined DataFrame:\n{df}\n{'-'*40}")
    else:
        df = pd.DataFrame()

    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False)
    output.seek(0)
    return output

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return 'No file part', 400
    file = request.files['file']
    if file.filename == '':
        return 'No selected file', 400
    if file and file.filename.endswith('.pdf'):
        try:
            excel_file = convert_pdf_to_excel(file)
            return send_file(excel_file, as_attachment=True, download_name='output.xlsx', mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
        except Exception as e:
            return f'Error processing file: {e}', 500
    else:
        return 'Invalid file format', 400

if __name__ == '__main__':
    app.run(debug=True)
