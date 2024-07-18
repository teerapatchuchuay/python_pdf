from flask import Flask, request, render_template, send_file, session
import pdfplumber
import pandas as pd
from io import BytesIO

app = Flask(__name__)
app.secret_key = b'_5#y2L"F4Q8z\n\xec]/'

def extract_data_from_pdf(pdf_file):
    data = []

    with pdfplumber.open(pdf_file) as pdf:
        for page in pdf.pages:
            # Extract text lines
            text = page.extract_text()
            if text:
                lines = text.split('\n')
                for line in lines:
                    parts = line.split()
                    if len(parts) >= 4:
                        description = " ".join(parts[:-3])
                        rate = parts[-3]
                        hours = parts[-2]
                        amount = parts[-1]
                        data.append([description, rate, hours, amount])

            # Extract tables
            tables = page.extract_tables()
            for table in tables:
                for row in table:
                    if len(row) >= 4:
                        description = row[0]
                        rate = row[-3]
                        hours = row[-2]
                        amount = row[-1]
                        data.append([description, rate, hours, amount])

    return data

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
            session['table_data'] = extract_data_from_pdf(file)
            return render_template('index.html', table_data=session['table_data'])
        except Exception as e:
            return f'Error processing file: {e}', 500
    else:
        return 'Invalid file format', 400

@app.route('/download')
def download_file():
    try:
        if 'table_data' in session:
            table_data = session['table_data']
            output = BytesIO()
            df = pd.DataFrame(table_data, columns=['Description', 'Rate', 'Hours', 'Amount'])
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df.to_excel(writer, index=False)
            output.seek(0)
            return send_file(output, as_attachment=True, download_name='output.xlsx', mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
        else:
            return 'No data to download', 404
    except Exception as e:
        return f'Error creating Excel file: {e}', 500

if __name__ == '__main__':
    app.run(debug=True)
