from flask import Flask, request, render_template, send_file, session
import pdfplumber
import pandas as pd
from io import BytesIO

app = Flask(__name__)
app.secret_key = b'_5#y2L"F4Q8z\n\xec]/'

def extract_bank_details(pdf_file):
    bank = account_name = bsb = "-"  # Default values
    
    with pdfplumber.open(pdf_file) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if text:
                lines = text.split('\n')
                for line in lines:
                    if 'Bank' in line:
                        bank = line.split('Bank', 1)[1].strip() or "-"
                    elif 'Account Name' in line:
                        account_name = line.split('Account Name', 1)[1].strip() or "-"
                    elif 'BSB' in line:
                        bsb = line.split('BSB', 1)[1].strip() or "-"
                        return bank, account_name, bsb  # Return once all data is found

    return bank, account_name, bsb  # Return defaults if not found

def extract_data_from_pdf(pdf_file):
    data = []
    bank, account_name, bsb = extract_bank_details(pdf_file)
    info_extracted = False  # Flag to ensure bank details are added only once

    with pdfplumber.open(pdf_file) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if text:
                lines = text.split('\n')
                for line in lines:
                    parts = line.split()
                    if len(parts) >= 4 and parts[-1].startswith('$') and parts[-3].startswith('$'):
                        description = " ".join(parts[:-3]) or "-"
                        rate = parts[-3] or "-"
                        hours = parts[-2] or "-"
                        amount = parts[-1] or "-"
                        if not info_extracted:
                            data.append([description, rate, hours, amount, bank, account_name, bsb])
                            info_extracted = True
                        else:
                            data.append([description, rate, hours, amount, "-", "-", "-"])  # Add placeholders for subsequent rows

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
            df = pd.DataFrame(table_data, columns=['Description', 'Rate', 'Hours', 'Amount', 'Bank', 'Account Name', 'BSB'])
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
