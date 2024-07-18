from flask import Flask, request, render_template, send_file, session
import pdfplumber
import pandas as pd
from io import BytesIO

app = Flask(__name__)
app.secret_key = b'_5#y2L"F4Q8z\n\xec]/'

def extract_bank_details(pdf_file):
    bank = account_name = bsb = "-"  
    
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
                        return bank, account_name, bsb  

    return bank, account_name, bsb  

def extract_data_from_pdf(pdf_file):
    data = []
    invoice_date = "-"  
    sub_total = tax = total = balance_due = "-"  

    with pdfplumber.open(pdf_file) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if text:
                lines = text.split('\n')
                for line in lines:
                    if "Invoice Date" in line and invoice_date == "-":
                        invoice_date = line.split("Invoice Date", 1)[1].strip() or "-"
                    elif "Sub Total" in line:
                        sub_total = line.split("Sub Total", 1)[1].strip() or "-"
                    elif "Tax" in line:
                        tax = line.split("Tax", 1)[1].strip() or "-"
                    elif "Total" in line:
                        total = line.split("Total", 1)[1].strip() or "-"
                    elif "Balance Due" in line:
                        balance_due = line.split("Balance Due", 1)[1].strip() or "-"
                    else:
                        parts = line.split()
                        if len(parts) >= 4 and is_currency(parts[-1]) and is_currency(parts[-2]):
                            item_description = " ".join(parts[:-2]) or "-"
                            qty = "-"  
                            rate = parts[-2] or "-"
                            amount = parts[-1] or "-"
                            
                            if len(parts) >= 5 and parts[-3].isdigit():
                                qty = parts[-3]
                                item_description = " ".join(parts[:-3]) or "-"
                            
                            data.append([item_description, qty, rate, amount, "-", "-", "-", "-", "-", "-", "-", "-"])
    
    data.append(["-", "-", "-", "-", "-", "-", "-", sub_total, tax, total, balance_due, invoice_date])
    
    return data


def is_currency(s):
    try:
        float(s.replace('$', '').replace(',', ''))
        return True
    except ValueError:
        return False

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
            df = pd.DataFrame(table_data, columns=['Item & Description', 'Qty', 'Rate', 'Amount', 'Bank', 'Account Name', 'BSB', 'Sub Total', 'Tax', 'Total', 'Balance Due', 'Invoice Date'])
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
