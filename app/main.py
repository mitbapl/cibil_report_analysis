from werkzeug.utils import secure_filename
from flask import Flask, request, render_template, send_file
import pandas as pd
import os
from io import BytesIO
import pdfplumber

app = Flask(__name__)
app.config['DEBUG'] = True
app.config['UPLOAD_FOLDER'] = 'uploads/'  # Folder to store uploaded files

@app.errorhandler(500)
def internal_error(error):
    return "Internal Server Error", 500

def extract_data_from_pdf(pdf_path):
    tables = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            page_tables = page.extract_tables()
            for table in page_tables:
                df = pd.DataFrame(table[1:], columns=table[0])  # Create DataFrame from extracted table
                df.columns = df.columns.str.strip()  # Normalize column names
                tables.append(df)
    
    return pd.concat(tables, ignore_index=True) if tables else pd.DataFrame()  # Return combined DataFrame

def process_data(df):
    # Extract personal details
    personal_data = {
        'Name': df['NAME'].iloc[0] if 'NAME' in df else None,
        'Date of Birth': df['DATE OF BIRTH'].iloc[0] if 'DATE OF BIRTH' in df else None,
        'Gender': df['GENDER'].iloc[0] if 'GENDER' in df else None,
        'Credit Vision Score': df['CIBIL TRANSUNION SCORE'].iloc[0] if 'CIBIL TRANSUNION SCORE' in df else None,
        'PAN': df['INCOME TAX ID'].iloc[0] if 'INCOME TAX ID' in df else None,
        'Mobile Phones': df[['MOBILE PHONE1', 'MOBILE PHONE2']].dropna().values.flatten().tolist(),
        'Office Phone': df['OFFICE PHONE'].iloc[0] if 'OFFICE PHONE' in df else None,
        'Email ID': df['EMAIL ID'].iloc[0] if 'EMAIL ID' in df else None,
        'Addresses': df['ADDRESS'].dropna().tolist() if 'ADDRESS' in df else []
    }

    # Extract credit details
    credit_data = df[['MEMBER NAME', 'ACCOUNT NUMBER', 'OPENED', 'SANCTIONED', 
                       'CURRENT BALANCE', 'OVERDUE', 'DPD', 'TYPE', 
                       'OWNERSHIP', 'LAST PAYMENT', 'CLOSED']].copy()
    
    # Perform credit analysis
    analysis_data = {
        'Total Accounts': len(credit_data),
        'Total Overdue Accounts': (credit_data['OVERDUE'] > 0).sum(),
        'Total Sanctioned Amount': credit_data['SANCTIONED'].sum(),
        'Total Overdue Amount': credit_data['OVERDUE'].sum(),
        'Credit Utilization (%)': (credit_data['CURRENT BALANCE'].sum() / 
                                   credit_data['SANCTIONED'].sum() * 100) 
                                   if credit_data['SANCTIONED'].sum() > 0 else None,
        'Average DPD': credit_data['DPD'].mean() if 'DPD' in credit_data.columns else None,
        'High Risk Accounts': (credit_data['DPD'] > 30).sum()
    }
    
    return personal_data, credit_data, analysis_data

def save_to_excel(personal_details, credit_details, analysis_data):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        pd.DataFrame([personal_details]).to_excel(writer, sheet_name='Personal Details', index=False)
        credit_details.to_excel(writer, sheet_name='Credit Details', index=False)
        pd.DataFrame([analysis_data]).to_excel(writer, sheet_name='CIBIL Analysis', index=False)
    output.seek(0)
    return output

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files or request.files['file'].filename == '':
        return 'No selected file or no file part'

    file = request.files['file']
    if file and allowed_file(file.filename):
        file_path = os.path.join(app.config['UPLOAD_FOLDER'], secure_filename(file.filename))
        os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
        file.save(file_path)

        try:
            extracted_data = extract_data_from_pdf(file_path)

            if extracted_data.empty:
                return "No data extracted from the PDF."

            personal_details, credit_details, analysis_data = process_data(extracted_data)
            excel_output = save_to_excel(personal_details, credit_details, analysis_data)

            return send_file(excel_output, as_attachment=True, download_name="extracted_credit_report.xlsx")

        except Exception as e:
            return f"An error occurred while processing the file: {str(e)}"
    
    return "Invalid file format. Please upload a PDF."

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() == 'pdf'

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000)
