from werkzeug.utils import secure_filename
from flask import Flask, request, render_template, send_file
import pandas as pd
import os
from io import BytesIO
import pdfplumber

app = Flask(__name__)
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
                df = pd.DataFrame(table[1:], columns=table[0])  # Create DataFrame using the first row as header
                df.columns = df.columns.str.strip()  # Normalize column names
                df = df.loc[:, ~df.columns.duplicated()]  # Remove duplicate columns
                tables.append(df)

    # Combine all extracted tables and return
    return pd.concat(tables, ignore_index=True) if tables else pd.DataFrame()

def extract_personal_details(df):
    personal_data = {
        'Name': df['NAME'].iloc[0] if 'NAME' in df.columns else None,
        'Date of Birth': df['DATE OF BIRTH'].iloc[0] if 'DATE OF BIRTH' in df.columns else None,
        'Gender': df['GENDER'].iloc[0] if 'GENDER' in df.columns else None,
        'Credit Vision Score': df['CIBIL TRANSUNION SCORE'].iloc[0] if 'CIBIL TRANSUNION SCORE' in df.columns else None,
        'PAN': df['INCOME TAX ID'].iloc[0] if 'INCOME TAX ID' in df.columns else None,
        'Mobile Phones': [],  # Initialize as an empty list
        'Office Phone': df['OFFICE PHONE'].iloc[0] if 'OFFICE PHONE' in df.columns else None,
        'Email ID': df['EMAIL ID'].iloc[0] if 'EMAIL ID' in df.columns else None,
        'Addresses': []  # Initialize as an empty list
    }

    # Handle mobile phones (checking if columns exist)
    for col in ['MOBILE PHONE1', 'MOBILE PHONE2']:  # You can add more if needed
        if col in df.columns:
            personal_data['Mobile Phones'].extend(df[col].dropna().tolist())

    # Handle addresses (if the column exists)
    if 'ADDRESS' in df.columns:
        personal_data['Addresses'].extend(df['ADDRESS'].dropna().tolist())

    return personal_data

def extract_credit_details(df):
    credit_columns = ['MEMBER NAME', 'ACCOUNT NUMBER', 'OPENED', 'SANCTIONED', 
                      'CURRENT BALANCE', 'OVERDUE', 'DPD', 'TYPE', 
                      'OWNERSHIP', 'LAST PAYMENT', 'CLOSED']
    
    # Ensure only existing columns are selected
    credit_columns = [col for col in credit_columns if col in df.columns]
    credit_data = df[credit_columns].drop_duplicates().copy()  # Drop duplicate rows
    return credit_data

def credit_analysis(credit_details):
    total_accounts = len(credit_details)
    total_overdue_accounts = (credit_details['OVERDUE'] > 0).sum() if 'OVERDUE' in credit_details.columns else 0
    total_sanctioned_amount = credit_details['SANCTIONED'].sum() if 'SANCTIONED' in credit_details.columns else 0
    total_overdue_amount = credit_details['OVERDUE'].sum() if 'OVERDUE' in credit_details.columns else 0
    total_current_balance = credit_details['CURRENT BALANCE'].sum() if 'CURRENT BALANCE' in credit_details.columns else 0
    average_dpd = credit_details['DPD'].mean() if 'DPD' in credit_details.columns else None

    credit_utilization = (total_current_balance / total_sanctioned_amount * 100) if total_sanctioned_amount > 0 else None
    high_risk_accounts = (credit_details['DPD'] > 30).sum() if 'DPD' in credit_details.columns else 0

    analysis_data = {
        'Total Accounts': total_accounts,
        'Total Overdue Accounts': total_overdue_accounts,
        'Total Sanctioned Amount': total_sanctioned_amount,
        'Total Overdue Amount': total_overdue_amount,
        'Credit Utilization (%)': credit_utilization,
        'Average DPD': average_dpd,
        'High Risk Accounts': high_risk_accounts
    }

    return analysis_data

def save_to_excel(personal_details, credit_details, analysis_data):
    output = BytesIO()  # In-memory output stream
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
    if 'file' not in request.files:
        return 'No file part'
    
    file = request.files['file']
    
    if file.filename == '':
        return 'No selected file'
    
    if file:
        file_path = os.path.join(app.config['UPLOAD_FOLDER'], secure_filename(file.filename))
        os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)  # Create upload folder if it doesn't exist
        file.save(file_path)

        try:
            extracted_data = extract_data_from_pdf(file_path)

            if extracted_data.empty:
                return "No data extracted from the PDF."

            personal_details = extract_personal_details(extracted_data)
            credit_details = extract_credit_details(extracted_data)
            analysis_data = credit_analysis(credit_details)

            excel_output = save_to_excel(personal_details, credit_details, analysis_data)

            return send_file(excel_output, as_attachment=True, download_name="extracted_credit_report.xlsx")

        except Exception as e:
            return f"An error occurred while processing the file: {str(e)}"
    
    return "Invalid file format. Please upload a PDF."

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000)
