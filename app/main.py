from werkzeug.utils import secure_filename
from flask import Flask, request, render_template, send_file
import pandas as pd
import os
from io import BytesIO
import pdfplumber
import re

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads/'  # Folder to store uploaded files

@app.errorhandler(500)
def internal_error(error):
    return "Internal Server Error", 500

def extract_data_from_pdf(pdf_path):
    text_data = []
    tables = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text_data.append(page.extract_text())  # Extract page text
            page_tables = page.extract_tables()    # Extract tables if any
            for table in page_tables:
                df = pd.DataFrame(table[1:], columns=table[0])  # Create DataFrame using the first row as header
                df.columns = df.columns.str.strip()  # Normalize column names
                df = df.loc[:, ~df.columns.duplicated()]  # Remove duplicate columns
                tables.append(df)

    # Combine all extracted tables and return
    combined_table = pd.concat(tables, ignore_index=True) if tables else pd.DataFrame()
    
    return combined_table, " ".join(text_data)  # Return both table and raw text data

def extract_personal_details(text_data):
    # Using regular expressions to extract specific details from raw text
    name = re.search(r"NAME:\s*([A-Za-z\s]+)", text_data)
    dob = re.search(r"DATE OF BIRTH:\s*([0-9\-]+)", text_data)
    gender = re.search(r"GENDER:\s*(\w+)", text_data)
    cibil_score = re.search(r"CIBIL TRANSUNION SCORE:\s*(\d+)", text_data)
    pan = re.search(r"INCOME TAX ID:\s*([A-Z0-9]+)", text_data)
    address_matches = re.findall(r"ADDRESS:\s*([\w\s,]+)\s*TAMIL NADU\s*[0-9]+", text_data)
    
    personal_data = {
        'Name': name.group(1) if name else None,
        'Date of Birth': dob.group(1) if dob else None,
        'Gender': gender.group(1) if gender else None,
        'Credit Vision Score': cibil_score.group(1) if cibil_score else None,
        'PAN': pan.group(1) if pan else None,
        'Addresses': address_matches if address_matches else [],
        'Mobile Phones': [],  # Initialize as an empty list
        'Office Phone': None,
        'Email ID': None
    }

    return personal_data

def extract_credit_details(text_data):
    # Extract credit account details using regular expressions
    accounts = []
    account_pattern = re.compile(r"ACCOUNT(?:\s+MEMBER NAME:.+?)\s+(TYPE:.+?)(?:\s+DATES|END OF REPORT)", re.DOTALL)
    account_matches = account_pattern.findall(text_data)

    for account in account_matches:
        account_data = {
            'Member Name': re.search(r"MEMBER NAME:\s*(.+)", account).group(1) if re.search(r"MEMBER NAME:\s*(.+)", account) else None,
            'Account Number': re.search(r"ACCOUNT NUMBER:\s*(.+)", account).group(1) if re.search(r"ACCOUNT NUMBER:\s*(.+)", account) else None,
            'Type': re.search(r"TYPE:\s*(.+)", account).group(1) if re.search(r"TYPE:\s*(.+)", account) else None,
            'Opened': re.search(r"OPENED:\s*(.+)", account).group(1) if re.search(r"OPENED:\s*(.+)", account) else None,
            'Sanctioned': re.search(r"SANCTIONED:\s*(\d+)", account).group(1) if re.search(r"SANCTIONED:\s*(\d+)", account) else None,
            'Current Balance': re.search(r"CURRENT BALANCE:\s*(\d+)", account).group(1) if re.search(r"CURRENT BALANCE:\s*(\d+)", account) else None,
            'Overdue': re.search(r"OVERDUE:\s*(\d+)", account).group(1) if re.search(r"OVERDUE:\s*(\d+)", account) else None,
            'DPD': re.search(r"DPD:\s*(\d+)", account).group(1) if re.search(r"DPD:\s*(\d+)", account) else None,
        }
        accounts.append(account_data)
    
    # Convert to DataFrame
    credit_data = pd.DataFrame(accounts)
    
    return credit_data

def credit_analysis(credit_details):
    total_accounts = len(credit_details)
    total_overdue_accounts = (credit_details['Overdue'].astype(int) > 0).sum() if 'Overdue' in credit_details.columns else 0
    total_sanctioned_amount = credit_details['Sanctioned'].astype(float).sum() if 'Sanctioned' in credit_details.columns else 0
    total_overdue_amount = credit_details['Overdue'].astype(float).sum() if 'Overdue' in credit_details.columns else 0
    total_current_balance = credit_details['Current Balance'].astype(float).sum() if 'Current Balance' in credit_details.columns else 0
    average_dpd = credit_details['DPD'].astype(float).mean() if 'DPD' in credit_details.columns else None

    credit_utilization = (total_current_balance / total_sanctioned_amount * 100) if total_sanctioned_amount > 0 else None
    high_risk_accounts = (credit_details['DPD'].astype(int) > 30).sum() if 'DPD' in credit_details.columns else 0

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
            tables, text_data = extract_data_from_pdf(file_path)

            if tables.empty and not text_data:
                return "No data extracted from the PDF."

            personal_details = extract_personal_details(text_data)
            credit_details = extract_credit_details(text_data)
            analysis_data = credit_analysis(credit_details)

            excel_output = save_to_excel(personal_details, credit_details, analysis_data)

            return send_file(excel_output, as_attachment=True, download_name="extracted_credit_report.xlsx")

        except Exception as e:
            return f"An error occurred while processing the file: {str(e)}"
    
    return "Invalid file format. Please upload a PDF."

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000)
