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
            text_data.append(page.extract_text())  # Extract text from each page
    return "\n".join(text_data) 

def extract_personal_details(text_data):
    """Extract personal details using regular expressions from the text."""
    name = re.search(r"NAME:\s*([A-Za-z\s]+)", text_data)
    dob = re.search(r"DATE OF BIRTH:\s*([0-9\-]+)", text_data)
    gender = re.search(r"GENDER:\s*(\w+)", text_data)
    cibil_score = re.search(r"CIBIL TRANSUNION SCORE:\s*(\d+)", text_data)
    pan = re.search(r"INCOME TAX ID:\s*([A-Z0-9]+)", text_data)
    address_matches = re.findall(r"ADDRESS(?:ES)?:\s*([\w\s,]+)\s*[0-9]{6}", text_data)  # Matches the address ending with pincode
    
    personal_data = {
        'Name': name.group(1) if name else None,
        'Date of Birth': dob.group(1) if dob else None,
        'Gender': gender.group(1) if gender else None,
        'Credit Vision Score': cibil_score.group(1) if cibil_score else None,
        'PAN': pan.group(1) if pan else None,
        'Addresses': address_matches if address_matches else [],
    }

    return personal_data

def extract_credit_details(text_data):
    """Extract credit account details using regular expressions."""
    accounts = []
    
    account_pattern = re.compile(r"ACCOUNT\s+MEMBER NAME:(.*?)\nTYPE:(.*?)\nOWNERSHIP:(.*?)\n(?:DATES)\s+OPENED:(.*?)\nLAST PAYMENT:(.*?)\nCLOSED:(.*?)\nSANCTIONED:(.*?)\nCURRENT BALANCE:(.*?)\nOVERDUE:(.*?)\n", re.DOTALL)
    account_matches = account_pattern.findall(text_data)

    for account in account_matches:
        account_data = {
            'Member Name': account[0].strip(),
            'Type': account[1].strip(),
            'Ownership': account[2].strip(),
            'Opened': account[3].strip(),
            'Last Payment': account[4].strip(),
            'Closed': account[5].strip(),
            'Sanctioned': account[6].strip(),
            'Current Balance': account[7].strip(),
            'Overdue': account[8].strip(),
        }
        accounts.append(account_data)
    
    credit_data = pd.DataFrame(accounts)
    
    return credit_data

def credit_analysis(credit_details):
    """Perform credit analysis based on the extracted credit details."""
    total_accounts = len(credit_details)
    total_overdue_accounts = (credit_details['Overdue'].astype(int) > 0).sum() if 'Overdue' in credit_details.columns else 0
    total_sanctioned_amount = credit_details['Sanctioned'].astype(float).sum() if 'Sanctioned' in credit_details.columns else 0
    total_overdue_amount = credit_details['Overdue'].astype(float).sum() if 'Overdue' in credit_details.columns else 0
    total_current_balance = credit_details['Current Balance'].astype(float).sum() if 'Current Balance' in credit_details.columns else 0
    average_dpd = None  # Not calculated in this version since DPD isn't extracted

    credit_utilization = (total_current_balance / total_sanctioned_amount * 100) if total_sanctioned_amount > 0 else None

    analysis_data = {
        'Total Accounts': total_accounts,
        'Total Overdue Accounts': total_overdue_accounts,
        'Total Sanctioned Amount': total_sanctioned_amount,
        'Total Overdue Amount': total_overdue_amount,
        'Credit Utilization (%)': credit_utilization,
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
            # Extract text from PDF (corrected: expect only one return value)
            text_data = extract_text_from_pdf(pdf_path)
# Extract personal details and credit details using the text_data
            personal_details = extract_personal_details(text_data)
            credit_details = extract_credit_details(text_data)
# Perform the credit analysis on the extracted credit details
            analysis_data = credit_analysis(credit_details)
# Output the results
# personal_details, credit_details, analysis_data
            excel_output = save_to_excel(personal_details, credit_details, analysis_data)
            return send_file(excel_output, as_attachment=True, download_name="extracted_credit_report.xlsx")

        except Exception as e:
            return f"An error occurred while processing the file: {str(e)}"
    
    return "Invalid file format. Please upload a PDF."

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000)
