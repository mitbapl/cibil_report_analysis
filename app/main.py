import re
import spacy
import pdfplumber
import pandas as pd
from flask import Flask, request, render_template, send_file
from werkzeug.utils import secure_filename
import os
from io import BytesIO

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads/'  # Folder to store uploaded files

# Initialize spaCy model
nlp = spacy.load("en_core_web_sm")

def extract_text_from_pdf(pdf_path):
    """Extract all text from the PDF."""
    text_data = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text_data.append(page.extract_text())  # Extract text from each page
    return "\n".join(text_data)  # Combine all the pages' text

def process_pdf_text_with_spacy_PD(text):
    """Process PDF text and extract personal details using spaCy and regex."""
    sections = re.findall(r'CONSUMER CIR\n([\s\S]*?)(?=\nACCOUNT DATES AMOUNTS STATUS)', text)
    data_list_PD = []

    patterns = {
        'DATE': r'DATE:(\S+)(?=\b|$)',
        'MEMBER ID': r'MEMBER ID: (.+?)(?=\b|$)', 
        'TIME': r'TIME: (\S+)(?=\b|$)',
        'NAME': r'NAME: ([\s\S]*?)(?=\n)',
        'DATE OF BIRTH': r'DATE OF BIRTH: (\S+)(?=\b|$)', 
        'GENDER': r'GENDER: (.+?)(?=\b|$)',
        'CREDITVISION® SCORE': r'CREDITVISION® SCORE (.+?)(?=\b|$)',
        'INCOME TAX ID': r'\b([A-Z]{5}[0-9]{4}[A-Z])\b',
        'VOTER ID NUMBER': r'VOTER ID NUMBER (.+?)(?=\b|$)',
        'LICENSE NUMBER': r'LICENSE NUMBER (.+?)(?=\b|$)',
        'UNIVERSAL ID NUMBER (UID)': r'\b([0-9]{12})\b',
        'OFFICE PHONE': r'OFFICE PHONE (.+?)(?=\b|$)',
        'MOBILE PHONE': r'MOBILE PHONE (.+?)(?=\b|$)',
        'All Accounts TOTAL': r'All Accounts TOTAL: ([\d,]+)(?=\b|$)',
        'HIGH CR/SANC. AMT': r'HIGH CR/SANC. AMT: ([\d,]+)(?=\b|$)', 
        'CURRENT': r'CURRENT: (\S+)(?=\b|$)',
        'OVERDUE': r'OVERDUE: ([\d,]+)(?=\b|$)',
        'RECENT': r'RECENT: (\S+)(?=\b|$)',
        'OLDEST': r'OLDEST: (\S+)(?=\b|$)',
        'ZERO-BALANCE': r'ZERO-BALANCE: (.+?)(?=\b|$)\n',
    }

    for section in sections:
        doc = nlp(section)
        fields = {field: None for field in patterns}  # Initialize fields
        for field, pattern in patterns.items():
            match = re.search(pattern, section)
            if match:
                fields[field] = match.group(1).strip()
        data_list_PD.append(fields)
    
    return data_list_PD

def process_pdf_text_with_spacy_CD(text):
    """Process PDF text and extract credit details using spaCy and regex."""
    sections = re.findall(r'ACCOUNT(?:.*?)\nTYPE:.*?(?=\nACCOUNT|\Z)', text, re.DOTALL)
    data_list_CD = []

    patterns = {
        'MEMBER NAME': r'MEMBER NAME:\s*(.*)',
        'ACCOUNT NUMBER': r'ACCOUNT NUMBER:\s*(.*)',
        'TYPE': r'TYPE:\s*(.*)',
        'OWNERSHIP': r'OWNERSHIP:\s*(.*)',
        'OPENED': r'OPENED:\s*(.*)',
        'LAST PAYMENT': r'LAST PAYMENT:\s*(.*)',
        'CLOSED': r'CLOSED:\s*(.*)',
        'SANCTIONED': r'SANCTIONED:\s*(.*)',
        'CURRENT BALANCE': r'CURRENT BALANCE:\s*(.*)',
        'OVERDUE': r'OVERDUE:\s*(.*)',
        'DPD': r'DPD:\s*(.*)',
    }

    for section in sections:
        doc = nlp(section)
        fields = {field: None for field in patterns}  # Initialize fields
        for field, pattern in patterns.items():
            match = re.search(pattern, section)
            if match:
                fields[field] = match.group(1).strip()
        data_list_CD.append(fields)

    return data_list_CD

def credit_analysis(credit_details):
    """Perform credit analysis based on the extracted credit details."""
    total_accounts = len(credit_details)
    total_overdue_accounts = sum(1 for cd in credit_details if cd['OVERDUE'] and int(cd['OVERDUE']) > 0)
    total_sanctioned_amount = sum(float(cd['SANCTIONED'].replace(',', '')) for cd in credit_details if cd['SANCTIONED'])
    total_overdue_amount = sum(float(cd['OVERDUE'].replace(',', '')) for cd in credit_details if cd['OVERDUE'])
    total_current_balance = sum(float(cd['CURRENT BALANCE'].replace(',', '')) for cd in credit_details if cd['CURRENT BALANCE'])
    
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
    """Save extracted data to an Excel file."""
    output = BytesIO()  # In-memory output stream
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        pd.DataFrame(personal_details).to_excel(writer, sheet_name='Personal Details', index=False)
        pd.DataFrame(credit_details).to_excel(writer, sheet_name='Credit Details', index=False)
        pd.DataFrame([analysis_data]).to_excel(writer, sheet_name='Credit Analysis', index=False)

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
            # Extract text from the uploaded PDF
            text_data = extract_text_from_pdf(file_path)
            
            # Extract personal and credit details using NLP-based functions
            personal_details = process_pdf_text_with_spacy_PD(text_data)
            credit_details = process_pdf_text_with_spacy_CD(text_data)

            # Perform credit analysis on the extracted credit details
            analysis_data = credit_analysis(credit_details)

            # Save the results to an Excel file
            excel_output = save_to_excel(personal_details, credit_details, analysis_data)

            return send_file(excel_output, as_attachment=True, download_name="extracted_credit_report.xlsx")

        except Exception as e:
            return f"An error occurred while processing the file: {str(e)}"
    
    return "Invalid file format. Please upload a PDF."

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000)
