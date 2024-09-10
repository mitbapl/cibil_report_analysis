from flask import Flask, request, render_template, send_file
import re
import pandas as pd
import tabula
import os
from io import BytesIO
from app import create_app

app = create_app()
app.config['DEBUG'] = True
app.config['UPLOAD_FOLDER'] = 'uploads/'  # Folder to store uploaded files
@app.errorhandler(500)
def internal_error(error):
    return "Internal Server Error", 500

def extract_table_from_pdf(pdf_path):
    """
    Extract tables from the given PDF using Tabula.
    """
    tables = tabula.read_pdf(pdf_path, pages="all", multiple_tables=True)
    combined_df = pd.concat(tables, ignore_index=True)
    return combined_df

def cibil_score_analysis(cibil_score):
    """
    Analyze the CIBIL score and categorize it.
    """
    if cibil_score == 'NA' or cibil_score == 'NH':
        return "No credit history or no recent activity."
    try:
        score = int(cibil_score)
        if 750 <= score <= 900:
            return "Good: High chances of loan approval."
        elif 600 <= score < 750:
            return "Average: Medium chances, further analysis required."
        elif 300 <= score < 600:
            return "Low: High-risk category."
    except ValueError:
        return "Invalid CIBIL score."

def analyze_payment_history(df):
    """
    Analyze DPD (Days Past Due) for each credit facility.
    """
    df['DPD Flag'] = df['DPD'].apply(lambda x: 'Critical' if int(x) > 90 else ('Standard' if int(x) <= 90 else 'Unknown'))
    return df

def analyze_credit_utilization(balance, credit_limit):
    """
    Calculate and analyze credit utilization.
    """
    try:
        utilization = (balance / credit_limit) * 100
        if utilization > 30:
            return f"High: {utilization:.2f}% utilization."
        return f"Normal: {utilization:.2f}% utilization."
    except ZeroDivisionError:
        return "Credit limit not available."

def analyze_credit_inquiries(inquiries):
    """
    Analyze multiple recent credit inquiries.
    """
    recent_inquiries = [inquiry for inquiry in inquiries if '6 months' in inquiry or '1 year' in inquiry]
    if len(recent_inquiries) > 3:
        return "High credit hunger: Multiple credit inquiries detected."
    return "Normal inquiry behavior."

def age_of_credit_analysis(account_opening_dates):
    """
    Analyze the length of credit history based on account opening dates.
    """
    today = pd.Timestamp.now()
    credit_ages = [(today - pd.to_datetime(date)).days for date in account_opening_dates]
    average_age_years = sum(credit_ages) / len(credit_ages) / 365
    return f"Average Credit Age: {average_age_years:.2f} years."

def account_status_analysis(df):
    """
    Analyze account status (write-offs, settlements, disputes).
    """
    if 'Status' in df.columns and 'Dispute Status' in df.columns:
        write_offs = df[df['Status'] == 'Written-Off']
        settlements = df[df['Status'] == 'Settled']
        disputes = df[df['Dispute Status'] == 'Dispute']
        return {
            'write_offs': len(write_offs),
            'settlements': len(settlements),
            'disputes': len(disputes)
        }
    else:
        print("'Status' or 'Dispute Status' column is missing.")
        return {}


def employment_and_income_analysis(income, liabilities):
    """
    Calculate and analyze debt-to-income ratio.
    """
    try:
        dti_ratio = liabilities / income
        return f"Debt-to-Income Ratio: {dti_ratio:.2f}"
    except ZeroDivisionError:
        return "Income data not available."

def overdue_analysis(df):
    """
    Summarize overdue balances and flag overdue accounts, handling missing columns gracefully.
    """
    if df is not None and not df.empty:
        # Check if the 'Account Status' column exists
        if 'Account Status' in df.columns:
            overdue_accounts = df[df['Account Status'] == 'Overdue']
            
            # Check if 'Overdue Amount' exists in overdue_accounts
            if 'Overdue Amount' in overdue_accounts.columns:
                total_overdue = overdue_accounts['Overdue Amount'].sum()
                return f"Total overdue: {total_overdue}. Overdue accounts flagged."
            else:
                print("'Overdue Amount' column is missing.")
                return "No overdue amounts found."
        else:
            print("'Account Status' column is missing.")
            return "No overdue accounts found."
    else:
        print("No data available.")
        return "No data to analyze."

def save_to_excel(extracted_data, analysis_data):
    """
    Save the extracted data into an Excel file with two sheets:
    - Sheet1: Extracted tables from PDF.
    - Sheet2: CIBIL analysis data.
    """
    df_analysis = pd.DataFrame([analysis_data])
    output = BytesIO()

    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        extracted_data.to_excel(writer, sheet_name='Sheet1 - Extracted Tables', index=False)
        df_analysis.to_excel(writer, sheet_name='Sheet2 - CIBIL Analysis', index=False)
    
    output.seek(0)
    return output

@app.route('/')
def index():
    return render_template('index.html')

# @app.route('/upload', methods=['POST'])
@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return 'No file part'
    
    file = request.files['file']
    
    if file.filename == '':
        return 'No selected file'
    
    if file:
        file_path = os.path.join('/tmp', secure_filename(file.filename))
        file.save(file_path)
        
        # Extract tables using Tabula
        extracted_table = extract_table_from_pdf(file_path)
        
        if extracted_table is not None:
            # Extract personal and credit details
            personal_details = extract_personal_details(extracted_table)
            credit_details = extract_credit_details(extracted_table)
            
            # Perform credit analysis
            analysis_data = credit_analysis(credit_details)
            
            # Save everything to Excel
            excel_output = save_to_excel(personal_details, credit_details, analysis_data)
            
            return send_file(excel_output, as_attachment=True, download_name="credit_report_analysis.xlsx")
        else:
            return 'No table data found'

    return "Invalid file format. Please upload a PDF."

if __name__ == '__main__': 
     if not os.path.exists('uploads'): 
         os.makedirs('uploads') 
     app.run(host='0.0.0.0', port=5000)
