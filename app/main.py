from werkzeug.utils import secure_filename
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

def extract_credit_details(table):
    credit_data = []

    for index, row in table.iterrows():
        credit_record = {
            'Member Name': row.get('MEMBER NAME', None),
            'Account Number': row.get('ACCOUNT NUMBER', None),
            'Opened Date': row.get('OPENED', None),
            'Sanctioned Amount': row.get('SANCTIONED', None),
            'Last Payment Date': row.get('LAST PAYMENT', None),
            'Current Balance': row.get('CURRENT BALANCE', None),
            'Closed Date': row.get('CLOSED', None),
            'Loan Type': row.get('TYPE', None),
            'EMI': row.get('EMI', None),  # If available, or remove this line
            'Overdue': row.get('OVERDUE', None),
            'Ownership': row.get('OWNERSHIP', None),
            'DPD': row.get('DPD', None),  # Days Past Due
        }
        credit_data.append(credit_record)

    return pd.DataFrame(credit_data)

# Function to extract personal details from the DataFrame
def extract_personal_details(table):
    personal_data = {
        'Date': None,
        'Member ID': None,
        'Time': None,
        'Name': None,
        'Date of Birth': None,
        'Gender': None,
        'Credit Vision Score': None,
        'PAN': None,
        'Voter ID': None,
        'License Number': None,
        'UID': None,
        'Office Phone': None,
        'Mobile Phones': [],
        'Email ID': None,
        'Addresses': [],
        'All Accounts TOTAL': None,
        'High CR/Sanc. Amt': None,
        'Current': None,
        'Overdue': None,
        'Recent': None,
        'Oldest': None,
        'Zero-Balance': None,
    }

    for index, row in table.iterrows():
        # Mapping based on the structure in the report
        personal_data['Date'] = row.get('DATE', personal_data['Date'])
        personal_data['Member ID'] = row.get('MEMBER ID', personal_data['Member ID'])
        personal_data['Time'] = row.get('TIME', personal_data['Time'])
        personal_data['Name'] = row.get('NAME', personal_data['Name'])
        personal_data['Date of Birth'] = row.get('DATE OF BIRTH', personal_data['Date of Birth'])
        personal_data['Gender'] = row.get('GENDER', personal_data['Gender'])
        personal_data['Credit Vision Score'] = row.get('CIBIL TRANSUNION SCORE', personal_data['Credit Vision Score'])  # Update this field
        personal_data['PAN'] = row.get('INCOME TAX ID', personal_data['PAN'])  # Updated to match the document
        personal_data['Voter ID'] = row.get('VOTER ID', personal_data['Voter ID'])
        personal_data['License Number'] = row.get('PASSPORT NO', personal_data['License Number'])  # Assuming no 'LICENSE NUMBER' exists
        personal_data['UID'] = row.get('UNIVERSAL ID NUMBER (UID)', personal_data['UID'])
        personal_data['Office Phone'] = row.get('OFFICE PHONE', personal_data['Office Phone'])
        mobile1 = row.get('MOBILE PHONE1', None)
        mobile2 = row.get('MOBILE PHONE2', None)
        if mobile1:
            personal_data['Mobile Phones'].append(mobile1)
        if mobile2:
            personal_data['Mobile Phones'].append(mobile2)
        personal_data['Email ID'] = row.get('EMAIL ID', personal_data['Email ID'])
        address = row.get('ADDRESS', None)
        if address:
            personal_data['Addresses'].append(address)

        personal_data['All Accounts TOTAL'] = row.get('All Accounts TOTAL', personal_data['All Accounts TOTAL'])
        personal_data['High CR/Sanc. Amt'] = row.get('HIGH CR/SANC. AMT', personal_data['High CR/Sanc. Amt'])
        personal_data['Current'] = row.get('CURRENT', personal_data['Current'])
        personal_data['Overdue'] = row.get('OVERDUE', personal_data['Overdue'])
        personal_data['Recent'] = row.get('RECENT', personal_data['Recent'])
        personal_data['Oldest'] = row.get('OLDEST', personal_data['Oldest'])
        personal_data['Zero-Balance'] = row.get('ZERO-BALANCE', personal_data['Zero-Balance'])

    return personal_data

# Function to convert personal details dictionary to DataFrame
def convert_to_dataframe(personal_details_dict):
    return pd.DataFrame([personal_details_dict])

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

def credit_analysis(credit_details):
    # Initialize metrics to capture various credit analysis insights
    total_accounts = len(credit_details)
    total_overdue_accounts = len(credit_details[credit_details['Overdue'] > 0])
    total_closed_accounts = len(credit_details[credit_details['Closed Date'].notnull()])
    
    total_sanctioned_amount = credit_details['Sanctioned Amount'].sum()
    total_current_balance = credit_details['Current Balance'].sum()
    total_overdue_amount = credit_details['Overdue'].sum()
    
    # Calculate average overdue and average DPD (Days Past Due)
    average_dpd = credit_details['DPD'].mean() if 'DPD' in credit_details.columns else None
    
    # Calculate credit utilization percentage (sum of balances vs sanctioned amounts)
    credit_utilization = (total_current_balance / total_sanctioned_amount) * 100 if total_sanctioned_amount > 0 else None
    
    # Risk indicators: find if any accounts have high DPD, overdue amounts
    high_risk_accounts = credit_details[credit_details['DPD'] > 30]  # Example: flagging DPD over 30 days
    
    # Aggregate the analysis data into a dictionary
    analysis_data = {
        'Total Accounts': total_accounts,
        'Total Overdue Accounts': total_overdue_accounts,
        'Total Closed Accounts': total_closed_accounts,
        'Total Sanctioned Amount': total_sanctioned_amount,
        'Total Current Balance': total_current_balance,
        'Total Overdue Amount': total_overdue_amount,
        'Average DPD': average_dpd,
        'Credit Utilization (%)': credit_utilization,
        'High Risk Accounts': len(high_risk_accounts)
    }
    
    return analysis_data

def save_to_excel(personal_details, credit_details, analysis_data):
    """
    Save the extracted data into an Excel file with three sheets:
    - Sheet1: Personal details from the CIBIL report.
    - Sheet2: Credit details from the CIBIL report.
    - Sheet3: CIBIL analysis data.
    """
    df_analysis = pd.DataFrame([analysis_data])  # Convert analysis data to DataFrame
    output = BytesIO()  # In-memory output stream

    # Use Pandas Excel writer
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        # Save personal details to Sheet1
        personal_details.to_excel(writer, sheet_name='Sheet1 - Personal Details', index=False)
        
        # Save credit details to Sheet2
        credit_details.to_excel(writer, sheet_name='Sheet2 - Credit Details', index=False)
        
        # Save analysis data to Sheet3
        df_analysis.to_excel(writer, sheet_name='Sheet3 - CIBIL Analysis', index=False)

    # Seek to the beginning of the stream and return
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
        # Save the uploaded file temporarily
        file_path = os.path.join('/tmp', secure_filename(file.filename))
        file.save(file_path)

        try:
            # Extract tables using Tabula
            extracted_tables = tabula.read_pdf(file_path, pages="all", multiple_tables=True)

            # Initialize lists for personal details and credit details
            personal_details_list = []
            credit_details_list = []

            # Loop through each extracted table
            for idx, table in enumerate(extracted_tables):
                # Debug: Print table shape and type for inspection
                print(f"Table {idx}: Type: {type(table)}, Shape: {getattr(table, 'shape', 'Not a DataFrame')}")
                
                # Ensure the table is a DataFrame and has a valid 2D structure
                if isinstance(table, pd.DataFrame) and len(table.shape) == 2:
                    # Check if the DataFrame is valid (has at least 1 row and 1 column)
                    if table.shape[0] > 0 and table.shape[1] > 0:
                        print(f"Processing valid table {idx} with shape {table.shape}")

                        # Extract personal details
                        personal_details = extract_personal_details(table)
                        personal_details_df = convert_to_dataframe(personal_details)
                        if not personal_details_df.empty:
                            personal_details_list.append(personal_details_df)

                        # Extract credit details
                        credit_details = extract_credit_details(table)
                        credit_details_df = convert_to_dataframe(credit_details)
                        if not credit_details_df.empty:
                            credit_details_list.append(credit_details_df)
                    else:
                        print(f"Skipping empty table {idx} with shape {table.shape}")
                else:
                    print(f"Skipping invalid table {idx} with shape {table.shape} or invalid structure")

            # Combine all personal details and credit details into DataFrames
            all_personal_details = pd.concat(personal_details_list, ignore_index=True) if personal_details_list else pd.DataFrame()
            all_credit_details = pd.concat(credit_details_list, ignore_index=True) if credit_details_list else pd.DataFrame()

            # Debug: Print the final DataFrame shapes
            print(f"All Personal Details Shape: {all_personal_details.shape}")
            print(f"All Credit Details Shape: {all_credit_details.shape}")

            # Perform credit analysis
            analysis_data = credit_analysis(all_credit_details)

            # Save everything to Excel
            excel_output = save_to_excel(all_personal_details, all_credit_details, analysis_data)

            # Return the file as a downloadable response
            return send_file(excel_output, as_attachment=True, download_name="credit_report_analysis.xlsx")

        except Exception as e:
            # Return detailed error message for debugging
            return f"An error occurred while processing the file: {str(e)}"
    
    return "Invalid file format. Please upload a PDF."

    
if __name__ == '__main__': 
     if not os.path.exists('uploads'): 
         os.makedirs('uploads') 
     app.run(host='0.0.0.0', port=5000)
