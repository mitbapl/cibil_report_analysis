from werkzeug.utils import secure_filename
from flask import Flask, request, render_template, send_file
import pandas as pd
import os
from io import BytesIO
from app import create_app
import pdfplumber

app = create_app()
app.config['DEBUG'] = True
app.config['UPLOAD_FOLDER'] = 'uploads/'  # Folder to store uploaded files

@app.errorhandler(500)
def internal_error(error):
    return "Internal Server Error", 500

def extract_credit_details(table):
    credit_data = []

    for index, row in table.iterrows():
        credit_record = {
            'Member Name': row.get('MEMBER NAME', None),
            'Account Number': row.get('ACCOUNT NUMBER', None),
            'Opened Date': row.get('OPENED', None),
            'Sanctioned Amount': row.get('SANCTIONED', None),
            'Current Balance': row.get('CURRENT BALANCE', None),
            'Overdue': row.get('OVERDUE', None),
            'DPD': row.get('DPD', None),  # Days Past Due
            'Loan Type': row.get('TYPE', None),
            'Ownership': row.get('OWNERSHIP', None),
            'Last Payment Date': row.get('LAST PAYMENT', None),
            'Closed Date': row.get('CLOSED', None),
        }
        credit_data.append(credit_record)

    return pd.DataFrame(credit_data)

# Function to extract personal details from the DataFrame
def extract_personal_details(table):
    personal_data = {
        'Name': None,
        'Date of Birth': None,
        'Gender': None,
        'Credit Vision Score': None,
        'PAN': None,
        'Mobile Phones': [],
        'Office Phone': None,
        'Email ID': None,
        'Addresses': [],
        'Total Accounts': None,
        'Total Overdue Accounts': None,
        'Total Sanctioned Amount': None,
        'Total Overdue Amount': None,
        'Credit Utilization': None,
        'Average DPD': None,
        'High Risk Accounts': None,
    }

    for index, row in table.iterrows():
        personal_data['Name'] = row.get('NAME', personal_data['Name'])
        personal_data['Date of Birth'] = row.get('DATE OF BIRTH', personal_data['Date of Birth'])
        personal_data['Gender'] = row.get('GENDER', personal_data['Gender'])
        personal_data['Credit Vision Score'] = row.get('CIBIL TRANSUNION SCORE', personal_data['Credit Vision Score'])
        personal_data['PAN'] = row.get('INCOME TAX ID', personal_data['PAN'])
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

    return personal_data

# Function to convert personal details dictionary to DataFrame
def convert_to_dataframe(personal_details_dict):
    return pd.DataFrame([personal_details_dict])

def credit_analysis(credit_details):
    total_accounts = len(credit_details)
    total_overdue_accounts = len(credit_details[credit_details['Overdue'] > 0])
    total_sanctioned_amount = credit_details['Sanctioned Amount'].sum()
    total_overdue_amount = credit_details['Overdue'].sum()
    total_current_balance = credit_details['Current Balance'].sum()
    average_dpd = credit_details['DPD'].mean() if 'DPD' in credit_details.columns else None

    # Calculate credit utilization percentage (sum of balances vs sanctioned amounts)
    credit_utilization = (total_current_balance / total_sanctioned_amount) * 100 if total_sanctioned_amount > 0 else None

    # High risk accounts based on DPD criteria (e.g., DPD > 30)
    high_risk_accounts = credit_details[credit_details['DPD'] > 30] 

    analysis_data = {
        'Total Accounts': total_accounts,
        'Total Overdue Accounts': total_overdue_accounts,
        'Total Sanctioned Amount': total_sanctioned_amount,
        'Total Overdue Amount': total_overdue_amount,
        'Credit Utilization (%)': credit_utilization,
        'Average DPD': average_dpd,
        'High Risk Accounts': len(high_risk_accounts)
    }

    return analysis_data

def save_to_excel(personal_details, credit_details, analysis_data):
    output = BytesIO()  # In-memory output stream

    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        # Save personal details to Sheet1
        personal_details_df = convert_to_dataframe(personal_details)
        personal_details_df.to_excel(writer, sheet_name='Sheet1 - Personal Details', index=False)

        # Save credit details to Sheet2
        credit_details.to_excel(writer, sheet_name='Sheet2 - Credit Details', index=False)

        # Save analysis data to Sheet3
        df_analysis = pd.DataFrame([analysis_data])
        df_analysis.to_excel(writer, sheet_name='Sheet3 - CIBIL Analysis', index=False)

    output.seek(0)
    return output

@app.route('/')
def index():
    return render_template('index.html')

def extract_table_from_pdf(pdf_path):
    """
    Extract tables and raw text from the given PDF using pdfplumber and normalize the data.
    """
    tables = []
    with pdfplumber.open(pdf_path) as pdf:
        for page_num, page in enumerate(pdf.pages):
            # Extract tables from each page
            page_tables = page.extract_tables()
            
            # Extract text to help in normalization
            raw_text = page.extract_text()
            print(f"Page {page_num + 1} - Extracted Tables: {len(page_tables)}")

            # Convert raw text to DataFrame or list of dicts (normalization logic goes here)
            for table in page_tables:
                # Extract tables as DataFrames
                df = pd.DataFrame(table[1:], columns=table[0])  # Assumes first row is the header
                df = df.loc[:, ~df.columns.duplicated()].copy()  # Remove duplicate columns
                print(f"Extracted DataFrame from page {page_num + 1}:\n{df}")  # Debug output
                tables.append(df)
    
    # Combine tables into a single DataFrame
    if tables:  # Only combine if tables have been extracted
        combined_df = pd.concat(tables, ignore_index=True)
        
        # Normalize data: Remove empty rows, clean headers, and standardize formats
        combined_df = clean_and_normalize_data(combined_df)
        print(f"Combined DataFrame:\n{combined_df}")  # Debug output
    else:
        print("No tables extracted.")
        return pd.DataFrame()  # Return an empty DataFrame if no tables were found

    return combined_df

def clean_and_normalize_data(df):
    """
    Normalize the extracted DataFrame by cleaning headers, removing unnecessary rows, and standardizing columns.
    """
    # Strip whitespace from column names
    df.columns = [col.strip() if isinstance(col, str) else col for col in df.columns]
    
    # Ensure the index is unique by resetting it (if necessary)
    df = df.reset_index(drop=True)
    
    # Remove rows where the entire row is NaN
    df = df.dropna(how="all")
    
    # Here, we will not remove any rows for now to see what we have
    # Return the DataFrame without filtering
    return df


@app.route('/upload', methods=['POST'])
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
            # Extract tables using pdfplumber
            extracted_data = extract_table_from_pdf(file_path)

            if extracted_data.empty:
                return "No data extracted from the PDF."

            # Analyze data if needed and save to Excel
            excel_output = save_to_excel(extracted_data)

            return send_file(excel_output, as_attachment=True, download_name="extracted_credit_report.xlsx")

        except Exception as e:
            return f"An error occurred while processing the file: {str(e)}"
    
    return "Invalid file format. Please upload a PDF."


if __name__ == '__main__':
     if not os.path.exists('uploads'): 
         os.makedirs('uploads') 
     app.run(host='0.0.0.0', port=5000)
