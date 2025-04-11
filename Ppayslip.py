import pandas as pd
import json
import openpyxl
from fpdf import FPDF

# Read the Excel filez
df = pd.read_excel("employees.xlsx.xlsx")

# Display current columns to verify names
print("Current columns in the DataFrame:")
print(df.columns)

# Select the required columns (using exact column names from your data)
df = df[['NAME', "EMPLOYEE'S ID", 'EMAIL', 'BASIC SALARY', 'ALLOWANCEES', 'DEDUCTIONS']]

# Net salary calculation (using exact column names from your data)
df['Net salary'] = df['BASIC SALARY'] + df['ALLOWANCEES'] - df['DEDUCTIONS']

# Format currency values
pd.options.display.float_format = '${:,.2f}'.format

# Display results with limited width to prevent console buffer issues
print("\nEmployee Salary Detail:")
print(df[['NAME', "EMPLOYEE'S ID", 'EMAIL', 'BASIC SALARY', 'ALLOWANCEES', 'DEDUCTIONS', 'Net salary']].to_string(max_colwidth=20))

import pandas as pd
from fpdf import FPDF
import os

# Create payslips directory if it doesn't exist
os.makedirs('payslips', exist_ok=True)

def create_payslip(df_row, filename):
    """Generate a single payslip PDF and save it"""
    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Arial", size=15)

    # Header
    pdf.cell(200, 10, txt="PAYSLIP", ln=True, align='C')
    pdf.ln(10)

    # Personal Info
    pdf.set_font("Arial", size=12)
    pdf.cell(200, 10, txt=f"Employee Name: {df_row['NAME']}", ln=True, align='L')
    pdf.cell(200, 10, txt=f"Employee ID: {df_row['EMPLOYEE\'S ID']}", ln=True, align='L')
    pdf.cell(200, 10, txt=f"Email: {df_row['EMAIL']}", ln=True, align='L')
    pdf.ln(10)

    # Salary Details
    pdf.set_font("Arial", 'B', size=12)
    pdf.cell(200, 10, txt="Salary Details:", ln=True, align='L')
    pdf.ln(5)

    pdf.set_font("Arial", size=10)
    pdf.cell(200, 10, txt=f"Basic Salary: ${df_row['BASIC SALARY']:.2f}", ln=True, align='L')
    pdf.cell(200, 10, txt=f"Allowances: ${df_row['ALLOWANCEES']:.2f}", ln=True, align='L')
    pdf.cell(200, 10, txt=f"Deductions: ${df_row['DEDUCTIONS']:.2f}", ln=True, align='L')
    pdf.ln(10)

    # Net Salary
    pdf.set_font("Arial", 'B', size=12)
    pdf.cell(200, 10, txt="Net Salary:", ln=True, align='L')
    pdf.set_font("Arial", size=10)
    pdf.cell(200, 10, txt=f"${df_row['Net salary']:.2f}", ln=True, align='L')

    # Output to file
    pdf.output(filename)

# Sample Data
sample_data = {
    'NAME': 'TINASHE WUTETE',
    "EMPLOYEE'S ID": 'stk001',
    'EMAIL': 'nashegraphix@gmail.com',
    'BASIC SALARY': 300,
    'ALLOWANCEES': 200,
    'DEDUCTIONS': 25,
}

# Calculate Net Salary
sample_data['Net salary'] = sample_data['BASIC SALARY'] + sample_data['ALLOWANCEES'] - sample_data['DEDUCTIONS']

# Create one payslip
create_payslip(sample_data, 'payslips/sample_payslip.pdf')
print("✅ Sample payslip generated: payslips/sample_payslip.pdf")

import smtplib
import os
from email.message import EmailMessage
import pandas as pd

# === Email Configuration ===
EMAIL_ADDRESS = "your_email@example.com"     # <-- Replace with your email
EMAIL_PASSWORD = "your_email_password"       # <-- Replace with your email app password
SMTP_SERVER = "smtp.gmail.com"
SMTP_PORT = 587

def send_payslip_email(to_email, name, employee_id):
    msg = EmailMessage()
    msg['Subject'] = "Your Payslip for This Month"
    msg['From'] = EMAIL_ADDRESS
    msg['To'] = to_email

    msg.set_content(f"""
    Dear {name},

    Please find attached your payslip for this month.

    Best regards,
    HR Department
    """)

    filename = f"payslips/{employee_id}.pdf"
    try:
        with open(filename, "rb") as f:
            file_data = f.read()
            msg.add_attachment(file_data, maintype="application", subtype="pdf", filename=os.path.basename(filename))
    except FileNotFoundError:
        print(f"❌ Payslip for {employee_id} not found.")
        return

    with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as smtp:
        smtp.starttls()
        smtp.login("kaduhwashaine20@gmail.com", "axcr gqzl tgpg czsy")
        smtp.send_message(msg)
        print(f"✅ Email sent to {name} ({to_email})")

# === Sample Employee DataFrame (or load from Excel) ===
data = {
    'NAME': ['TINASHE WUTETE', 'LLOYD CHOGARI', 'TAFADZWA MZAYA', 'CARLTON SITHOLE', 'DENNY BINGURA'],
    "EMPLOYEE'S ID": ['stk001', 'stk002', 'stk003', 'stk004', 'stk005'],
    'EMAIL': ['nashegraphix@gmail.com', 'lloyddonnel44@gmail.com', 'munhuharaswit@gmail.com',
              'spliffking16@gmail.com', 'demykadwell@gmail.com'],
}

df = pd.DataFrame(data)

# === Send Emails to All Employees ===
for _, row in df.iterrows():
    send_payslip_email(row['EMAIL'], row['NAME'], row["EMPLOYEE'S ID"])
