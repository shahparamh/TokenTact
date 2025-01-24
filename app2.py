from flask import Flask, render_template, request, jsonify, send_file
import smtplib
import gspread
from google.oauth2.service_account import Credentials
import pandas as pd
from flask import send_file
from io import BytesIO
from reportlab.pdfgen import canvas
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.platypus import Spacer
from reportlab.lib.styles import ParagraphStyle
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage

app = Flask(__name__)

@app.route('/')
def index():
    return render_template('index2.html')

@app.route('/validate_email', methods=['POST'])
def validate_email():
    email = request.form['email']
    if email.endswith(("@ahduni.edu.in")):
        return jsonify({'status': 'success', 'message': 'Email Verified'})
    else:
        return jsonify({'status': 'error', 'message': 'You are not eligible for this service'})

@app.route('/send_email', methods=['POST'])
def send_email():
    token_number = request.form['token_number']
    receiver_email = request.form['receiver_email']
    photo = request.files['photo']

    # Set up your email configuration
    sender_email = "mail"
    # Use the application-specific password generated for your script
    sender_password = "pass"
    # Set up Google Sheets API credentials
    credentials = Credentials.from_service_account_file('keys.json', scopes=['https://spreadsheets.google.com/feeds'])

    # Connect to the Google Sheet
    gc = gspread.authorize(credentials)
    spreadsheet_key = 'API KEY'
    worksheet = gc.open_by_key(spreadsheet_key).sheet1

    try:
        # Get information based on the entered token number
        row = worksheet.find(str(token_number)).row
        student_name = worksheet.cell(row, 2).value
        receiver_email = worksheet.cell(row, 3).value
        # Increment the count in the 'count' column
        count_column = worksheet.find("count").col
        count_value = int(worksheet.cell(row, count_column).value)
        worksheet.update_cell(row, count_column, count_value + 1)
    except Exception as e:
        return jsonify({'status': 'error', 'message': 'Token number not found'})

    # Compose the email message
    subject = "Wrong Parking Alert"
    body = f"Dear {student_name},\n\n" \
           f"Vehicle with token number {token_number} is parked wrongly. Please take necessary action.\n\n" \
           f"AHMEDABAD UNIVERSITY/ByteBuilder/2023"

    # Set up the SMTP server and send the email
    try:
        with smtplib.SMTP("smtp.gmail.com", 587) as server:
            server.starttls()
            server.login(sender_email, sender_password)
            msg = MIMEMultipart()
            msg['From'] = sender_email
            msg['To'] = receiver_email
            msg['Subject'] = subject

            msg.attach(MIMEText(body, 'plain'))
            msg.attach(MIMEImage(photo.read()))
            server.sendmail(sender_email, receiver_email, msg.as_string())
        return jsonify({'status': 'success', 'message': 'Email sent successfully!'})
    except Exception as e:
        return jsonify({'status': 'error', 'message': f'An error occurred: {e}'})

@app.route('/download_report')
def download_report():
    try:
        # Set up Google Sheets API credentials
        credentials = Credentials.from_service_account_file('keys.json', scopes=['https://spreadsheets.google.com/feeds'])

        # Connect to the Google Sheet
        gc = gspread.authorize(credentials)
        spreadsheet_key = 'API KEY'
        worksheet = gc.open_by_key(spreadsheet_key).sheet1

        # Read data into a DataFrame
        data = worksheet.get_all_records()
        df = pd.DataFrame(data)

        # Save the DataFrame as a CSV
        csv_path = "report.csv"
        df.to_csv(csv_path, index=False, header=True)

        # Generate PDF using the CSV data
        pdf_path = "report.pdf"
        title = "Report (TokenTact)"
        generate_pdf(pdf_path, df, title)

        return send_file(pdf_path, as_attachment=True)
    except Exception as e:
        return jsonify({'status': 'error', 'message': f'An error occurred: {e}'})

def generate_pdf(pdf_path, data, title):
    doc = SimpleDocTemplate(pdf_path, pagesize=letter)
    content = []

    # Add title
    styles = getSampleStyleSheet()
    title_text = "<u>{}</u>".format(title)
    title_paragraph = Paragraph(title_text, styles['Title'])
    content.append(title_paragraph)

    # Add table header
    header = data.columns.tolist()
    table_data = [header] + [data.iloc[i].tolist() for i in range(len(data))]

    # Create table and apply style
    style = TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('GRID', (0, 0), (-1, -1), 1, colors.black),
    ])

    # Create table and apply style
    table = Table(table_data, style=style)

    # Add table to content
    content.append(table)

    content.append(Spacer(1, 20))

    # Add names of students with count greater than 5 at the bottom
    count_greater_than_5 = data[data['count'] > 5]['student_name'].tolist()
    bold_style = ParagraphStyle(name='Bold', parent=styles['Normal'], fontName='Helvetica-Bold')

    if count_greater_than_5:
        violation_text = " -> "+" "+", ".join(count_greater_than_5) + " has violated the rules of parking."
    else:
        violation_text = "No student has violated the rules of parking."

    violation_paragraph = Paragraph(violation_text, bold_style)
    content.append(violation_paragraph)

    # Build PDF document
    doc.build(content)

if __name__ == '__main__':
    app.run(debug=True, host= '0.0.0.0')
