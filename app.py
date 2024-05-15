from flask import Flask, request, send_file, jsonify
from pathlib import Path
import os
import pandas as pd
from docxtpl import DocxTemplate
from docx2pdf import convert
import tempfile
import pythoncom
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from email.mime.text import MIMEText
from dotenv import load_dotenv

load_dotenv()  # Load environment variables from .env file

app = Flask(__name__)

@app.route('/')
def index():
    return send_file('index.html')  # Ensure index.html is in the same directory as this script

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return jsonify(success=False), 400

    file = request.files['file']
    if file.filename == '':
        return jsonify(success=False), 400

    # SMTP configuration
    smtp_server = os.getenv('SMTP_SERVER')
    smtp_port = int(os.getenv('SMTP_PORT'))
    smtp_username = os.getenv('SMTP_USERNAME')
    smtp_password = os.getenv('SMTP_PASSWORD')

    # Create a temporary directory
    with tempfile.TemporaryDirectory() as temp_dir:
        temp_path = Path(temp_dir)
        word_template_path = Path('student.docx')
        excel_path = temp_path / file.filename
        file.save(excel_path)

        output_dir = temp_path / "DOC"
        pdf_output_dir = temp_path / "PDF"

        # Create output folders for the Word and PDF documents
        output_dir.mkdir(exist_ok=True)
        pdf_output_dir.mkdir(exist_ok=True)

        # Convert Excel sheet to pandas dataframe
        df = pd.read_excel(excel_path, sheet_name="Sheet1")

        # Iterate over each row in df and render Word document
        for record in df.to_dict(orient="records"):
            doc = DocxTemplate(word_template_path)
            doc.render(record)
            output_path = output_dir / f"{record['Name']}.docx"
            doc.save(output_path)

        # Convert generated Word documents to PDF
        convert_docx_to_pdf(output_dir, pdf_output_dir)

        # Send emails with the PDFs attached
        for record in df.to_dict(orient="records"):
            pdf_path = pdf_output_dir / f"{record['Name']}.pdf"
            send_email(smtp_server, smtp_port, smtp_username, smtp_password, record['email'], pdf_path)

    return jsonify(success=True)

def send_email(smtp_server, smtp_port, smtp_username, smtp_password, recipient_email, pdf_path):
    try:
        msg = MIMEMultipart()
        msg['From'] = smtp_username
        msg['To'] = recipient_email
        msg['Subject'] = 'Your PDF Document'

        body = 'Here is your Certificate.'
        msg.attach(MIMEText(body, 'plain'))

        with open(pdf_path, 'rb') as f:
            attach = MIMEApplication(f.read(), _subtype="pdf")
            attach.add_header('Content-Disposition', 'attachment', filename=str(pdf_path.name))
            msg.attach(attach)

        server = smtplib.SMTP(smtp_server, smtp_port)
        server.starttls()
        server.login(smtp_username, smtp_password)
        server.sendmail(smtp_username, recipient_email, msg.as_string())
        server.quit()

        print(f"Email sent to {recipient_email}")
    except Exception as e:
        print(f"Failed to send email to {recipient_email}. Error: {e}")

def convert_docx_to_pdf(docx_folder, pdf_folder):
    docx_files = [f for f in os.listdir(docx_folder) if f.endswith(".docx")]

    for docx_file in docx_files:
        docx_path = os.path.join(docx_folder, docx_file)
        pdf_name, _ = os.path.splitext(docx_file)
        pdf_path = os.path.join(pdf_folder, f"{pdf_name}.pdf")

        try:
            # Ensure COM is initialized for the conversion process
            pythoncom.CoInitialize()
            convert(docx_path, pdf_path)
            pythoncom.CoUninitialize()
        except Exception as e:
            print(f"Failed to convert {docx_file} to PDF. Error: {e}")

if __name__ == '__main__':
    app.run(debug=True)