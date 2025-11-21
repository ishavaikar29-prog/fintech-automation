import os
import requests
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders
from openpyxl import Workbook


def fetch_api_data(api_url, api_key):
    """Fetch data from API."""
    headers = {"Authorization": f"Bearer {api_key}"}
    res = requests.get(api_url, headers=headers, timeout=20)

    if res.status_code != 200:
        raise Exception(f"API Error: {res.status_code} → {res.text}")

    return res.json()


def create_excel(data, file_path="report.xlsx"):
    """Create an Excel file from a list/dict."""
    wb = Workbook()
    sheet = wb.active
    sheet.title = "Report"

    # If data is a list of dictionaries
    if isinstance(data, list) and len(data) > 0 and isinstance(data[0], dict):
        headers = list(data[0].keys())
        sheet.append(headers)

        for row in data:
            sheet.append(list(row.values()))

    else:
        sheet.append(["Data"])
        sheet.append([str(data)])

    wb.save(file_path)
    return file_path


def send_email_with_attachment(
    smtp_host,
    smtp_port,
    smtp_user,
    smtp_pass,
    to_email,
    subject,
    body,
    attachment_path
):
    """Send email with attachment."""
    msg = MIMEMultipart()
    msg["From"] = smtp_user
    msg["To"] = to_email
    msg["Subject"] = subject

    msg.attach(MIMEText(body, "plain"))

    # File attachment
    part = MIMEBase("application", "octet-stream")
    with open(attachment_path, "rb") as f:
        part.set_payload(f.read())
    encoders.encode_base64(part)
    part.add_header("Content-Disposition", f"attachment; filename={attachment_path}")

    msg.attach(part)

    # SMTP connection
    server = smtplib.SMTP(smtp_host, smtp_port)
    server.starttls()
    server.login(smtp_user, smtp_pass)
    server.send_message(msg)
    server.quit()

    print("Email sent successfully!")


if __name__ == "__main__":
    # Read from GitHub Secrets (env variables)
    API_URL = os.getenv("API_URL")
    API_KEY = os.getenv("API_KEY")

    SMTP_HOST = os.getenv("SMTP_HOST")
    SMTP_PORT = int(os.getenv("SMTP_PORT", "587"))
    SMTP_USER = os.getenv("SMTP_USER")
    SMTP_PASS = os.getenv("SMTP_PASS")
    TO_EMAIL = os.getenv("TO_EMAIL")

    # 1. Fetch data
    print("Fetching API data...")
    data = fetch_api_data(API_URL, API_KEY)

    # 2. Generate Excel
    print("Generating Excel...")
    excel_file = create_excel(data)

    # 3. Send Email
    print("Sending email...")
    send_email_with_attachment(
        smtp_host=SMTP_HOST,
        smtp_port=SMTP_PORT,
        smtp_user=SMTP_USER,
        smtp_pass=SMTP_PASS,
        to_email=TO_EMAIL,
        subject="Daily Automated Report",
        body="Hello,\n\nPlease find today's automated report attached.\n\nRegards,\nAutomation Bot",
        attachment_path=excel_file
    )
