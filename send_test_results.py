#!/usr/bin/env python3
"""
Send test results for one manager via email
"""

import sys
import smtplib
from pathlib import Path
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import json

def send_email(gmail_user, gmail_password, recipient, test_dir):
    """Send test results email"""

    # Read summary
    summary_file = Path(test_dir) / "test_summary.json"
    with open(summary_file) as f:
        summary = json.load(f)

    # Create email
    msg = MIMEMultipart()
    msg['From'] = gmail_user
    msg['To'] = recipient
    msg['Subject'] = f"Mizrahi Special Transactions Test - {summary['results'][0]['manager_name']}"

    # Email body
    body = f"""
Test Results for Special Transactions Actor (nQh62mdhpUTM5l65l)
{'='*60}

Manager: {summary['results'][0]['manager_name']}
Fund Code: {summary['results'][0]['fund_code']}
Status: {summary['results'][0]['status']}
Test Date: {summary['test_date']}

Actor Details:
- Actor ID: {summary['actor_id']}
- Build ID: {summary['build_id']}
- Run ID: {summary['results'][0]['run_id']}
- KV Store ID: {summary['results'][0]['kv_store_id']}

Files Retrieved:
"""

    # Attach files
    attachments = []
    for file_path_str in summary['results'][0]['files_saved']:
        file_path = Path(file_path_str)
        if file_path.exists():
            body += f"- {file_path.name} ({file_path.stat().st_size:,} bytes)\n"
            attachments.append(file_path)

    body += f"\n{'='*60}\n"
    body += f"All files are attached to this email.\n"

    msg.attach(MIMEText(body, 'plain', 'utf-8'))

    # Attach files
    for file_path in attachments:
        with open(file_path, 'rb') as f:
            part = MIMEBase('application', 'octet-stream')
            part.set_payload(f.read())

        encoders.encode_base64(part)
        part.add_header('Content-Disposition', f'attachment; filename= {file_path.name}')
        msg.attach(part)

    # Also attach summary JSON
    with open(summary_file, 'rb') as f:
        part = MIMEBase('application', 'json')
        part.set_payload(f.read())
    encoders.encode_base64(part)
    part.add_header('Content-Disposition', f'attachment; filename= test_summary.json')
    msg.attach(part)

    # Send email
    print(f"Sending email to {recipient}...")
    with smtplib.SMTP_SSL('smtp.gmail.com', 465) as server:
        server.login(gmail_user, gmail_password)
        server.send_message(msg)
    print("Email sent successfully!")

if __name__ == "__main__":
    if len(sys.argv) != 5:
        print("Usage: python send_test_results.py <gmail_user> <gmail_password> <recipient> <test_dir>")
        sys.exit(1)

    send_email(sys.argv[1], sys.argv[2], sys.argv[3], sys.argv[4])
