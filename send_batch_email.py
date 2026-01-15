#!/usr/bin/env python3
"""
Email Sender for Batch Processing Results
Sends the batch processing results via Gmail SMTP.

Usage:
    python send_batch_email.py --batch-dir ./batch_output/20260114_120000 --gmail-user your.email@gmail.com --gmail-app-password YOUR_APP_PASSWORD
"""

import argparse
import smtplib
import json
from pathlib import Path
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from datetime import datetime

def send_email_with_attachments(
    sender_email,
    sender_password,
    recipient_email,
    subject,
    body,
    attachments
):
    """Send email with attachments via Gmail SMTP"""

    # Create message
    msg = MIMEMultipart()
    msg['From'] = sender_email
    msg['To'] = recipient_email
    msg['Subject'] = subject

    # Add body
    msg.attach(MIMEText(body, 'plain', 'utf-8'))

    # Add attachments
    for file_path in attachments:
        if not file_path.exists():
            print(f"Warning: Attachment not found: {file_path}")
            continue

        with open(file_path, 'rb') as f:
            part = MIMEBase('application', 'octet-stream')
            part.set_payload(f.read())

        encoders.encode_base64(part)
        part.add_header(
            'Content-Disposition',
            f'attachment; filename= {file_path.name}'
        )
        msg.attach(part)

    # Send email
    print(f"Connecting to Gmail SMTP...")
    with smtplib.SMTP_SSL('smtp.gmail.com', 465) as server:
        print(f"Logging in as {sender_email}...")
        server.login(sender_email, sender_password)
        print(f"Sending email to {recipient_email}...")
        server.send_message(msg)
        print(f"Email sent successfully!")

def main():
    parser = argparse.ArgumentParser(
        description="Send batch processing results via email"
    )
    parser.add_argument(
        "--batch-dir",
        type=Path,
        required=True,
        help="Path to batch output directory"
    )
    parser.add_argument(
        "--gmail-user",
        required=True,
        help="Gmail email address"
    )
    parser.add_argument(
        "--gmail-app-password",
        required=True,
        help="Gmail App Password (not your regular password!)"
    )
    parser.add_argument(
        "--recipient",
        default="elay.g@82labs.io",
        help="Email recipient (default: elay.g@82labs.io)"
    )

    args = parser.parse_args()

    # Validate batch directory
    if not args.batch_dir.exists():
        print(f"Error: Batch directory not found: {args.batch_dir}")
        return 1

    # Load batch summary
    summary_file = args.batch_dir / "batch_summary.txt"
    json_summary_file = args.batch_dir / "batch_summary.json"

    if not summary_file.exists():
        print(f"Error: batch_summary.txt not found in {args.batch_dir}")
        return 1

    # Read email body
    email_body = summary_file.read_text(encoding='utf-8')

    # Collect all attachments
    attachments = []

    # Add summary files
    attachments.append(summary_file)
    if json_summary_file.exists():
        attachments.append(json_summary_file)

    # Add all XLSX and JSON files from manager directories
    for manager_dir in args.batch_dir.iterdir():
        if manager_dir.is_dir():
            for file in manager_dir.glob("*.xlsx"):
                attachments.append(file)
            for file in manager_dir.glob("*.json"):
                attachments.append(file)

    # Send email
    subject = f"Mizrahi Special Transactions - Batch Results ({datetime.now().strftime('%Y-%m-%d')})"

    print("="*60)
    print("SENDING BATCH RESULTS EMAIL")
    print("="*60)
    print(f"From: {args.gmail_user}")
    print(f"To: {args.recipient}")
    print(f"Subject: {subject}")
    print(f"Attachments: {len(attachments)}")
    for att in attachments:
        print(f"  - {att.name}")
    print("="*60)

    try:
        send_email_with_attachments(
            args.gmail_user,
            args.gmail_app_password,
            args.recipient,
            subject,
            email_body,
            attachments
        )
        print("\n" + "="*60)
        print("SUCCESS: Email sent successfully!")
        print("="*60)
        return 0
    except Exception as e:
        print(f"\nERROR: Failed to send email: {e}")
        print("\nTroubleshooting:")
        print("1. Make sure you're using a Gmail App Password (not your regular password)")
        print("2. Enable 2-factor authentication on your Google account")
        print("3. Generate an App Password at: https://myaccount.google.com/apppasswords")
        return 1

if __name__ == "__main__":
    import sys
    sys.exit(main())
