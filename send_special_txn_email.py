#!/usr/bin/env python3
"""
Email Sender for Special Transactions Reports
Sends two emails per manager using Resend API:
1. Report email with Excel attachment
2. Transaction samples email (conditional on samples existing)

Usage:
    python send_special_txn_email.py \
        --manager-name "מגדל" \
        --xlsx-path "./output/מגדל_special_transactions_report.xlsx" \
        --email-json-path "./output/מגדל_email.json" \
        --report-month "דצמבר 2025" \
        --recipient "manager@example.com"
"""

import argparse
import json
import os
import sys
import io
from pathlib import Path

# Ensure UTF-8 encoding for Hebrew text
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

# Try to import resend - only required for actual sending, not test mode
try:
    import resend
    RESEND_AVAILABLE = True
except ImportError:
    RESEND_AVAILABLE = False
    resend = None  # Set to None so we can check later


def send_report_email(manager_name, xlsx_path, report_month, recipient):
    """
    Email 1: Rich HTML report with Excel attachment

    Args:
        manager_name: Manager name (e.g., "מגדל")
        xlsx_path: Path to Excel report file
        report_month: Report month in Hebrew (e.g., "דצמבר 2025")
        recipient: Email recipient address

    Returns:
        dict: Resend API response with email ID
    """
    html_body = build_report_html(manager_name, xlsx_path.name, report_month)

    # Read Excel file as bytes
    with open(xlsx_path, 'rb') as f:
        xlsx_content = f.read()

    params = {
        "from": "Mizrahi Special Transactions <noreply@82labs.io>",
        "to": [recipient],
        "subject": f"דוח עסקאות מיוחדות - {manager_name}",
        "html": html_body,
        "attachments": [
            {
                "filename": xlsx_path.name,
                "content": list(xlsx_content)  # Resend expects list of bytes
            }
        ]
    }

    return resend.Emails.send(params)


def send_samples_email(email_json_path, recipient):
    """
    Email 2: Transaction samples (conditional on content)

    Args:
        email_json_path: Path to email.json file
        recipient: Email recipient address

    Returns:
        dict: Resend API response with email ID, or None if no samples
    """
    # Read email.json
    with open(email_json_path, encoding="utf-8") as f:
        data = json.load(f)

    # Return early if no samples
    if not data:
        return None

    email_data = data[0]
    decision_1 = email_data.get("decision_1_transactions")
    decision_2 = email_data.get("decision_2_transactions")

    if not decision_1 and not decision_2:
        return None

    # Build Hebrew text email with tables
    html_body = build_samples_html(email_data)

    params = {
        "from": "Mizrahi Special Transactions <noreply@82labs.io>",
        "to": [recipient],
        "subject": f"בדיקת עסקאות מיוחדות - {email_data['manager_name']} - {email_data['report_month']}",
        "html": html_body
    }

    return resend.Emails.send(params)


def build_report_html(manager_name, filename, report_month):
    """Generate rich HTML for Email 1 (based on n8n template)"""
    return f"""<!DOCTYPE html>
<html dir="rtl" lang="he">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>דוח עסקאות מיוחדות</title>
</head>
<body style="margin: 0; padding: 0; font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; background-color: #f4f6f8; direction: rtl; line-height: 1.7;">
    <table cellpadding="0" cellspacing="0" border="0" width="100%" style="background-color: #f4f6f8; padding: 40px 20px;">
        <tr>
            <td align="center">
                <table cellpadding="0" cellspacing="0" border="0" width="600" style="max-width: 600px; background-color: #ffffff; border-radius: 12px; box-shadow: 0 4px 20px rgba(0, 0, 0, 0.06); overflow: hidden;">

                    <!-- Header -->
                    <tr>
                        <td style="background-color: #ffffff; padding: 30px 40px; border-bottom: 1px solid #eef1f4; text-align: center;">
                            <h1 style="color: #1e293b; margin: 0; font-size: 26px; font-weight: 600;">דוח עסקאות מיוחדות</h1>
                        </td>
                    </tr>

                    <!-- Status Banner -->
                    <tr>
                        <td style="background: linear-gradient(135deg, #3B82F6 0%, #1D4ED8 100%); padding: 45px 40px; text-align: center;">
                            <div style="width: 70px; height: 70px; background-color: rgba(255,255,255,0.2); border-radius: 50%; margin: 0 auto 20px; line-height: 70px;">
                                <span style="font-size: 32px; color: white;">📊</span>
                            </div>
                            <h2 style="color: #ffffff; margin: 0 0 8px 0; font-size: 24px; font-weight: 600;">בדיקות #1-#7 הושלמו</h2>
                            <p style="color: rgba(255,255,255,0.9); margin: 0; font-size: 15px; font-weight: 400;">{manager_name}</p>
                        </td>
                    </tr>

                    <!-- Main Content -->
                    <tr>
                        <td style="padding: 40px;">

                            <!-- Info Card -->
                            <table cellpadding="0" cellspacing="0" border="0" width="100%" style="background-color: #f8fafc; border-radius: 8px; border-right: 4px solid #3B82F6;">
                                <tr>
                                    <td style="padding: 25px;">
                                        <h3 style="color: #1e293b; margin: 0 0 15px 0; font-size: 17px; font-weight: 600;">פרטי הדוח</h3>

                                        <table cellpadding="0" cellspacing="0" border="0" width="100%">
                                            <tr>
                                                <td style="padding: 8px 0; color: #64748b; font-size: 14px; width: 120px;">מנהל קרן:</td>
                                                <td style="padding: 8px 0; color: #1e293b; font-size: 14px; font-weight: 500;">{manager_name}</td>
                                            </tr>
                                            <tr>
                                                <td style="padding: 8px 0; color: #64748b; font-size: 14px;">שם הדוח:</td>
                                                <td style="padding: 8px 0; color: #1e293b; font-size: 14px; font-weight: 500;">{filename}</td>
                                            </tr>
                                            <tr>
                                                <td style="padding: 8px 0; color: #64748b; font-size: 14px;">חודש דיווח:</td>
                                                <td style="padding: 8px 0; color: #1e293b; font-size: 14px; font-weight: 500;">{report_month}</td>
                                            </tr>
                                        </table>
                                    </td>
                                </tr>
                            </table>

                            <!-- Checks Summary -->
                            <table cellpadding="0" cellspacing="0" border="0" width="100%" style="margin-top: 25px;">
                                <tr>
                                    <td style="background-color: #f0fdf4; border-radius: 8px; padding: 18px 20px;">
                                        <p style="color: #166534; font-size: 14px; margin: 0 0 10px 0; font-weight: 500;">בדיקות שבוצעו:</p>
                                        <ul style="color: #166534; font-size: 13px; margin: 10px 0 0 0; padding-right: 20px;">
                                            <li>בדיקה #1: עסקאות בין-קרנות</li>
                                            <li>בדיקה #3: תאריכים בטווח חודש הדיווח</li>
                                            <li>בדיקה #4ד: שלמות נתונים בסיסית</li>
                                            <li>בדיקה #4א: התאמת סוג עסקה לאופן החלטה</li>
                                            <li>בדיקה #4ב: דח"צ - התנגדות להחלטה</li>
                                            <li>בדיקה #4ג: דח"צ 1 לא קיבל החלטה</li>
                                            <li>בדיקה #5א: דגימה - אופן החלטה 1</li>
                                            <li>בדיקה #5ב: דגימה - אופן החלטה 2</li>
                                            <li>בדיקה #6: השוואת מחירים לבורסה</li>
                                            <li>בדיקה #7: ניירות ערך בעייתיים</li>
                                        </ul>
                                    </td>
                                </tr>
                            </table>

                            <!-- Attachment Notice -->
                            <table cellpadding="0" cellspacing="0" border="0" width="100%" style="margin-top: 25px;">
                                <tr>
                                    <td style="background-color: #eff6ff; border-radius: 8px; padding: 18px 20px; text-align: center;">
                                        <span style="color: #3b82f6; font-size: 18px; margin-left: 8px;">📎</span>
                                        <span style="color: #1e40af; font-size: 14px;">הדוח המלא מצורף למייל זה כקובץ Excel</span>
                                    </td>
                                </tr>
                            </table>

                        </td>
                    </tr>

                    <!-- Footer -->
                    <tr>
                        <td style="background-color: #f8fafc; padding: 20px 40px; text-align: center; border-top: 1px solid #e2e8f0;">
                            <p style="margin: 0; color: #94a3b8; font-size: 12px;">
                                Powered by <a href="https://82labs.com" target="_blank" style="color: #64748b; text-decoration: none; font-weight: 600;">82Labs</a>
                            </p>
                        </td>
                    </tr>

                </table>
            </td>
        </tr>
    </table>
</body>
</html>"""


def build_samples_html(email_data):
    """Generate Hebrew email with transaction tables for Email 2"""
    manager_name = email_data["manager_name"]
    report_month = email_data["report_month"]
    decision_1 = email_data.get("decision_1_transactions")
    decision_2 = email_data.get("decision_2_transactions")

    # Build tables
    decision_1_section = ""
    if decision_1:
        decision_1_table = _build_txn_table(decision_1)
        decision_1_section = f"""
        <h3 style="color: #1e293b; font-size: 16px; margin: 25px 0 10px 0;">אופן החלטה 1:</h3>
        {decision_1_table}
        <p style="margin: 15px 0 0 0; color: #1e293b; font-size: 14px;">
            אנא ציינו בהתאם לאיזה סעיף בנוהל אושרה העסקה וצרפו אסמכתאות רלוונטיות בהתאם לנדרש בהתאם לנוהל
        </p>
        """

    decision_2_section = ""
    if decision_2:
        decision_2_table = _build_txn_table(decision_2)
        decision_2_section = f"""
        <h3 style="color: #1e293b; font-size: 16px; margin: 25px 0 10px 0;">אופן החלטה 2:</h3>
        {decision_2_table}
        <p style="margin: 15px 0 0 0; color: #1e293b; font-size: 14px;">
            אנא ציינו את מועד אישור העסקה וכן צרפו פרוטוקול/ טיוטת פרוטוקול/ חומר נלווה שהוצג בפני הדירקטוריון לטובת אישור העסקה.
        </p>
        """

    return f"""<!DOCTYPE html>
<html dir="rtl" lang="he">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>בדיקת עסקאות מיוחדות</title>
</head>
<body style="margin: 0; padding: 0; font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; background-color: #f4f6f8; direction: rtl; line-height: 1.7;">
    <table cellpadding="0" cellspacing="0" border="0" width="100%" style="background-color: #f4f6f8; padding: 40px 20px;">
        <tr>
            <td align="center">
                <table cellpadding="0" cellspacing="0" border="0" width="800" style="max-width: 800px; background-color: #ffffff; border-radius: 8px; padding: 40px;">
                    <tr>
                        <td>
                            <p style="margin: 0 0 15px 0; color: #1e293b; font-size: 14px;">שלום רב,</p>

                            <p style="margin: 0 0 20px 0; color: #1e293b; font-size: 14px;">
                                בהתאם לבדיקת עסקאות מיוחדות לחודש {report_month}, עלו לנו עסקאות שעליהן נבקש לקבל פרטים נוספים:
                            </p>

                            {decision_1_section}

                            {decision_2_section}

                            <p style="margin: 25px 0 15px 0; color: #1e293b; font-size: 14px;">
                                נשמח לקבל את התייחסותכם בהקדם האפשרי.
                            </p>

                            <p style="margin: 15px 0 0 0; color: #1e293b; font-size: 14px;">
                                בברכה,<br>
                                צוות מזרחי טפחות החברה לנאמנות
                            </p>
                        </td>
                    </tr>
                </table>
            </td>
        </tr>
    </table>
</body>
</html>"""


def _build_txn_table(transactions):
    """Build HTML table from transaction list"""
    if not transactions:
        return ""

    rows = []
    for txn in transactions:
        rows.append(f"""
        <tr>
            <td style="border: 1px solid #ddd; padding: 10px; text-align: center; font-size: 13px; color: #1e293b;">{txn['fund_number']}</td>
            <td style="border: 1px solid #ddd; padding: 10px; text-align: right; font-size: 13px; color: #1e293b;">{txn['fund_name']}</td>
            <td style="border: 1px solid #ddd; padding: 10px; text-align: right; font-size: 13px; color: #1e293b;">{txn['security_name']}</td>
            <td style="border: 1px solid #ddd; padding: 10px; text-align: center; font-size: 13px; color: #1e293b;">{txn['security_number']}</td>
            <td style="border: 1px solid #ddd; padding: 10px; text-align: center; font-size: 13px; color: #1e293b;">{txn['quantity']}</td>
            <td style="border: 1px solid #ddd; padding: 10px; text-align: center; font-size: 13px; color: #1e293b;">{txn['price']}</td>
            <td style="border: 1px solid #ddd; padding: 10px; text-align: center; font-size: 13px; color: #1e293b;">{txn['date']}</td>
            <td style="border: 1px solid #ddd; padding: 10px; text-align: center; font-size: 13px; color: #1e293b;">{txn['type']}</td>
        </tr>""")

    return f"""
    <table cellpadding="0" cellspacing="0" border="0" width="100%" style="border-collapse: collapse; margin: 15px 0;">
        <thead>
            <tr style="background-color: #f0f0f0;">
                <th style="border: 1px solid #ddd; padding: 12px; font-size: 13px; font-weight: 600; color: #1e293b;">מספר קרן</th>
                <th style="border: 1px solid #ddd; padding: 12px; font-size: 13px; font-weight: 600; color: #1e293b;">שם קרן</th>
                <th style="border: 1px solid #ddd; padding: 12px; font-size: 13px; font-weight: 600; color: #1e293b;">שם נייר</th>
                <th style="border: 1px solid #ddd; padding: 12px; font-size: 13px; font-weight: 600; color: #1e293b;">מספר נייר</th>
                <th style="border: 1px solid #ddd; padding: 12px; font-size: 13px; font-weight: 600; color: #1e293b;">כמות</th>
                <th style="border: 1px solid #ddd; padding: 12px; font-size: 13px; font-weight: 600; color: #1e293b;">מחיר</th>
                <th style="border: 1px solid #ddd; padding: 12px; font-size: 13px; font-weight: 600; color: #1e293b;">תאריך</th>
                <th style="border: 1px solid #ddd; padding: 12px; font-size: 13px; font-weight: 600; color: #1e293b;">סוג</th>
            </tr>
        </thead>
        <tbody>
            {''.join(rows)}
        </tbody>
    </table>"""


def save_report_html(manager_name, xlsx_path, report_month, output_dir):
    """Save report email HTML to file instead of sending"""
    html_body = build_report_html(manager_name, xlsx_path.name, report_month)

    output_file = output_dir / f"{manager_name}_email_1_report.html"
    output_file.write_text(html_body, encoding='utf-8')

    return output_file


def save_samples_html(email_json_path, output_dir, manager_name):
    """Save samples email HTML to file instead of sending"""
    # Read email.json
    with open(email_json_path, encoding="utf-8") as f:
        data = json.load(f)

    # Return early if no samples
    if not data:
        return None

    email_data = data[0]
    decision_1 = email_data.get("decision_1_transactions")
    decision_2 = email_data.get("decision_2_transactions")

    if not decision_1 and not decision_2:
        return None

    # Build Hebrew text email with tables
    html_body = build_samples_html(email_data)

    output_file = output_dir / f"{manager_name}_email_2_samples.html"
    output_file.write_text(html_body, encoding='utf-8')

    return output_file


def main():
    parser = argparse.ArgumentParser(
        description="Send special transactions emails via Resend (or save as HTML in test mode)"
    )
    parser.add_argument(
        "--manager-name",
        required=True,
        help="Manager name (e.g., מגדל)"
    )
    parser.add_argument(
        "--xlsx-path",
        required=True,
        type=Path,
        help="Path to Excel report"
    )
    parser.add_argument(
        "--email-json-path",
        required=True,
        type=Path,
        help="Path to email.json"
    )
    parser.add_argument(
        "--report-month",
        required=True,
        help="Report month in Hebrew (e.g., דצמבר 2025)"
    )
    parser.add_argument(
        "--recipient",
        required=True,
        help="Email recipient"
    )
    parser.add_argument(
        "--resend-api-key",
        help="Resend API key (or use RESEND_API_KEY env var)"
    )
    parser.add_argument(
        "--test-mode",
        action="store_true",
        help="Test mode: save emails as HTML files instead of sending"
    )
    parser.add_argument(
        "--output-dir",
        type=Path,
        help="Output directory for HTML files in test mode (defaults to xlsx-path parent dir)"
    )

    args = parser.parse_args()

    # Determine if we're in test mode
    test_mode = args.test_mode or not (args.resend_api_key or os.getenv("RESEND_API_KEY"))

    # If not in test mode, check if resend is available
    if not test_mode and not RESEND_AVAILABLE:
        print("="*60)
        print("ERROR: Resend package not installed")
        print("="*60)
        print("The resend package is required for sending emails.")
        print("\nOptions:")
        print("  1. Install resend: pip install resend")
        print("  2. Use test mode: add --test-mode flag to generate HTML previews")
        print("\nTest mode example:")
        print(f"  python {Path(__file__).name} --test-mode ...")
        return 1

    if test_mode:
        print("="*60)
        print("TEST MODE: Saving emails as HTML files")
        print("="*60)
        print(f"Reason: {'--test-mode flag set' if args.test_mode else 'RESEND_API_KEY not configured'}")
    else:
        # Set API key for sending
        if args.resend_api_key:
            resend.api_key = args.resend_api_key
        else:
            resend.api_key = os.getenv("RESEND_API_KEY")

    # Validate files exist
    if not args.xlsx_path.exists():
        print(f"Error: Excel file not found: {args.xlsx_path}")
        return 1

    if not args.email_json_path.exists():
        print(f"Error: Email JSON file not found: {args.email_json_path}")
        return 1

    # Determine output directory for test mode
    output_dir = args.output_dir if args.output_dir else args.xlsx_path.parent

    if not test_mode:
        print("="*60)
        print("SENDING SPECIAL TRANSACTIONS EMAILS")
        print("="*60)

    print(f"Manager: {args.manager_name}")
    print(f"Report Month: {args.report_month}")
    print(f"Recipient: {args.recipient}")
    print(f"Excel File: {args.xlsx_path.name}")
    if test_mode:
        print(f"Output Directory: {output_dir}")
    print("="*60)

    try:
        if test_mode:
            # TEST MODE: Save HTML files
            print(f"\n[1/2] Saving report email as HTML...")
            html_file_1 = save_report_html(
                args.manager_name,
                args.xlsx_path,
                args.report_month,
                output_dir
            )
            print(f"✓ Report email saved: {html_file_1.name}")

            print(f"\n[2/2] Checking for transaction samples...")
            html_file_2 = save_samples_html(
                args.email_json_path,
                output_dir,
                args.manager_name
            )

            if html_file_2:
                print(f"✓ Samples email saved: {html_file_2.name}")
            else:
                print("⊘ No samples to save (no decision_1 or decision_2 transactions)")

            print("\n" + "="*60)
            print("SUCCESS: Email HTML files created!")
            print("="*60)
            print(f"\nHTML files saved to: {output_dir}")
            if html_file_2:
                print(f"  - {html_file_1.name}")
                print(f"  - {html_file_2.name}")
            else:
                print(f"  - {html_file_1.name}")
            print("\nOpen these files in a browser to preview the emails.")
            return 0

        else:
            # NORMAL MODE: Send emails via Resend
            print(f"\n[1/2] Sending report email...")
            result1 = send_report_email(
                args.manager_name,
                args.xlsx_path,
                args.report_month,
                args.recipient
            )
            print(f"✓ Report email sent successfully")
            print(f"  Email ID: {result1['id']}")

            print(f"\n[2/2] Checking for transaction samples...")
            result2 = send_samples_email(args.email_json_path, args.recipient)

            if result2:
                print(f"✓ Samples email sent successfully")
                print(f"  Email ID: {result2['id']}")
            else:
                print("⊘ No samples to send (no decision_1 or decision_2 transactions)")

            print("\n" + "="*60)
            print("SUCCESS: Email process completed!")
            print("="*60)
            return 0

    except Exception as e:
        if test_mode:
            print(f"\nERROR: Failed to create HTML files: {e}")
            return 1
        else:
            print(f"\nERROR: Failed to send emails: {e}")
            print("\nTroubleshooting:")
            print("1. Verify RESEND_API_KEY is correct")
            print("2. Check that you have a verified sender domain in Resend")
            print("3. Ensure the 'from' email domain matches your Resend configuration")
            print("4. Visit https://resend.com/docs for more information")
            return 1


if __name__ == "__main__":
    sys.exit(main())
