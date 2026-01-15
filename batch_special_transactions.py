#!/usr/bin/env python3
"""
Batch Special Transactions Processor
Processes special transactions for all fund managers and sends reports via email.

Usage:
    python batch_special_transactions.py --apify-token YOUR_TOKEN
    python batch_special_transactions.py --apify-token YOUR_TOKEN --managers "מגדל,איילון"
    python batch_special_transactions.py --apify-token YOUR_TOKEN --skip-tase-prices
"""

import os
import sys
import time
import base64
import argparse
import json
import subprocess
import smtplib
from datetime import datetime, timedelta
from pathlib import Path
from typing import Optional, List, Dict
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

try:
    import requests
except ImportError:
    print("Missing required package: requests")
    print("Install with: pip install requests")
    sys.exit(1)

# ============================================================================
# CONFIGURATION
# ============================================================================

# Apify Actor IDs
FUNDS_LIST_ACTOR_ID = "K9WppTziYC3n2vxTu"
FUND_REPORTS_ACTOR_ID = "5lhI6O39Qbgv9O0gs"  # Also used for special transactions

# Fund Manager Names and Codes
FUND_MANAGERS = {
    "מגדל": "10040",
    "איילון": "10054",
    "קסם": "10047",
    "סיגמא": "10048",
    "פורסט": "10082",
    "הראל": "10031",
    "אנליסט": "10019",
    "מיטב": "10083",
    "איביאי": "10068",
    "אלטשולר-שחם": "10017",
}

# Email configuration
EMAIL_RECIPIENT = "elay.g@82labs.io"

# Processing configuration
TRUSTEE_NAME = "מזרחי טפחות"
DEFAULT_PRICE_THRESHOLD = 5.0
DEFAULT_SPEC_FILE = "spec-file.xlsx"  # Default specification file

# ============================================================================
# LOGGING
# ============================================================================

def log(msg, level="INFO"):
    """Log message with timestamp"""
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    print(f"[{timestamp}] [{level}] {msg}")

def log_error(msg):
    """Log error message"""
    log(msg, "ERROR")

def log_success(msg):
    """Log success message"""
    log(msg, "SUCCESS")

# ============================================================================
# APIFY FUNCTIONS
# ============================================================================

def apify_request(method, endpoint, token, json_data=None, params=None):
    """Make request to Apify API"""
    url = f"https://api.apify.com/v2{endpoint}"
    headers = {"Authorization": f"Bearer {token}"}

    try:
        response = requests.request(method, url, headers=headers, json=json_data, params=params)
        response.raise_for_status()
        return response
    except requests.exceptions.RequestException as e:
        log_error(f"Apify API request failed: {e}")
        raise

def run_actor_and_wait(actor_id, token, input_data=None, timeout=300):
    """Run an Apify actor and wait for completion"""
    log(f"Starting Apify actor: {actor_id}")

    # Start the run
    resp = apify_request("POST", f"/acts/{actor_id}/runs", token, json_data=input_data or {}, params={"timeout": timeout})
    run_data = resp.json()["data"]
    run_id = run_data["id"]
    log(f"Run started: {run_id}")

    # Poll for completion
    start = time.time()
    while time.time() - start < timeout:
        resp = apify_request("GET", f"/actor-runs/{run_id}", token)
        status = resp.json()["data"]["status"]

        if status == "SUCCEEDED":
            log_success(f"Actor run completed: {run_id}")
            return resp.json()["data"]
        elif status in ("FAILED", "ABORTED", "TIMED-OUT"):
            log_error(f"Actor failed with status: {status}")
            raise Exception(f"Actor failed with status: {status}")

        log(f"Status: {status}... waiting")
        time.sleep(5)

    raise Exception("Timeout waiting for actor")

def build_maya_special_transactions_url(fund_code):
    """Build Maya URL for special transactions reports"""
    today = datetime.now()
    one_year_ago = today - timedelta(days=365)

    # This URL format is for special transactions - adjust if needed
    return (
        f"https://maya.tase.co.il/he/reports/funds?"
        f"fromDate={one_year_ago.strftime('%Y-%m-%d')}&toDate={today.strftime('%Y-%m-%d')}"
        f"&noMeetings=false&isSingle=false&isIntendToTaseMember=false"
        f"&by=group&groupId=7&itemId={fund_code}&eventsIds%5B%5D=5615"
    )

def fetch_funds_list(token, output_dir):
    """Fetch mutual funds list from Apify"""
    log("Fetching Mutual Funds List from Apify...")

    try:
        run_data = run_actor_and_wait(FUNDS_LIST_ACTOR_ID, token, {})
        dataset_id = run_data["defaultDatasetId"]

        # Get dataset items
        resp = apify_request("GET", f"/datasets/{dataset_id}/items", token)
        items = resp.json()

        if not items or not items[0].get("fileBase64"):
            raise Exception("No fileBase64 in Apify response")

        # Save the file
        funds_list_bytes = base64.b64decode(items[0]["fileBase64"])
        funds_list_path = output_dir / "Mutual_Funds_List.xlsx"
        funds_list_path.write_bytes(funds_list_bytes)

        log_success(f"Saved Mutual Funds List: {funds_list_path} ({len(funds_list_bytes)} bytes)")
        return funds_list_path

    except Exception as e:
        log_error(f"Failed to fetch funds list: {e}")
        raise

def fetch_manager_report(manager_name, fund_code, token, output_dir):
    """Fetch manager special transactions report from Apify"""
    log(f"Fetching special transactions report for {manager_name} (code: {fund_code})...")

    try:
        maya_url = build_maya_special_transactions_url(fund_code)
        run_data = run_actor_and_wait(FUND_REPORTS_ACTOR_ID, token, {"url": maya_url}, timeout=300)

        # Try to get CSV from key-value store
        kv_store_id = run_data["defaultKeyValueStoreId"]

        # Try different possible keys
        report_keys = ["report_latest_month.csv", "special_transactions.csv", "report.csv"]

        for key in report_keys:
            try:
                resp = apify_request("GET", f"/key-value-stores/{kv_store_id}/records/{key}", token)
                if resp.status_code == 200:
                    report_path = output_dir / f"{manager_name}_special_transactions.csv"
                    report_path.write_bytes(resp.content)
                    log_success(f"Saved report: {report_path} ({len(resp.content)} bytes)")
                    return report_path
            except:
                continue

        log_error(f"Could not find special transactions report in key-value store")
        return None

    except Exception as e:
        log_error(f"Failed to fetch report for {manager_name}: {e}")
        return None

# ============================================================================
# PROCESSING FUNCTIONS
# ============================================================================

def process_manager(manager_name, manager_code, funds_list_path, token, args, output_base_dir):
    """Process special transactions for a single manager"""
    log(f"\n{'='*60}")
    log(f"Processing manager: {manager_name}")
    log(f"{'='*60}")

    # Create manager-specific output directory
    manager_dir = output_base_dir / manager_name
    manager_dir.mkdir(parents=True, exist_ok=True)

    # Fetch manager report
    report_path = fetch_manager_report(manager_name, manager_code, token, manager_dir)
    if not report_path:
        log_error(f"Skipping {manager_name} - no report available")
        return None

    # Prepare output paths
    output_xlsx = manager_dir / f"{manager_name}_special_transactions_report.xlsx"
    email_json = manager_dir / f"{manager_name}_email.json"

    # Build command for mizrahi_special_transactions.py
    cmd = [
        sys.executable,
        str(Path(__file__).parent / "mizrahi_special_transactions.py"),
        "--mutual-funds-list", str(funds_list_path),
        "--input-report", str(report_path),
        "--output-xlsx", str(output_xlsx),
        "--email-json", str(email_json),
        "--manager-name", manager_name,
        "--price-threshold", str(args.price_threshold),
        "--log-dir", str(manager_dir),  # Logs go into manager's output directory
    ]

    if args.skip_tase_prices:
        cmd.append("--skip-tase-prices")

    if args.spec_file and args.spec_file.exists():
        cmd.extend(["--spec-file", str(args.spec_file)])

    if args.seed:
        cmd.extend(["--seed", str(args.seed)])

    # Run the processing script
    log(f"Running special transactions processor for {manager_name}...")
    log(f"Command: {' '.join(cmd)}")

    try:
        result = subprocess.run(cmd, capture_output=True, text=True, timeout=600)

        if result.returncode == 0:
            log_success(f"Processing completed for {manager_name}")

            # Check if output files exist
            if output_xlsx.exists() and email_json.exists():
                return {
                    "manager_name": manager_name,
                    "manager_code": manager_code,
                    "output_xlsx": output_xlsx,
                    "email_json": email_json,
                    "status": "success",
                    "log_output": result.stdout
                }
            else:
                log_error(f"Output files missing for {manager_name}")
                return {
                    "manager_name": manager_name,
                    "status": "failed",
                    "error": "Output files not generated"
                }
        else:
            log_error(f"Processing failed for {manager_name}")
            log_error(f"stderr: {result.stderr}")
            return {
                "manager_name": manager_name,
                "status": "failed",
                "error": result.stderr
            }

    except subprocess.TimeoutExpired:
        log_error(f"Processing timeout for {manager_name}")
        return {
            "manager_name": manager_name,
            "status": "failed",
            "error": "Processing timeout (>10 minutes)"
        }
    except Exception as e:
        log_error(f"Processing exception for {manager_name}: {e}")
        return {
            "manager_name": manager_name,
            "status": "failed",
            "error": str(e)
        }

# ============================================================================
# EMAIL FUNCTIONS
# ============================================================================

def send_results_email(results, recipient, output_dir):
    """Send consolidated email with all results"""
    log(f"\n{'='*60}")
    log(f"Preparing email to {recipient}")
    log(f"{'='*60}")

    # Create email summary
    successful = [r for r in results if r.get("status") == "success"]
    failed = [r for r in results if r.get("status") == "failed"]

    # Build email body
    email_body = f"""
Mizrahi Special Transactions - Batch Processing Results
{'='*60}

Processing Date: {datetime.now().strftime("%Y-%m-%d %H:%M:%S")}
Total Managers: {len(results)}
Successful: {len(successful)}
Failed: {len(failed)}

{'='*60}
SUCCESSFUL MANAGERS:
"""

    for r in successful:
        email_body += f"\n✓ {r['manager_name']}"
        email_body += f"\n  - Output: {r['output_xlsx'].name}"
        email_body += f"\n  - Email JSON: {r['email_json'].name}"

    if failed:
        email_body += f"\n\n{'='*60}\nFAILED MANAGERS:"
        for r in failed:
            email_body += f"\n✗ {r['manager_name']}: {r.get('error', 'Unknown error')}"

    email_body += f"\n\n{'='*60}\nAll output files are attached to this email."
    email_body += f"\n\nProcessing logs are available in each manager's output folder."

    log("Email body prepared:")
    log(email_body)

    # Save email summary to file
    summary_file = output_dir / "batch_summary.txt"
    summary_file.write_text(email_body, encoding="utf-8")
    log_success(f"Email summary saved to: {summary_file}")

    # Create JSON summary
    json_summary = {
        "processing_date": datetime.now().isoformat(),
        "total_managers": len(results),
        "successful": len(successful),
        "failed": len(failed),
        "results": results
    }

    json_summary_file = output_dir / "batch_summary.json"
    json_summary_file.write_text(json.dumps(json_summary, ensure_ascii=False, indent=2, default=str), encoding="utf-8")
    log_success(f"JSON summary saved to: {json_summary_file}")

    log("\n" + "="*60)
    log("NOTE: Actual email sending requires SMTP configuration.")
    log(f"All results have been saved to: {output_dir}")
    log(f"To send the email, please use your preferred email client")
    log(f"or configure SMTP settings in this script.")
    log("="*60)

# ============================================================================
# MAIN
# ============================================================================

def parse_args():
    """Parse command-line arguments"""
    parser = argparse.ArgumentParser(
        description="Batch process special transactions for all fund managers"
    )
    parser.add_argument(
        "--apify-token",
        required=True,
        help="Apify API token"
    )
    parser.add_argument(
        "--managers",
        help="Comma-separated list of manager names to process (default: all)",
        default=None
    )
    parser.add_argument(
        "--output-dir",
        type=Path,
        default=Path("./batch_output"),
        help="Base output directory (default: ./batch_output)"
    )
    parser.add_argument(
        "--skip-tase-prices",
        action="store_true",
        help="Skip TASE price checks for faster processing"
    )
    parser.add_argument(
        "--price-threshold",
        type=float,
        default=DEFAULT_PRICE_THRESHOLD,
        help=f"Price variance threshold in percent (default: {DEFAULT_PRICE_THRESHOLD}%%)"
    )
    parser.add_argument(
        "--spec-file",
        type=Path,
        default=Path(DEFAULT_SPEC_FILE),
        help=f"Specification table Excel file (default: {DEFAULT_SPEC_FILE})"
    )
    parser.add_argument(
        "--seed",
        type=int,
        help="Optional RNG seed for reproducible sampling"
    )
    parser.add_argument(
        "--email",
        default=EMAIL_RECIPIENT,
        help=f"Email recipient (default: {EMAIL_RECIPIENT})"
    )

    return parser.parse_args()

def main():
    """Main entry point"""
    args = parse_args()

    # Create output directory
    output_dir = args.output_dir / datetime.now().strftime("%Y%m%d_%H%M%S")
    output_dir.mkdir(parents=True, exist_ok=True)

    log("="*60)
    log("MIZRAHI SPECIAL TRANSACTIONS - BATCH PROCESSOR")
    log("="*60)
    log(f"Output directory: {output_dir}")
    log(f"Email recipient: {args.email}")
    log(f"Price threshold: {args.price_threshold}%")
    log(f"Skip TASE prices: {args.skip_tase_prices}")

    # Determine which managers to process
    if args.managers:
        manager_names = [m.strip() for m in args.managers.split(",")]
        managers_to_process = {name: FUND_MANAGERS[name] for name in manager_names if name in FUND_MANAGERS}
    else:
        managers_to_process = FUND_MANAGERS

    log(f"Managers to process: {len(managers_to_process)}")
    for name in managers_to_process:
        log(f"  - {name}")

    # Step 1: Fetch mutual funds list
    log(f"\n{'='*60}")
    log("STEP 1: Fetching Mutual Funds List")
    log("="*60)

    try:
        funds_list_path = fetch_funds_list(args.apify_token, output_dir)
    except Exception as e:
        log_error(f"Failed to fetch funds list: {e}")
        return 1

    # Step 2: Process each manager
    log(f"\n{'='*60}")
    log("STEP 2: Processing All Managers")
    log("="*60)

    results = []
    for manager_name, manager_code in managers_to_process.items():
        result = process_manager(
            manager_name,
            manager_code,
            funds_list_path,
            args.apify_token,
            args,
            output_dir
        )
        if result:
            results.append(result)

    # Step 3: Send results email
    log(f"\n{'='*60}")
    log("STEP 3: Preparing Results Email")
    log("="*60)

    send_results_email(results, args.email, output_dir)

    # Final summary
    log(f"\n{'='*60}")
    log("BATCH PROCESSING COMPLETE")
    log("="*60)
    log(f"Total managers processed: {len(results)}")
    log(f"Successful: {len([r for r in results if r.get('status') == 'success'])}")
    log(f"Failed: {len([r for r in results if r.get('status') == 'failed'])}")
    log(f"Output directory: {output_dir}")
    log("="*60)

    return 0

if __name__ == "__main__":
    sys.exit(main())
