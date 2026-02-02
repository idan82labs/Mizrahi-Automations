#!/usr/bin/env python3
"""
K.303 Disclosure Report Automation - Complete Standalone Script
Fetches files from Apify actors and runs disclosure_k303_validator.

Usage:
    python k303_automation_complete.py --fund-name "מגדל"
    python k303_automation_complete.py --fund-name "מגדל" --output-dir ./reports --keep-temp
"""

import os
import sys
import time
import base64
import argparse
import tempfile
import shutil
import subprocess
from datetime import datetime, timedelta
from pathlib import Path

# Try to import required packages
try:
    import requests
    import pandas as pd
except ImportError as e:
    print(f"Missing required package: {e}")
    print("Install with: pip install pandas openpyxl requests")
    sys.exit(1)

# ============================================================================
# CONFIGURATION
# ============================================================================

APIFY_TOKEN = os.environ.get("APIFY_TOKEN", "")

# Apify Actor IDs
FUNDS_LIST_ACTOR_ID = "K9WppTziYC3n2vxTu"
K303_REPORTS_ACTOR_ID = "iTpNz9ixbdQCmH43C"

# Fund Manager Codes
FUND_MANAGER_CODES = {
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

# ============================================================================
# LOGGING
# ============================================================================

def log(msg: str) -> None:
    """Print timestamped log message."""
    print(f"[{datetime.now().strftime('%H:%M:%S')}] {msg}")


# ============================================================================
# APIFY FUNCTIONS
# ============================================================================

def apify_request(method: str, endpoint: str, json_data: dict = None, params: dict = None):
    """Make request to Apify API."""
    url = f"https://api.apify.com/v2{endpoint}"
    headers = {"Authorization": f"Bearer {APIFY_TOKEN}"}
    response = requests.request(method, url, headers=headers, json=json_data, params=params)
    response.raise_for_status()
    return response


def run_actor_and_wait(actor_id: str, input_data: dict = None, timeout: int = 180) -> dict:
    """Run an Apify actor and wait for completion."""
    log(f"Starting Apify actor: {actor_id}")

    resp = apify_request(
        "POST",
        f"/acts/{actor_id}/runs",
        json_data=input_data or {},
        params={"timeout": timeout}
    )
    run_data = resp.json()["data"]
    run_id = run_data["id"]
    log(f"Run started: {run_id}")

    start = time.time()
    while time.time() - start < timeout:
        resp = apify_request("GET", f"/actor-runs/{run_id}")
        status = resp.json()["data"]["status"]

        if status == "SUCCEEDED":
            log(f"Run completed: {run_id}")
            return resp.json()["data"]
        elif status in ("FAILED", "ABORTED", "TIMED-OUT"):
            raise Exception(f"Actor failed with status: {status}")

        log(f"Status: {status}... waiting")
        time.sleep(3)

    raise Exception("Timeout waiting for actor")


# ============================================================================
# STEP 1: Fetch Funds List
# ============================================================================

def fetch_funds_list(temp_dir: Path) -> Path:
    """Fetch master funds list from Apify and save as CSV."""
    log("Fetching funds list (master Excel)...")

    run_data = run_actor_and_wait(FUNDS_LIST_ACTOR_ID, {})
    dataset_id = run_data["defaultDatasetId"]

    resp = apify_request("GET", f"/datasets/{dataset_id}/items")
    items = resp.json()

    if not items or not items[0].get("fileBase64"):
        raise Exception("No fileBase64 in funds list response")

    # Save as Excel first
    xlsx_path = temp_dir / "funds_list.xlsx"
    xlsx_bytes = base64.b64decode(items[0]["fileBase64"])
    xlsx_path.write_bytes(xlsx_bytes)
    log(f"Saved funds list: {xlsx_path} ({len(xlsx_bytes):,} bytes)")

    # Convert to CSV for validator
    csv_path = temp_dir / "funds_list.csv"
    df = pd.read_excel(xlsx_path)
    df.to_csv(csv_path, index=False, encoding='utf-8-sig')
    log(f"Converted to CSV: {csv_path}")

    return csv_path


# ============================================================================
# STEP 2: Fetch K.303 Reports
# ============================================================================

def fetch_k303_reports(fund_name: str, temp_dir: Path) -> tuple[Path, Path, str | None]:
    """Fetch K.303 disclosure reports from Apify.

    Returns:
        Tuple of (current_report_path, previous_report_path, report_month)
        report_month is YYYY-MM format derived from the latest report's upload date (minus 1 month)
    """
    log(f"Fetching K.303 reports for: {fund_name}")

    # Run the K.303 actor with fund manager name
    run_data = run_actor_and_wait(
        K303_REPORTS_ACTOR_ID,
        {"fundManagerName": fund_name},
        timeout=300  # K.303 downloads can be slow
    )

    kv_store_id = run_data["defaultKeyValueStoreId"]
    dataset_id = run_data["defaultDatasetId"]
    log(f"Key-Value Store ID: {kv_store_id}")

    # Get report metadata from dataset to extract report month from report name
    # Example: "נתוני גילוי נאות לחודש דצמבר 2025" -> 2025-12
    # The report name already contains the actual month the report is FOR (not upload date)
    report_month = None
    import re
    try:
        resp = apify_request("GET", f"/datasets/{dataset_id}/items")
        items = resp.json()

        # Check if the actor run actually succeeded in downloading files
        if items and items[0].get("status") == "failed":
            error_msg = items[0].get("error", "Unknown error")
            raise Exception(f"K.303 actor failed to download files: {error_msg}")

        if items and items[0].get("downloadedFiles"):
            first_file = items[0]["downloadedFiles"][0]
            report_name = first_file.get("reportName", "")
            log(f"Report name: {report_name}")

            # Extract month and year from Hebrew report name
            # Format: "נתוני גילוי נאות לחודש <month> <year>" or similar
            # The month in the report name is the actual report month (not upload month)
            hebrew_months = {
                "ינואר": "01", "פברואר": "02", "מרץ": "03", "אפריל": "04",
                "מאי": "05", "יוני": "06", "יולי": "07", "אוגוסט": "08",
                "ספטמבר": "09", "אוקטובר": "10", "נובמבר": "11", "דצמבר": "12",
            }

            for heb_month, num_month in hebrew_months.items():
                if heb_month in report_name:
                    # Find year (4 digits)
                    year_match = re.search(r'20\d{2}', report_name)
                    if year_match:
                        year = year_match.group()
                        report_month = f"{year}-{num_month}"
                        log(f"Extracted report month from name: {report_month}")
                        break
    except Exception as e:
        log(f"Warning: Could not extract report date from metadata: {e}")
        raise

    # Download current month report
    current_path = temp_dir / "report_latest_month.csv"
    resp = apify_request("GET", f"/key-value-stores/{kv_store_id}/records/report_latest_month.csv")
    current_path.write_bytes(resp.content)
    log(f"Downloaded current report: {current_path} ({len(resp.content):,} bytes)")

    # Download previous month report
    previous_path = temp_dir / "report_previous_month.csv"
    resp = apify_request("GET", f"/key-value-stores/{kv_store_id}/records/report_previous_month.csv")
    previous_path.write_bytes(resp.content)
    log(f"Downloaded previous report: {previous_path} ({len(resp.content):,} bytes)")

    return current_path, previous_path, report_month


# ============================================================================
# STEP 3: Run K.303 Validator
# ============================================================================

def run_k303_validator(
    funds_list_path: Path,
    current_report_path: Path,
    previous_report_path: Path,
    output_xlsx_path: Path,
    manager_name: str,
    report_month: str,
) -> Path:
    """Run the K.303 disclosure validator as subprocess."""
    log("Running K.303 disclosure validator...")

    validator_script = Path(__file__).parent / "disclosure_k303_validator.py"

    cmd = [
        sys.executable,
        str(validator_script),
        "--mutual-funds-list", str(funds_list_path),
        "--current-report", str(current_report_path),
        "--previous-report", str(previous_report_path),
        "--output-xlsx", str(output_xlsx_path),
        "--manager-name", manager_name,
        "--report-month", report_month,
    ]

    log(f"Running: {' '.join(cmd[:4])}...")
    result = subprocess.run(cmd, capture_output=True, text=True, encoding='utf-8')

    if result.returncode != 0:
        log(f"Validator stderr: {result.stderr}")
        raise Exception(f"Validator failed with code {result.returncode}")

    # Print validator output
    for line in result.stdout.strip().split('\n'):
        if line:
            log(f"  {line}")

    return output_xlsx_path


# ============================================================================
# MAIN
# ============================================================================

def main():
    parser = argparse.ArgumentParser(
        description="K.303 Disclosure Report Automation"
    )
    parser.add_argument(
        "--fund-name",
        required=True,
        choices=list(FUND_MANAGER_CODES.keys()),
        help="Fund manager name (Hebrew)"
    )
    parser.add_argument(
        "--output-dir",
        type=Path,
        default=Path(__file__).parent / "output",
        help="Output directory for reports (default: mizrahi_5/output)"
    )
    parser.add_argument(
        "--report-month",
        type=str,
        default=None,
        help="Report month YYYY-MM (default: previous month)"
    )
    parser.add_argument(
        "--keep-temp",
        action="store_true",
        help="Keep temporary files after completion"
    )

    args = parser.parse_args()

    log("=" * 60)
    log("K.303 DISCLOSURE REPORT AUTOMATION")
    log("=" * 60)
    log(f"Fund Manager: {args.fund_name}")
    log(f"Output Dir:   {args.output_dir}")
    log("=" * 60)

    # Create temp directory
    temp_dir = Path(tempfile.mkdtemp(prefix="k303_"))
    log(f"Temp directory: {temp_dir}")

    try:
        # STEP 1: Fetch funds list
        log("\n--- STEP 1: Fetching master funds list ---")
        funds_list_path = fetch_funds_list(temp_dir)

        # STEP 2: Fetch K.303 reports
        log("\n--- STEP 2: Fetching K.303 disclosure reports ---")
        current_path, previous_path, scraped_report_month = fetch_k303_reports(args.fund_name, temp_dir)

        # Determine report month: CLI arg > scraped date > fallback to previous month
        if args.report_month:
            report_month = args.report_month
            log(f"Using CLI-specified report month: {report_month}")
        elif scraped_report_month:
            report_month = scraped_report_month
            log(f"Using report month from scraped data: {report_month}")
        else:
            # Fallback to previous month
            today = datetime.now()
            first_of_month = today.replace(day=1)
            last_month = first_of_month - timedelta(days=1)
            report_month = last_month.strftime("%Y-%m")
            log(f"Fallback: using previous month: {report_month}")

        log(f"Report Month: {report_month}")

        # STEP 3: Run validator
        log("\n--- STEP 3: Running K.303 validator ---")
        args.output_dir.mkdir(parents=True, exist_ok=True)
        output_filename = f"k303_validation_{args.fund_name}_{report_month}.xlsx"
        output_path = args.output_dir / output_filename

        run_k303_validator(
            funds_list_path=funds_list_path,
            current_report_path=current_path,
            previous_report_path=previous_path,
            output_xlsx_path=output_path,
            manager_name=args.fund_name,
            report_month=report_month,
        )

        # Summary
        log("\n" + "=" * 60)
        log("SUCCESS!")
        log("=" * 60)
        log(f"Output file: {output_path}")

        if args.keep_temp:
            log(f"Temp files kept at: {temp_dir}")

    except Exception as e:
        log(f"\nERROR: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)

    finally:
        if not args.keep_temp:
            shutil.rmtree(temp_dir, ignore_errors=True)
            log("Temp files cleaned up")


if __name__ == "__main__":
    main()
