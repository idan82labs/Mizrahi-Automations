#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Test Special Transactions Actor (nQh62mdhpUTM5l65l - עסקה מתואמת/מחוץ לבורסה)
Tests the coordinated/off-exchange transactions actor for all fund managers.

Usage:
    python test_special_transactions_all_managers.py --apify-token YOUR_TOKEN
    python test_special_transactions_all_managers.py --apify-token YOUR_TOKEN --managers "מגדל,איילון"
"""

import os
import sys
import io
import time
import base64
import argparse
import json
import subprocess
from datetime import datetime, timedelta
from pathlib import Path
from typing import Dict, List

# Ensure UTF-8 output on Windows
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

try:
    import requests
except ImportError:
    print("Missing required package: requests")
    print("Install with: pip install requests")
    sys.exit(1)

# ============================================================================
# CONFIGURATION
# ============================================================================

# Apify Actor and Build IDs
FUNDS_LIST_ACTOR_ID = "K9WppTziYC3n2vxTu"
SPECIAL_TRANSACTIONS_ACTOR_ID = "5lhI6O39Qbgv9O0gs"
SPECIAL_TRANSACTIONS_BUILD_ID = "nQh62mdhpUTM5l65l"  # Build 0.0.10 - עסקה מתואמת/מחוץ לבורסה

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
        response = requests.request(method, url, headers=headers, json=json_data, params=params, timeout=30)
        response.raise_for_status()
        return response
    except requests.exceptions.RequestException as e:
        log_error(f"Apify API request failed: {e}")
        raise

def run_actor_with_build(actor_id, build_id, token, input_data=None, timeout=300):
    """Run an Apify actor (uses latest build automatically)"""
    log(f"Starting actor: {actor_id}")
    if build_id:
        log(f"Note: Will use latest build (build ID: {build_id})")

    # Start the run (uses latest build by default)
    endpoint = f"/acts/{actor_id}/runs"
    resp = apify_request("POST", endpoint, token, json_data=input_data or {}, params={"timeout": timeout})
    run_data = resp.json()["data"]
    run_id = run_data["id"]
    build_used = run_data.get("buildId", "unknown")
    log(f"Run started: {run_id} (build: {build_used})")

    # Poll for completion
    start = time.time()
    while time.time() - start < timeout:
        resp = apify_request("GET", f"/actor-runs/{run_id}", token)
        status_data = resp.json()["data"]
        status = status_data["status"]

        if status == "SUCCEEDED":
            log_success(f"Actor run completed: {run_id}")
            return status_data
        elif status in ("FAILED", "ABORTED", "TIMED-OUT"):
            log_error(f"Actor failed with status: {status}")
            raise Exception(f"Actor failed with status: {status}")

        log(f"Status: {status}... waiting")
        time.sleep(5)

    raise Exception("Timeout waiting for actor")

def build_maya_special_transactions_url(fund_code):
    """Build Maya URL for special transactions (coordinated/off-exchange)"""
    today = datetime.now()
    one_year_ago = today - timedelta(days=365)

    # Special transactions event ID (עסקה מתואמת/מחוץ לבורסה)
    return (
        f"https://maya.tase.co.il/he/reports/funds?"
        f"fromDate={one_year_ago.strftime('%Y-%m-%d')}&toDate={today.strftime('%Y-%m-%d')}"
        f"&noMeetings=false&isSingle=false&isIntendToTaseMember=false"
        f"&by=group&groupId=7&itemId={fund_code}&eventsIds%5B%5D=5615"  # 5615 is special transactions
    )

def fetch_funds_list(token, output_dir):
    """Fetch mutual funds list from Apify"""
    log("Fetching Mutual Funds List from Apify...")

    try:
        run_data = run_actor_with_build(FUNDS_LIST_ACTOR_ID, None, token, {})
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

def test_manager_special_transactions(manager_name, fund_code, token, output_dir):
    """Test special transactions fetch for a single manager"""
    log(f"\n{'='*60}")
    log(f"Testing manager: {manager_name} (code: {fund_code})")
    log(f"{'='*60}")

    # Create manager directory
    manager_dir = output_dir / manager_name
    manager_dir.mkdir(parents=True, exist_ok=True)

    try:
        # Build Maya URL for special transactions
        maya_url = build_maya_special_transactions_url(fund_code)
        log(f"Maya URL: {maya_url[:100]}...")

        # Run actor with specific build for special transactions
        input_data = {"url": maya_url}
        run_data = run_actor_with_build(
            SPECIAL_TRANSACTIONS_ACTOR_ID,
            SPECIAL_TRANSACTIONS_BUILD_ID,
            token,
            input_data,
            timeout=300
        )

        # Get results from key-value store
        kv_store_id = run_data["defaultKeyValueStoreId"]
        log(f"Key-Value Store ID: {kv_store_id}")

        # Fetch only the latest month report
        report_key = "report_latest_month.csv"
        files_saved = []

        try:
            resp = apify_request("GET", f"/key-value-stores/{kv_store_id}/records/{report_key}", token)
            if resp.status_code == 200 and len(resp.content) > 0:
                file_path = manager_dir / f"{manager_name}_{report_key.replace('.', '_')}.csv"
                file_path.write_bytes(resp.content)
                files_saved.append(file_path)
                log_success(f"Saved {report_key}: {file_path} ({len(resp.content)} bytes)")
            else:
                log_error(f"Report '{report_key}' not found or empty")
        except Exception as e:
            log_error(f"Failed to fetch {report_key}: {e}")

        # Get dataset info if available
        dataset_id = run_data.get("defaultDatasetId")
        if dataset_id:
            try:
                resp = apify_request("GET", f"/datasets/{dataset_id}/items", token)
                dataset_items = resp.json()
                if dataset_items:
                    dataset_file = manager_dir / f"{manager_name}_dataset.json"
                    dataset_file.write_text(json.dumps(dataset_items, ensure_ascii=False, indent=2), encoding='utf-8')
                    files_saved.append(dataset_file)
                    log_success(f"Saved dataset: {dataset_file}")
            except Exception as e:
                log(f"Could not fetch dataset: {e}")

        if files_saved:
            return {
                "manager_name": manager_name,
                "fund_code": fund_code,
                "status": "success",
                "files_saved": [str(f) for f in files_saved],
                "run_id": run_data["id"],
                "kv_store_id": kv_store_id
            }
        else:
            return {
                "manager_name": manager_name,
                "fund_code": fund_code,
                "status": "no_data",
                "message": "No CSV/data files found in key-value store",
                "run_id": run_data["id"]
            }

    except Exception as e:
        log_error(f"Failed to test {manager_name}: {e}")
        return {
            "manager_name": manager_name,
            "fund_code": fund_code,
            "status": "failed",
            "error": str(e)
        }

# ============================================================================
# MAIN
# ============================================================================

def parse_args():
    """Parse command-line arguments"""
    parser = argparse.ArgumentParser(
        description="Test special transactions actor for all fund managers"
    )
    parser.add_argument(
        "--apify-token",
        required=True,
        help="Apify API token"
    )
    parser.add_argument(
        "--managers",
        help="Comma-separated list of manager names to test (default: all)",
        default=None
    )
    parser.add_argument(
        "--output-dir",
        type=Path,
        default=Path("./special_transactions_test"),
        help="Output directory (default: ./special_transactions_test)"
    )
    parser.add_argument(
        "--skip-funds-list",
        action="store_true",
        help="Skip fetching the mutual funds list"
    )

    return parser.parse_args()

def main():
    """Main entry point"""
    args = parse_args()

    # Create output directory with timestamp
    output_dir = args.output_dir / datetime.now().strftime("%Y%m%d_%H%M%S")
    output_dir.mkdir(parents=True, exist_ok=True)

    log("="*60)
    log("SPECIAL TRANSACTIONS ACTOR TEST")
    log("Actor ID: " + SPECIAL_TRANSACTIONS_ACTOR_ID)
    log("Build ID: " + SPECIAL_TRANSACTIONS_BUILD_ID)
    log("="*60)
    log(f"Output directory: {output_dir}")

    # Determine which managers to test
    if args.managers:
        manager_names = [m.strip() for m in args.managers.split(",")]
        managers_to_test = {name: FUND_MANAGERS[name] for name in manager_names if name in FUND_MANAGERS}
    else:
        managers_to_test = FUND_MANAGERS

    log(f"Managers to test: {len(managers_to_test)}")
    for name in managers_to_test:
        log(f"  - {name} ({managers_to_test[name]})")

    # Step 1: Fetch mutual funds list (optional)
    if not args.skip_funds_list:
        log(f"\n{'='*60}")
        log("STEP 1: Fetching Mutual Funds List")
        log("="*60)
        try:
            funds_list_path = fetch_funds_list(args.apify_token, output_dir)
        except Exception as e:
            log_error(f"Failed to fetch funds list (continuing anyway): {e}")

    # Step 2: Test each manager
    log(f"\n{'='*60}")
    log("STEP 2: Testing Special Transactions for Each Manager")
    log("="*60)

    results = []
    for manager_name, fund_code in managers_to_test.items():
        result = test_manager_special_transactions(
            manager_name,
            fund_code,
            args.apify_token,
            output_dir
        )
        results.append(result)

    # Step 3: Generate summary
    log(f"\n{'='*60}")
    log("TEST RESULTS SUMMARY")
    log("="*60)

    successful = [r for r in results if r["status"] == "success"]
    no_data = [r for r in results if r["status"] == "no_data"]
    failed = [r for r in results if r["status"] == "failed"]

    log(f"\nTotal managers tested: {len(results)}")
    log(f"✓ Successful (data found): {len(successful)}")
    log(f"⚠ No data available: {len(no_data)}")
    log(f"✗ Failed: {len(failed)}")

    if successful:
        log("\n--- SUCCESSFUL MANAGERS ---")
        for r in successful:
            log(f"✓ {r['manager_name']}")
            for f in r['files_saved']:
                log(f"  - {Path(f).name}")

    if no_data:
        log("\n--- NO DATA AVAILABLE ---")
        for r in no_data:
            log(f"⚠ {r['manager_name']}: {r.get('message', 'No data')}")

    if failed:
        log("\n--- FAILED MANAGERS ---")
        for r in failed:
            log(f"✗ {r['manager_name']}: {r.get('error', 'Unknown error')}")

    # Save JSON summary
    summary = {
        "test_date": datetime.now().isoformat(),
        "actor_id": SPECIAL_TRANSACTIONS_ACTOR_ID,
        "build_id": SPECIAL_TRANSACTIONS_BUILD_ID,
        "total_managers": len(results),
        "successful": len(successful),
        "no_data": len(no_data),
        "failed": len(failed),
        "results": results
    }

    summary_file = output_dir / "test_summary.json"
    summary_file.write_text(json.dumps(summary, ensure_ascii=False, indent=2, default=str), encoding='utf-8')
    log_success(f"\nSummary saved to: {summary_file}")

    log("\n" + "="*60)
    log(f"All results saved to: {output_dir}")
    log("="*60)

    return 0 if len(failed) == 0 else 1

if __name__ == "__main__":
    sys.exit(main())
