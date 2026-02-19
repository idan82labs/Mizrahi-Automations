"""
Fund Automation Test Script
Run this to test the Apify integration locally.

Usage: python test_fund_automation.py
"""

import os
import sys
import time
import base64
import requests
from datetime import datetime, timedelta

# ============ CONFIGURATION ============
APIFY_TOKEN = "YOUR_APIFY_TOKEN_HERE"
TEST_FUND_NAME = "מגדל"
TEST_FUND_CODE = "10040"

# Apify Actor IDs
FUNDS_LIST_ACTOR_ID = "K9WppTziYC3n2vxTu"
FUND_REPORTS_ACTOR_ID = "5lhI6O39Qbgv9O0gs"
# =======================================


def log(msg):
    print(f"[{datetime.now().strftime('%H:%M:%S')}] {msg}")


def apify_request(method, endpoint, json_data=None, params=None):
    """Make request to Apify API"""
    url = f"https://api.apify.com/v2{endpoint}"
    headers = {"Authorization": f"Bearer {APIFY_TOKEN}"}
    
    response = requests.request(method, url, headers=headers, json=json_data, params=params)
    response.raise_for_status()
    return response


def run_actor_and_wait(actor_id, input_data=None, timeout=180):
    """Run an Apify actor and wait for completion"""
    log(f"Starting actor: {actor_id}")
    
    # Start the run
    resp = apify_request("POST", f"/acts/{actor_id}/runs", json_data=input_data or {}, params={"timeout": timeout})
    run_data = resp.json()["data"]
    run_id = run_data["id"]
    log(f"Run started: {run_id}")
    
    # Poll for completion
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


def build_maya_url(fund_code):
    """Build Maya URL for fund reports"""
    today = datetime.now()
    one_year_ago = today - timedelta(days=365)
    
    return (
        f"https://maya.tase.co.il/he/reports/funds?"
        f"fromDate={one_year_ago.strftime('%Y-%m-%d')}&toDate={today.strftime('%Y-%m-%d')}"
        f"&noMeetings=false&isSingle=false&isIntendToTaseMember=false"
        f"&by=group&groupId=7&itemId={fund_code}&eventsIds%5B%5D=5618"
    )


def main():
    log("=" * 50)
    log("FUND AUTOMATION TEST")
    log(f"Testing with: {TEST_FUND_NAME} (code: {TEST_FUND_CODE})")
    log("=" * 50)
    
    # Create output directory
    os.makedirs("test_output", exist_ok=True)
    
    # ===== STEP 1: Fetch funds list =====
    log("\n--- STEP 1: Fetching funds list (master Excel) ---")
    try:
        run_data = run_actor_and_wait(FUNDS_LIST_ACTOR_ID, {})
        dataset_id = run_data["defaultDatasetId"]
        log(f"Dataset ID: {dataset_id}")
        
        # Get dataset items
        resp = apify_request("GET", f"/datasets/{dataset_id}/items")
        items = resp.json()
        
        if not items:
            log("ERROR: No items in dataset!")
            return
        
        file_base64 = items[0].get("fileBase64")
        if not file_base64:
            log(f"ERROR: No fileBase64 in response. Keys: {items[0].keys()}")
            return
        
        # Save the file
        funds_list_bytes = base64.b64decode(file_base64)
        with open("test_output/funds_list.xlsx", "wb") as f:
            f.write(funds_list_bytes)
        log(f"SUCCESS: Saved funds_list.xlsx ({len(funds_list_bytes)} bytes)")
        
    except Exception as e:
        log(f"ERROR in Step 1: {e}")
        return
    
    # ===== STEP 2: Fetch fund reports =====
    log("\n--- STEP 2: Fetching fund reports (CSVs from Maya) ---")
    try:
        maya_url = build_maya_url(TEST_FUND_CODE)
        log(f"Maya URL: {maya_url[:80]}...")
        
        run_data = run_actor_and_wait(FUND_REPORTS_ACTOR_ID, {"url": maya_url}, timeout=180)
        
        # Get report name from dataset
        dataset_id = run_data["defaultDatasetId"]
        resp = apify_request("GET", f"/datasets/{dataset_id}/items")
        items = resp.json()
        
        report_name = "Unknown"
        if items and items[0].get("downloadedFiles"):
            report_name = items[0]["downloadedFiles"][0].get("reportName", "Unknown")
        log(f"Report name: {report_name}")
        
        # Get CSVs from key-value store
        kv_store_id = run_data["defaultKeyValueStoreId"]
        log(f"Key-Value Store ID: {kv_store_id}")
        
        # Download current month CSV
        resp = apify_request("GET", f"/key-value-stores/{kv_store_id}/records/report_latest_month.csv")
        with open("test_output/current_month.csv", "wb") as f:
            f.write(resp.content)
        log(f"SUCCESS: Saved current_month.csv ({len(resp.content)} bytes)")
        
        # Download previous month CSV
        resp = apify_request("GET", f"/key-value-stores/{kv_store_id}/records/report_previous_month.csv")
        with open("test_output/previous_month.csv", "wb") as f:
            f.write(resp.content)
        log(f"SUCCESS: Saved previous_month.csv ({len(resp.content)} bytes)")
        
    except Exception as e:
        log(f"ERROR in Step 2: {e}")
        import traceback
        traceback.print_exc()
        return
    
    # ===== SUMMARY =====
    log("\n" + "=" * 50)
    log("TEST COMPLETE!")
    log("=" * 50)
    log("\nFiles created in test_output/:")
    for f in os.listdir("test_output"):
        size = os.path.getsize(f"test_output/{f}")
        log(f"  - {f} ({size:,} bytes)")
    
    log("\nIf all 3 files exist with reasonable sizes, the Apify integration works!")
    log("Next step: Deploy to server and test full processing.")


if __name__ == "__main__":
    main()
