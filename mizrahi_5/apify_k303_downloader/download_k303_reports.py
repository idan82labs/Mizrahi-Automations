#!/usr/bin/env python3
"""
K.303 Report Downloader using Apify Puppeteer Scraper

Downloads K.303 disclosure reports from the ISA Magna website
and prepares them for the disclosure_k303_validator.py script.

Usage:
    python download_k303_reports.py --manager "מגדל" --output-dir ./reports

Requirements:
    pip install requests
"""

import argparse
import json
import os
import sys
import time
import urllib.parse
from pathlib import Path

import requests

# Apify configuration
APIFY_TOKEN = os.environ.get("APIFY_TOKEN", "YOUR_APIFY_TOKEN_HERE")
PUPPETEER_SCRAPER_ACTOR_ID = "apify/puppeteer-scraper"
APIFY_API_BASE = "https://api.apify.com/v2"

# Fund manager name mappings (Hebrew to URL-encoded)
MANAGER_SEARCH_TERMS = {
    "מגדל": "מגדל קרנות נאמנות",
    "איילון": "איילון קרנות נאמנות",
    "הראל": "הראל קרנות נאמנות",
    "אנליסט": "אנליסט קרנות נאמנות",
    "קסם": "קסם קרנות נאמנות",
    "מיטב": "מיטב קרנות נאמנות",
}


def get_magna_search_url(manager_name: str) -> str:
    """Build the Magna ISA search URL for K.303 reports."""
    search_term = MANAGER_SEARCH_TERMS.get(manager_name, manager_name)
    encoded_term = urllib.parse.quote(search_term)
    # ק303 URL-encoded is %D7%A7303
    return f"https://www.magna.isa.gov.il/?form=%D7%A7303&q={encoded_term}"


def run_puppeteer_scraper(manager_name: str, token: str) -> dict:
    """Run the Apify Puppeteer Scraper to download K.303 reports."""

    search_url = get_magna_search_url(manager_name)
    print(f"Search URL: {search_url}")

    # Load the page function from file
    page_function_path = Path(__file__).parent / "pageFunction.js"
    if page_function_path.exists():
        with open(page_function_path, "r", encoding="utf-8") as f:
            page_function = f.read()
    else:
        raise FileNotFoundError(f"pageFunction.js not found at {page_function_path}")

    # Build actor input
    actor_input = {
        "startUrls": [
            {
                "url": search_url,
                "userData": {
                    "label": "SEARCH_RESULTS",
                    "managerName": manager_name
                }
            }
        ],
        "keepUrlFragments": False,
        "useChrome": False,
        "headless": True,
        "ignoreSslErrors": False,
        "ignoreCorsAndCsp": False,
        "downloadMedia": True,
        "downloadCss": False,
        "maxConcurrency": 1,
        "maxRequestRetries": 3,
        "maxRequestsPerCrawl": 10,
        "navigationTimeoutSecs": 60,
        "pageLoadTimeoutSecs": 60,
        "proxyConfiguration": {
            "useApifyProxy": True
        },
        "preNavigationHooks": """
async ({ page, request }) => {
    await page.setExtraHTTPHeaders({
        'Accept-Language': 'he-IL,he;q=0.9,en;q=0.8'
    });
}
""",
        "pageFunction": page_function
    }

    # Start the actor run
    print(f"Starting Puppeteer Scraper for manager: {manager_name}...")

    run_url = f"{APIFY_API_BASE}/acts/{PUPPETEER_SCRAPER_ACTOR_ID}/runs"
    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json"
    }

    response = requests.post(run_url, headers=headers, json=actor_input, timeout=60)
    response.raise_for_status()

    run_data = response.json()
    run_id = run_data["data"]["id"]
    print(f"Actor run started: {run_id}")

    return run_data["data"]


def wait_for_run(run_id: str, token: str, timeout: int = 300, poll_interval: int = 5) -> dict:
    """Wait for the actor run to complete."""

    run_url = f"{APIFY_API_BASE}/actor-runs/{run_id}"
    headers = {"Authorization": f"Bearer {token}"}

    start_time = time.time()

    while time.time() - start_time < timeout:
        response = requests.get(run_url, headers=headers, timeout=30)
        response.raise_for_status()

        run_data = response.json()["data"]
        status = run_data["status"]

        print(f"  Status: {status}")

        if status == "SUCCEEDED":
            return run_data
        elif status in ("FAILED", "ABORTED", "TIMED-OUT"):
            raise RuntimeError(f"Actor run failed with status: {status}")

        time.sleep(poll_interval)

    raise TimeoutError(f"Actor run did not complete within {timeout} seconds")


def download_results(run_data: dict, token: str, output_dir: Path) -> list[Path]:
    """Download the XLSX files from the key-value store."""

    kv_store_id = run_data.get("defaultKeyValueStoreId")
    if not kv_store_id:
        raise ValueError("No key-value store found in run data")

    print(f"Downloading files from key-value store: {kv_store_id}")

    # List all keys in the store
    list_url = f"{APIFY_API_BASE}/key-value-stores/{kv_store_id}/keys"
    headers = {"Authorization": f"Bearer {token}"}

    response = requests.get(list_url, headers=headers, timeout=30)
    response.raise_for_status()

    keys_data = response.json()
    keys = [item["key"] for item in keys_data.get("data", {}).get("items", [])]

    print(f"Found {len(keys)} items in store: {keys}")

    # Download XLSX files
    downloaded_files = []
    output_dir.mkdir(parents=True, exist_ok=True)

    for key in keys:
        if key.endswith(".xlsx") or key.endswith(".xls"):
            file_url = f"{APIFY_API_BASE}/key-value-stores/{kv_store_id}/records/{key}"
            response = requests.get(file_url, headers=headers, timeout=60)
            response.raise_for_status()

            output_path = output_dir / key
            with open(output_path, "wb") as f:
                f.write(response.content)

            print(f"  Downloaded: {output_path} ({len(response.content)} bytes)")
            downloaded_files.append(output_path)

    return downloaded_files


def get_dataset_results(run_data: dict, token: str) -> list[dict]:
    """Get the dataset results (metadata about downloads)."""

    dataset_id = run_data.get("defaultDatasetId")
    if not dataset_id:
        return []

    dataset_url = f"{APIFY_API_BASE}/datasets/{dataset_id}/items"
    headers = {"Authorization": f"Bearer {token}"}

    response = requests.get(dataset_url, headers=headers, timeout=30)
    response.raise_for_status()

    return response.json()


def main():
    parser = argparse.ArgumentParser(
        description="Download K.303 reports from ISA Magna using Apify"
    )
    parser.add_argument(
        "--manager", "-m",
        required=True,
        help="Fund manager name in Hebrew (e.g., מגדל, איילון, הראל)"
    )
    parser.add_argument(
        "--output-dir", "-o",
        type=Path,
        default=Path("./k303_reports"),
        help="Output directory for downloaded files"
    )
    parser.add_argument(
        "--token", "-t",
        default=APIFY_TOKEN,
        help="Apify API token (or set APIFY_TOKEN env var)"
    )
    parser.add_argument(
        "--timeout",
        type=int,
        default=300,
        help="Timeout in seconds for the actor run"
    )

    args = parser.parse_args()

    if args.token == "YOUR_APIFY_TOKEN_HERE":
        print("ERROR: Please set your Apify API token via --token or APIFY_TOKEN env var")
        sys.exit(1)

    try:
        # Run the scraper
        run_data = run_puppeteer_scraper(args.manager, args.token)

        # Wait for completion
        print("Waiting for actor to complete...")
        run_data = wait_for_run(run_data["id"], args.token, args.timeout)

        # Download the files
        downloaded_files = download_results(run_data, args.token, args.output_dir)

        # Get metadata
        metadata = get_dataset_results(run_data, args.token)

        print("\n" + "=" * 60)
        print("DOWNLOAD COMPLETE")
        print("=" * 60)
        print(f"Downloaded {len(downloaded_files)} files to: {args.output_dir}")

        for f in downloaded_files:
            print(f"  - {f.name}")

        if metadata:
            print("\nMetadata:")
            print(json.dumps(metadata, indent=2, ensure_ascii=False))

        # Suggest next command
        if len(downloaded_files) >= 2:
            current = next((f for f in downloaded_files if "current" in f.name), downloaded_files[0])
            previous = next((f for f in downloaded_files if "previous" in f.name), downloaded_files[1] if len(downloaded_files) > 1 else None)

            print("\n" + "-" * 60)
            print("To run validation, use:")
            print(f"""
python disclosure_k303_validator.py \\
    --mutual-funds-list "Mutual_Funds_List.csv" \\
    --current-report "{current}" \\
    --previous-report "{previous}" \\
    --output-xlsx "validation_output.xlsx" \\
    --manager-name "{args.manager}"
""")

    except Exception as e:
        print(f"ERROR: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()
