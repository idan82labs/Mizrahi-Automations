#!/usr/bin/env python3
"""
Mizrahi_4 Daily Tracking Automation - Complete Orchestrator Script
Validates inputs, fetches data from Apify actors, and runs mizrahi_4_logic.py.

Usage:
    python Mizrahi_4/mizrahi_4_automation.py
    python Mizrahi_4/mizrahi_4_automation.py --output-dir ./reports --keep-temp
"""

import os
import sys
import base64
import argparse
import tempfile
import shutil
import subprocess
from datetime import datetime
from pathlib import Path

try:
    import pandas as pd
    from dotenv import load_dotenv
except ImportError as e:
    print(f"Missing required package: {e}")
    print("Install with: pip install pandas openpyxl requests python-dotenv")
    sys.exit(1)

# Add parent directory to path for shared imports
sys.path.insert(0, str(Path(__file__).parent.parent))

from shared.apify_client import apify_request, run_actor_and_wait, log
from shared.constants import APIFY_ACTORS

# Load .env file from project root
load_dotenv(Path(__file__).parent.parent / ".env")

# ============================================================================
# CONFIGURATION
# ============================================================================

APIFY_TOKEN = os.environ.get("APIFY_TOKEN", "")

# Default input file paths (relative to Mizrahi_4/)
DEFAULT_FUND_INDEX_TABLE = "input/fund-index-table.xlsx"
DEFAULT_BFIX_PRICES = "input/BFIX PRICE BLOOMBERG.xlsx"
DEFAULT_BLOOMBERG_INDEX = "input/bloomberg-index.xlsx"


# ============================================================================
# STEP 1: VALIDATE LOCAL INPUT FILES
# ============================================================================

def validate_local_inputs(
    fund_index_table: Path,
    bfix_prices: Path,
    bloomberg_index: Path | None
) -> None:
    """Validate that required local input files exist and are readable."""
    log("Validating local input files...")

    # Required files
    if not fund_index_table.exists():
        raise FileNotFoundError(f"fund-index-table.xlsx not found: {fund_index_table}")
    log(f"  ✓ {fund_index_table.name}")

    if not bfix_prices.exists():
        raise FileNotFoundError(f"BFIX PRICE BLOOMBERG.xlsx not found: {bfix_prices}")
    log(f"  ✓ {bfix_prices.name}")

    # Optional file
    if bloomberg_index:
        if not bloomberg_index.exists():
            log(f"  ⚠ bloomberg-index.xlsx not found (optional): {bloomberg_index}")
        else:
            log(f"  ✓ {bloomberg_index.name}")

    # Validate fund-index-table is readable and has expected columns
    try:
        df = pd.read_excel(fund_index_table)
        required_cols = ["מס' קרן", "מספר מדד", "לינק", "מקור נתונים"]
        missing = [c for c in required_cols if c not in df.columns]
        if missing:
            raise ValueError(f"fund-index-table.xlsx missing columns: {missing}")
        log(f"  ✓ fund-index-table.xlsx has {len(df)} funds")
    except Exception as e:
        raise ValueError(f"Cannot read fund-index-table.xlsx: {e}")

    log("All required local files validated")


# ============================================================================
# STEP 2: FETCH FUNDS LIST
# ============================================================================

def fetch_funds_list(temp_dir: Path) -> Path:
    """Fetch master funds list from Apify and save as CSV."""
    log("Fetching funds list from Apify...")

    run_data = run_actor_and_wait(APIFY_TOKEN, APIFY_ACTORS["FUNDS_LIST"], {})
    dataset_id = run_data["defaultDatasetId"]

    resp = apify_request(APIFY_TOKEN, "GET", f"/datasets/{dataset_id}/items")
    items = resp.json()

    if not items or not items[0].get("fileBase64"):
        raise Exception("No fileBase64 in funds list response")

    # Save as Excel first
    xlsx_path = temp_dir / "funds_list.xlsx"
    xlsx_bytes = base64.b64decode(items[0]["fileBase64"])
    xlsx_path.write_bytes(xlsx_bytes)
    log(f"Saved funds list: {xlsx_path} ({len(xlsx_bytes):,} bytes)")

    # Convert to CSV for mizrahi_4_logic.py
    csv_path = temp_dir / "funds_list.csv"
    df = pd.read_excel(xlsx_path)
    df.to_csv(csv_path, index=False, encoding='utf-8-sig')
    log(f"Converted to CSV: {csv_path}")

    return csv_path


# ============================================================================
# STEP 3: FETCH INDX HISTORICAL DATA
# ============================================================================

def get_indx_urls_from_fund_table(fund_index_table: Path) -> list[str]:
    """Extract INDX URLs for funds with 'אינדקס' data source."""
    df = pd.read_excel(fund_index_table)

    # Filter to funds with "אינדקס" data source
    indx_funds = df[df["מקור נתונים"] == "אינדקס"]

    # Extract URLs from column C (לינק)
    urls = indx_funds["לינק"].dropna().tolist()

    # Deduplicate (multiple funds may share same index)
    urls = list(set(urls))

    log(f"Found {len(urls)} unique INDX URLs for {len(indx_funds)} funds")
    return urls


def fetch_indx_historical_data(
    fund_index_table: Path,
    temp_dir: Path,
    save_local_dir: Path | None = None
) -> Path | None:
    """Fetch INDX historical data files from Apify.

    Args:
        fund_index_table: Path to fund-index-table.xlsx
        temp_dir: Temp directory for downloads
        save_local_dir: If provided, save files here; otherwise use temp_dir

    Returns:
        Path to directory containing the downloaded files
    """
    log("Fetching INDX historical data from Apify...")

    # Get URLs to scrape
    index_urls = get_indx_urls_from_fund_table(fund_index_table)

    if not index_urls:
        log("No INDX funds found - skipping INDX data fetch")
        return None

    log(f"Index URLs to fetch: {index_urls}")

    # Run INDX actor
    run_data = run_actor_and_wait(
        APIFY_TOKEN,
        APIFY_ACTORS["INDX_HISTORICAL"],
        {"indexUrls": index_urls},
        timeout=300  # Allow more time for multiple downloads
    )

    kv_store_id = run_data["defaultKeyValueStoreId"]
    log(f"Key-Value Store ID: {kv_store_id}")

    # Determine output directory
    if save_local_dir:
        records_dir = save_local_dir
        records_dir.mkdir(parents=True, exist_ok=True)
        log(f"Saving INDX files to local directory: {records_dir}")
    else:
        records_dir = temp_dir / "indx_records"
        records_dir.mkdir(exist_ok=True)
        log("Saving INDX files to temp directory")

    # List all keys in the KV store
    resp = apify_request(APIFY_TOKEN, "GET", f"/key-value-stores/{kv_store_id}/keys")
    keys = resp.json()["data"]["items"]

    # Download all Excel files
    downloaded = 0
    for key_info in keys:
        key = key_info["key"]
        if key.endswith("_Historical_Data.xlsx"):
            resp = apify_request(APIFY_TOKEN, "GET", f"/key-value-stores/{kv_store_id}/records/{key}")
            file_path = records_dir / key
            file_path.write_bytes(resp.content)
            log(f"  Downloaded: {key} ({len(resp.content):,} bytes)")
            downloaded += 1

    log(f"Downloaded {downloaded} INDX historical data files")
    return records_dir


# ============================================================================
# STEP 4: RUN MIZRAHI_4_LOGIC.PY
# ============================================================================

def run_mizrahi_4_logic(
    funds_list_path: Path,
    fund_index_table: Path,
    bfix_prices: Path,
    bloomberg_index: Path | None,
    indx_records: Path | None,
    output_xlsx: Path,
) -> Path:
    """Run mizrahi_4_logic.py as subprocess."""
    log("Running mizrahi_4_logic.py...")

    logic_script = Path(__file__).parent / "mizrahi_4_logic.py"

    cmd = [
        sys.executable,
        str(logic_script),
        "--mutual-funds-list", str(funds_list_path),
        "--fund-index-table", str(fund_index_table),
        "--bfix-prices", str(bfix_prices),
        "--output-xlsx", str(output_xlsx),
    ]

    # Add optional arguments
    if bloomberg_index and bloomberg_index.exists():
        cmd.extend(["--bloomberg-index", str(bloomberg_index)])

    if indx_records and indx_records.exists():
        cmd.extend(["--indx-records", str(indx_records)])

    log(f"Running: {' '.join(cmd[:4])}...")
    result = subprocess.run(cmd, capture_output=True, text=True, encoding='utf-8')

    if result.returncode != 0:
        log(f"Logic script stderr: {result.stderr}")
        raise Exception(f"mizrahi_4_logic.py failed with code {result.returncode}")

    # Print output (last few lines)
    for line in result.stdout.strip().split('\n')[-10:]:
        if line:
            log(f"  {line}")

    return output_xlsx


# ============================================================================
# MAIN
# ============================================================================

def main():
    parser = argparse.ArgumentParser(
        description="Mizrahi_4 Daily Tracking Automation"
    )
    parser.add_argument(
        "--fund-index-table",
        type=Path,
        default=Path(__file__).parent / DEFAULT_FUND_INDEX_TABLE,
        help="Path to fund-index-table.xlsx"
    )
    parser.add_argument(
        "--bfix-prices",
        type=Path,
        default=Path(__file__).parent / DEFAULT_BFIX_PRICES,
        help="Path to BFIX PRICE BLOOMBERG.xlsx"
    )
    parser.add_argument(
        "--bloomberg-index",
        type=Path,
        default=Path(__file__).parent / DEFAULT_BLOOMBERG_INDEX,
        help="Path to bloomberg-index.xlsx (optional)"
    )
    parser.add_argument(
        "--output-dir",
        type=Path,
        default=Path(__file__).parent / "output",
        help="Output directory for reports"
    )
    parser.add_argument(
        "--keep-temp",
        action="store_true",
        help="Keep temporary files after completion"
    )
    parser.add_argument(
        "--skip-indx",
        action="store_true",
        help="Skip fetching INDX historical data"
    )
    parser.add_argument(
        "--save-indx-local",
        type=Path,
        default=None,
        help="Save downloaded INDX files to this local directory (disabled by default, uses temp location)"
    )

    args = parser.parse_args()

    log("=" * 60)
    log("MIZRAHI_4 DAILY TRACKING AUTOMATION")
    log("=" * 60)
    log(f"Fund Index Table: {args.fund_index_table}")
    log(f"BFIX Prices:      {args.bfix_prices}")
    log(f"Bloomberg Index:  {args.bloomberg_index}")
    log(f"Output Dir:       {args.output_dir}")
    log(f"Save INDX Local:  {args.save_indx_local or '(disabled, using temp)'}")
    log("=" * 60)

    # Check APIFY_TOKEN
    if not APIFY_TOKEN:
        log("ERROR: APIFY_TOKEN environment variable not set")
        sys.exit(1)

    # Create temp directory
    temp_dir = Path(tempfile.mkdtemp(prefix="mizrahi4_"))
    log(f"Temp directory: {temp_dir}")

    try:
        # STEP 1: Validate local inputs
        log("\n--- STEP 1: Validating local input files ---")
        validate_local_inputs(
            args.fund_index_table,
            args.bfix_prices,
            args.bloomberg_index
        )

        # STEP 2: Fetch funds list
        log("\n--- STEP 2: Fetching mutual funds list ---")
        funds_list_path = fetch_funds_list(temp_dir)

        # STEP 3: Fetch INDX historical data
        indx_records = None
        if not args.skip_indx:
            log("\n--- STEP 3: Fetching INDX historical data ---")
            indx_records = fetch_indx_historical_data(
                args.fund_index_table,
                temp_dir,
                save_local_dir=args.save_indx_local
            )
        else:
            log("\n--- STEP 3: Skipping INDX historical data (--skip-indx) ---")

        # STEP 4: Run logic script
        log("\n--- STEP 4: Running mizrahi_4_logic.py ---")
        args.output_dir.mkdir(parents=True, exist_ok=True)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_filename = f"mizrahi_4_output_{timestamp}.xlsx"
        output_path = args.output_dir / output_filename

        run_mizrahi_4_logic(
            funds_list_path=funds_list_path,
            fund_index_table=args.fund_index_table,
            bfix_prices=args.bfix_prices,
            bloomberg_index=args.bloomberg_index,
            indx_records=indx_records,
            output_xlsx=output_path,
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
