#!/usr/bin/env python3
"""
Mizrahi Special Transactions - CLI Processor (single script)

Implements checks #1–#5.1 from the current specification.

Inputs:
  - Mutual Funds List (XLSX): filtered to Mizrahi trustee funds (by 'שם נאמן')
  - Manager special transactions report (CSV or XLSX): e.g., 1702431.csv or 1702431.xlsx

Outputs:
  - Output XLSX: summary + exceptions + samples (+ out-of-scope funds)
  - Email JSON: for n8n workflow - contains ONLY two JSON objects with transaction info (no full email body)

Dependencies:
  - Python 3.10+
  - openpyxl

Example:
  python mizrahi_special_transactions.py \
    --mutual-funds-list "Mutual Funds List.xlsx" \
    --input-report "1702431.csv" \
    --output-xlsx "output.xlsx" \
    --email-json "email.json" \
    --seed 123

Notes:
  - Hebrew column headers are expected (as in your provided files).
  - Dates in the manager report are often stored as numbers DDMMYYYY without leading zeros.
  - Times are often stored as numbers HHMMSS without leading zeros.
  - CSV files are expected to be UTF-8 encoded (BOM is handled automatically).
  - Unique ID for transactions is built from security number + date (as per specification step 1).
"""

from __future__ import annotations

import argparse
import csv
import datetime as dt
import json
import random
import re
import time
from collections import Counter, defaultdict
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Optional

import openpyxl
from openpyxl.styles import Font

# Selenium imports (optional - for price checks)
try:
    from selenium import webdriver
    from selenium.webdriver.chrome.service import Service
    from selenium.webdriver.chrome.options import Options
    from selenium.webdriver.common.by import By
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC
    from webdriver_manager.chrome import ChromeDriverManager
    SELENIUM_AVAILABLE = True
except ImportError:
    SELENIUM_AVAILABLE = False


# -----------------------------
# Configuration (column headers)
# -----------------------------

MIZRAHI_TRUSTEE_NAME_DEFAULT = 'מזרחי טפחות חברה לנאמנות בע"מ'

# Mutual Funds List headers
MF_COL_FUND_ID = "מספר בורסה"
MF_COL_TRUSTEE = "שם נאמן"

# Manager report headers
R_COL_FUND_NO = "מספר קרן"
R_COL_FUND_NAME = "שם קרן"
R_COL_SECURITY_NAME = "שם נייר"
R_COL_SECURITY_NO = "מספר נייר"
R_COL_QUANTITY = "כמות"
R_COL_PRICE = "מחיר"
R_COL_DATE = "תאריך"
R_COL_TIME = "שעה"
R_COL_TYPE = "סוג"
R_COL_DECISION = "אופן החלטה"
R_COL_REPORT_DATE = "ת.דוח"  # used to infer month if --report-month omitted

# Decision-method rules (check #4)
TYPE_REQUIRES_DECISION_1 = {12, 22}
TYPE_REQUIRES_DECISION_1_OR_2 = {31, 32, 33, 34, 35, 36}

# Price checks (spec #6)
TASE_PRICE_CHECK_TYPES = {12, 21, 22}  # Types requiring TASE price comparison
TASE_SAMPLES_PER_TYPE = 2
TASE_VARIANCE_THRESHOLD = 0.05  # 5%
PRICE_LIMIT_TYPES = {31, 32, 33, 34, 35, 36}  # Types with price > 100 check
PRICE_LIMIT = 100.0

# Problematic securities lists (spec #7)
PROBLEMATIC_LISTS_CONFIG = {
    'low_liquidity': {
        'url': 'https://market.tase.co.il/he/market_data/securities/data/all?dType=1&cl1=0&cl2=2',
        'name_he': 'דלי סחירות',
    },
    'maintenance': {
        'url': 'https://market.tase.co.il/he/market_data/securities/data/all?dType=1&cl1=0&cl2=3',
        'name_he': 'רשימת שימור',
    },
    'suspended': {
        'url': 'https://market.tase.co.il/he/market_data/securities/data/all?dType=1&cl1=0&cl2=4',
        'name_he': 'מושעים',
    },
}


# -----------------------------
# Data structures
# -----------------------------

@dataclass(frozen=True)
class TxnRow:
    row_num: int  # row number in the original sheet (1-based)
    fund_no: Optional[int]
    fund_name: Optional[str]
    security_name: Optional[str]
    security_no: Optional[str]
    quantity: Optional[float]
    price: Optional[float]
    tx_date: Optional[dt.date]
    tx_time: Optional[dt.time]
    tx_type: Optional[int]
    decision_method: Optional[int]
    report_date: Optional[dt.date]

    @property
    def unique_id(self) -> str:
        """Spec: unique id = security number + date."""
        d = self.tx_date.strftime("%d%m%Y") if self.tx_date else ""
        s = self.security_no or ""
        return f"{s}|{d}"


@dataclass(frozen=True)
class ExceptionRow:
    check_id: str
    reason: str
    row: TxnRow
    group_key: str = ""


@dataclass(frozen=True)
class Samples:
    decision_1: Optional[TxnRow]
    decision_2: Optional[TxnRow]


@dataclass
class PriceCheckResult:
    """Result of TASE price comparison check."""
    row: TxnRow
    tase_closing_price: Optional[float] = None
    variance_pct: Optional[float] = None
    is_exception: bool = False
    error_message: Optional[str] = None


@dataclass
class PriceLimitResult:
    """Result of price > 100 check for types 31-36."""
    row: TxnRow
    is_exception: bool = False


@dataclass
class ProblematicSecurityResult:
    """Result of checking a transaction against problematic lists."""
    row: TxnRow
    matched_lists: list = field(default_factory=list)
    is_exception: bool = False


# -----------------------------
# Helpers: parsing / normalization
# -----------------------------

def _norm_spaces(s: str) -> str:
    return " ".join((s or "").strip().split())


def _to_str(v: Any) -> Optional[str]:
    if v is None:
        return None
    s = str(v).strip()
    return s if s else None


def _to_int(v: Any) -> Optional[int]:
    if v is None or v == "":
        return None
    try:
        return int(float(v))
    except Exception:
        return None


def _to_float(v: Any) -> Optional[float]:
    if v is None or v == "":
        return None
    try:
        return float(v)
    except Exception:
        return None


def _parse_ddmmyyyy(v: Any) -> Optional[dt.date]:
    """Parse manager report numeric date DDMMYYYY with zero-padding."""
    if v is None or v == "":
        return None
    s = str(int(v)) if isinstance(v, (int, float)) else str(v)
    s = s.strip().zfill(8)
    try:
        day = int(s[0:2])
        month = int(s[2:4])
        year = int(s[4:8])
        return dt.date(year, month, day)
    except Exception:
        return None


def _parse_hhmmss(v: Any) -> Optional[dt.time]:
    """Parse manager report numeric time HHMMSS with zero-padding."""
    if v is None or v == "":
        return None
    s = str(int(v)) if isinstance(v, (int, float)) else str(v)
    s = s.strip().zfill(6)
    try:
        hh = int(s[0:2])
        mm = int(s[2:4])
        ss = int(s[4:6])
        return dt.time(hh, mm, ss)
    except Exception:
        return None


def _headers_row1(ws) -> dict[str, int]:
    """Return mapping from header text to column index, assuming headers are in row 1."""
    headers: dict[str, int] = {}
    for c in range(1, ws.max_column + 1):
        v = ws.cell(1, c).value
        if isinstance(v, str) and v.strip():
            headers[v.strip()] = c
    return headers


# -----------------------------
# Loaders
# -----------------------------

def load_mizrahi_fund_ids(mutual_funds_list_path: Path, trustee_name: str) -> set[int]:
    """Load fund IDs under Mizrahi trusteeship from the Mutual Funds List.

    Uses:
      - fund id column: מספר בורסה
      - trustee name column: שם נאמן
    """
    wb = openpyxl.load_workbook(mutual_funds_list_path, data_only=True)
    ws = wb[wb.sheetnames[0]]

    headers = _headers_row1(ws)
    fund_id_col = headers.get(MF_COL_FUND_ID)
    trustee_col = headers.get(MF_COL_TRUSTEE)

    if fund_id_col is None or trustee_col is None:
        wb.close()
        raise ValueError(f"Mutual Funds List must contain columns: '{MF_COL_FUND_ID}', '{MF_COL_TRUSTEE}'")

    target = _norm_spaces(trustee_name)
    fund_ids: set[int] = set()

    for r in range(2, ws.max_row + 1):
        trustee = ws.cell(r, trustee_col).value
        if trustee is None:
            continue
        if _norm_spaces(str(trustee)) != target:
            continue

        fid_val = ws.cell(r, fund_id_col).value
        try:
            fid = int(float(fid_val))
        except Exception:
            continue

        fund_ids.add(fid)

    wb.close()
    return fund_ids


def load_manager_report_xlsx(input_report_path: Path) -> tuple[list[TxnRow], dict]:
    """Load manager special-transactions report from XLSX. Expects headers in row 1."""
    wb = openpyxl.load_workbook(input_report_path, data_only=True)
    ws = wb[wb.sheetnames[0]]
    headers = _headers_row1(ws)

    def col(name: str) -> Optional[int]:
        return headers.get(name)

    required = [R_COL_FUND_NO, R_COL_SECURITY_NO, R_COL_QUANTITY, R_COL_PRICE, R_COL_DATE, R_COL_TIME, R_COL_TYPE, R_COL_DECISION]
    missing = [r for r in required if col(r) is None]
    if missing:
        wb.close()
        raise ValueError(f"Manager report missing required columns: {missing}")

    rows: list[TxnRow] = []
    rec_col = headers.get("מס. רשומה")  # optional, helps skip trailing note rows

    for r in range(2, ws.max_row + 1):
        rec = ws.cell(r, rec_col).value if rec_col else None
        fund_no = _to_int(ws.cell(r, col(R_COL_FUND_NO)).value)
        sec_no = _to_str(ws.cell(r, col(R_COL_SECURITY_NO)).value)

        if fund_no is None and sec_no is None and rec is None:
            continue

        row = TxnRow(
            row_num=r,
            fund_no=fund_no,
            fund_name=_to_str(ws.cell(r, col(R_COL_FUND_NAME)).value) if col(R_COL_FUND_NAME) else None,
            security_name=_to_str(ws.cell(r, col(R_COL_SECURITY_NAME)).value) if col(R_COL_SECURITY_NAME) else None,
            security_no=sec_no,
            quantity=_to_float(ws.cell(r, col(R_COL_QUANTITY)).value),
            price=_to_float(ws.cell(r, col(R_COL_PRICE)).value),
            tx_date=_parse_ddmmyyyy(ws.cell(r, col(R_COL_DATE)).value),
            tx_time=_parse_hhmmss(ws.cell(r, col(R_COL_TIME)).value),
            tx_type=_to_int(ws.cell(r, col(R_COL_TYPE)).value),
            decision_method=_to_int(ws.cell(r, col(R_COL_DECISION)).value),
            report_date=_parse_ddmmyyyy(ws.cell(r, col(R_COL_REPORT_DATE)).value) if col(R_COL_REPORT_DATE) else None,
        )
        rows.append(row)

    inferred_month = None
    for row in rows:
        if row.report_date:
            inferred_month = f"{row.report_date.year:04d}-{row.report_date.month:02d}"
            break

    wb.close()
    meta = {"sheet": ws.title, "rows_parsed": len(rows), "report_month_inferred": inferred_month}
    return rows, meta


def load_manager_report_csv(input_report_path: Path) -> tuple[list[TxnRow], dict]:
    """Load manager special-transactions report from CSV. Expects headers in row 1.

    The CSV may have a BOM (byte order mark) which is handled by utf-8-sig encoding.
    """
    rows: list[TxnRow] = []

    with open(input_report_path, 'r', encoding='utf-8-sig') as f:
        reader = csv.DictReader(f)

        # Normalize header names (strip whitespace)
        if reader.fieldnames:
            reader.fieldnames = [h.strip() for h in reader.fieldnames]

        required = [R_COL_FUND_NO, R_COL_SECURITY_NO, R_COL_QUANTITY, R_COL_PRICE, R_COL_DATE, R_COL_TIME, R_COL_TYPE, R_COL_DECISION]
        missing = [r for r in required if r not in (reader.fieldnames or [])]
        if missing:
            raise ValueError(f"Manager report CSV missing required columns: {missing}")

        for row_num, csv_row in enumerate(reader, start=2):  # start=2 because row 1 is header
            fund_no = _to_int(csv_row.get(R_COL_FUND_NO))
            sec_no = _to_str(csv_row.get(R_COL_SECURITY_NO))
            rec = csv_row.get("מס. רשומה")

            if fund_no is None and sec_no is None and not rec:
                continue

            row = TxnRow(
                row_num=row_num,
                fund_no=fund_no,
                fund_name=_to_str(csv_row.get(R_COL_FUND_NAME)),
                security_name=_to_str(csv_row.get(R_COL_SECURITY_NAME)),
                security_no=sec_no,
                quantity=_to_float(csv_row.get(R_COL_QUANTITY)),
                price=_to_float(csv_row.get(R_COL_PRICE)),
                tx_date=_parse_ddmmyyyy(csv_row.get(R_COL_DATE)),
                tx_time=_parse_hhmmss(csv_row.get(R_COL_TIME)),
                tx_type=_to_int(csv_row.get(R_COL_TYPE)),
                decision_method=_to_int(csv_row.get(R_COL_DECISION)),
                report_date=_parse_ddmmyyyy(csv_row.get(R_COL_REPORT_DATE)),
            )
            rows.append(row)

    inferred_month = None
    for row in rows:
        if row.report_date:
            inferred_month = f"{row.report_date.year:04d}-{row.report_date.month:02d}"
            break

    meta = {"source": str(input_report_path), "rows_parsed": len(rows), "report_month_inferred": inferred_month}
    return rows, meta


def load_manager_report(input_report_path: Path) -> tuple[list[TxnRow], dict]:
    """Load manager special-transactions report. Detects format by file extension."""
    suffix = input_report_path.suffix.lower()
    if suffix == '.csv':
        return load_manager_report_csv(input_report_path)
    elif suffix in ('.xlsx', '.xls'):
        return load_manager_report_xlsx(input_report_path)
    else:
        raise ValueError(f"Unsupported file format: {suffix}. Use .csv or .xlsx")


# -----------------------------
# Checks (spec 1–5.1)
# -----------------------------

def check_1_duplicates_exact(rows: list[TxnRow]) -> list[ExceptionRow]:
    """Spec #1: identical date/time/security number/quantity/price -> flag all rows in the group."""
    buckets: dict[tuple, list[TxnRow]] = defaultdict(list)
    for r in rows:
        key = (
            r.tx_date.isoformat() if r.tx_date else "",
            r.tx_time.isoformat() if r.tx_time else "",
            r.security_no or "",
            r.quantity if r.quantity is not None else "",
            r.price if r.price is not None else "",
        )
        buckets[key].append(r)

    out: list[ExceptionRow] = []
    for key, group in buckets.items():
        if len(group) <= 1:
            continue
        group_key = "|".join(map(str, key))
        for row in group:
            out.append(ExceptionRow(check_id="CHK_1", reason="DUPLICATE_EXACT", row=row, group_key=group_key))
    return out


def check_1_abs_quantity_pairs(rows: list[TxnRow]) -> list[ExceptionRow]:
    """Spec #1: within unique_id(security+date), if there are two rows with same abs(quantity) but DIFFERENT SIGNS -> flag.

    Only flags when one quantity is positive and the other is negative (inter-fund transactions).
    Does NOT flag if both quantities have the same sign.
    """
    by_uid: dict[str, list[TxnRow]] = defaultdict(list)
    for r in rows:
        by_uid[r.unique_id].append(r)

    out: list[ExceptionRow] = []
    for uid, group in by_uid.items():
        abs_map: dict[float, list[TxnRow]] = defaultdict(list)
        for r in group:
            if r.quantity is None or r.quantity == 0:
                continue
            abs_map[abs(r.quantity)].append(r)

        for abs_qty, rs in abs_map.items():
            # Check if there are rows with OPPOSITE signs (one positive, one negative)
            has_positive = any(r.quantity > 0 for r in rs if r.quantity is not None)
            has_negative = any(r.quantity < 0 for r in rs if r.quantity is not None)

            # Only flag if we have BOTH positive and negative quantities with same abs value
            if has_positive and has_negative:
                for r in rs:
                    out.append(ExceptionRow(check_id="CHK_1", reason="עסקה בין קרנות", row=r, group_key=f"{uid}|abs={abs_qty}"))

    return out


def check_3_dates_in_report_month(rows: list[TxnRow], report_month: str) -> list[ExceptionRow]:
    """Spec #3: tx_date must be within same month (YYYY-MM)."""
    y, m = report_month.split("-")
    year = int(y)
    month = int(m)

    out: list[ExceptionRow] = []
    for r in rows:
        if r.tx_date is None:
            out.append(ExceptionRow(check_id="CHK_3", reason="MISSING_TX_DATE", row=r, group_key=report_month))
            continue
        if r.tx_date.year != year or r.tx_date.month != month:
            out.append(ExceptionRow(check_id="CHK_3", reason="DATE_OUT_OF_REPORT_MONTH", row=r, group_key=report_month))
    return out


def check_4_decision_method_rules(rows: list[TxnRow]) -> list[ExceptionRow]:
    """Spec #4: decision method allowed values depend on type."""
    out: list[ExceptionRow] = []
    for r in rows:
        if r.tx_type is None or r.decision_method is None:
            out.append(ExceptionRow(check_id="CHK_4", reason="MISSING_TYPE_OR_DECISION_METHOD", row=r))
            continue

        if r.tx_type in TYPE_REQUIRES_DECISION_1 and r.decision_method != 1:
            out.append(ExceptionRow(check_id="CHK_4", reason=f"TYPE_{r.tx_type}_REQUIRES_DECISION_1", row=r, group_key=f"type={r.tx_type}"))

        if r.tx_type in TYPE_REQUIRES_DECISION_1_OR_2 and r.decision_method not in (1, 2):
            out.append(ExceptionRow(check_id="CHK_4", reason=f"TYPE_{r.tx_type}_REQUIRES_DECISION_1_OR_2", row=r, group_key=f"type={r.tx_type}"))

    return out


def pick_samples(valid_rows: list[TxnRow], seed: Optional[int]) -> Samples:
    """Spec #5: random transaction with decision method 1 and 2 from valid lines."""
    rng = random.Random(seed)
    dm1 = [r for r in valid_rows if r.decision_method == 1]
    dm2 = [r for r in valid_rows if r.decision_method == 2]
    s1 = rng.choice(dm1) if dm1 else None
    s2 = rng.choice(dm2) if dm2 else None
    return Samples(decision_1=s1, decision_2=s2)


# -----------------------------
# Checks (spec #6 - Price)
# -----------------------------

def _init_selenium_driver():
    """Initialize Selenium Chrome driver."""
    if not SELENIUM_AVAILABLE:
        return None
    options = Options()
    options.add_argument('--headless')
    options.add_argument('--no-sandbox')
    options.add_argument('--disable-dev-shm-usage')
    options.add_argument('--disable-gpu')
    options.add_argument('--window-size=1920,1080')
    service = Service(ChromeDriverManager().install())
    return webdriver.Chrome(service=service, options=options)


def _fetch_tase_closing_price(driver, security_no: str, tx_date: dt.date) -> tuple[Optional[float], Optional[str]]:
    """Fetch closing price from TASE website."""
    date_str = tx_date.strftime('%Y-%m-%d')
    url = f'https://market.tase.co.il/he/market_data/security/{security_no}/historical_data/eod?pType=8&oId=0{security_no}&dFrom={date_str}&dTo={date_str}'

    try:
        driver.get(url)
        time.sleep(3)
        try:
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.TAG_NAME, 'table')))
        except:
            pass
        time.sleep(2)

        body = driver.find_element(By.TAG_NAME, 'body')
        text = body.text

        if "לא נמצאו תוצאות" in text or "אין נתונים" in text:
            return None, "לא נמצאו נתונים לתאריך זה"

        tables = driver.find_elements(By.TAG_NAME, 'table')
        if not tables:
            return None, "לא נמצאה טבלה"

        for table in tables:
            rows = table.find_elements(By.TAG_NAME, 'tr')
            for row in rows:
                cells = row.find_elements(By.TAG_NAME, 'td')
                if cells:
                    row_text = [cell.text.strip() for cell in cells]
                    date_formatted = tx_date.strftime('%d/%m/%Y')
                    if row_text and date_formatted in row_text[0]:
                        if len(row_text) >= 2:
                            try:
                                price_str = row_text[1].replace(',', '')
                                return float(price_str), None
                            except ValueError:
                                return None, f"לא ניתן לפרסר מחיר: {row_text[1]}"

        return None, "התאריך לא נמצא בטבלה"
    except Exception as e:
        return None, f"שגיאה: {str(e)}"


def check_6_tase_prices(rows: list[TxnRow], seed: Optional[int] = None) -> list[PriceCheckResult]:
    """Spec #6 Part 1: Sample transactions of types 12,21,22 and compare with TASE closing price."""
    if not SELENIUM_AVAILABLE:
        print("WARNING: Selenium not available, skipping TASE price checks")
        return []

    results: list[PriceCheckResult] = []
    rng = random.Random(seed)

    # Select samples for each type
    samples_by_type: dict[int, list[TxnRow]] = {}
    for tx_type in TASE_PRICE_CHECK_TYPES:
        type_rows = [r for r in rows if r.tx_type == tx_type and r.security_no and r.tx_date and r.price]
        if type_rows:
            samples_by_type[tx_type] = rng.sample(type_rows, min(TASE_SAMPLES_PER_TYPE, len(type_rows)))

    if not any(samples_by_type.values()):
        return results

    driver = _init_selenium_driver()
    if not driver:
        return results

    try:
        for tx_type, txns in samples_by_type.items():
            for txn in txns:
                closing_price, error = _fetch_tase_closing_price(driver, txn.security_no, txn.tx_date)
                if error:
                    results.append(PriceCheckResult(row=txn, error_message=error))
                else:
                    variance = abs(txn.price - closing_price) / closing_price if closing_price else 0
                    is_exception = variance > TASE_VARIANCE_THRESHOLD
                    results.append(PriceCheckResult(
                        row=txn,
                        tase_closing_price=closing_price,
                        variance_pct=variance * 100,
                        is_exception=is_exception
                    ))
    finally:
        driver.quit()

    return results


def check_6_price_limits(rows: list[TxnRow]) -> list[PriceLimitResult]:
    """Spec #6 Part 2: Check types 31-36 for price > 100."""
    results: list[PriceLimitResult] = []
    for row in rows:
        if row.tx_type in PRICE_LIMIT_TYPES and row.price is not None:
            is_exception = row.price > PRICE_LIMIT
            if is_exception:
                results.append(PriceLimitResult(row=row, is_exception=True))
    return results


# -----------------------------
# Checks (spec #7 - Problematic Securities)
# -----------------------------

def _fetch_problematic_list(driver, list_type: str, url: str) -> set[str]:
    """Fetch a problematic securities list from TASE website."""
    securities: set[str] = set()
    try:
        driver.get(url)
        time.sleep(3)
        try:
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.TAG_NAME, 'table')))
        except:
            pass
        time.sleep(2)

        tables = driver.find_elements(By.TAG_NAME, 'table')
        for table in tables:
            rows = table.find_elements(By.TAG_NAME, 'tr')
            for row in rows:
                cells = row.find_elements(By.TAG_NAME, 'td')
                if len(cells) >= 4:
                    cell_texts = [cell.text.strip() for cell in cells]
                    for text in cell_texts:
                        if re.match(r'^\d{7}$', text):
                            securities.add(text)
                            break
    except Exception as e:
        print(f"    Error fetching {list_type}: {e}")
    return securities


def fetch_problematic_lists(cache_path: Optional[Path] = None) -> dict[str, set[str]]:
    """Fetch all problematic securities lists from TASE."""
    # Try to load from cache
    if cache_path and cache_path.exists():
        try:
            cached = json.loads(cache_path.read_text(encoding='utf-8'))
            return {k: set(v) for k, v in cached.items()}
        except:
            pass

    if not SELENIUM_AVAILABLE:
        print("WARNING: Selenium not available, skipping problematic securities fetch")
        return {}

    driver = _init_selenium_driver()
    if not driver:
        return {}

    all_lists: dict[str, set[str]] = {}
    try:
        for list_type, config in PROBLEMATIC_LISTS_CONFIG.items():
            print(f"  Fetching {config['name_he']}...")
            securities = _fetch_problematic_list(driver, list_type, config['url'])
            all_lists[list_type] = securities
            print(f"    Found {len(securities)} securities")
    finally:
        driver.quit()

    # Save to cache
    if cache_path:
        cache_data = {k: list(v) for k, v in all_lists.items()}
        cache_path.write_text(json.dumps(cache_data, ensure_ascii=False, indent=2), encoding='utf-8')

    return all_lists


def check_7_problematic_securities(rows: list[TxnRow], problematic_lists: dict[str, set[str]]) -> list[ProblematicSecurityResult]:
    """Spec #7: Check all transactions against problematic securities lists."""
    results: list[ProblematicSecurityResult] = []

    for row in rows:
        if not row.security_no:
            continue

        matched_lists = []
        for list_type, security_nos in problematic_lists.items():
            if row.security_no in security_nos:
                matched_lists.append(PROBLEMATIC_LISTS_CONFIG[list_type]['name_he'])

        if matched_lists:
            results.append(ProblematicSecurityResult(
                row=row,
                matched_lists=matched_lists,
                is_exception=True
            ))

    return results


def build_email_json(samples: Samples) -> list[dict[str, Any]]:
    """Spec #5.1: JSON file should contain only two JSON objects with transaction info.

    Output is always a list with 2 items:
      - item 0: decision method 1 sample (or empty object with nulls if unavailable)
      - item 1: decision method 2 sample (or empty object with nulls if unavailable)

    Each object contains:
      Fund number, Fund name, Security name, Security number, Quantity, Price, Date, Type, Decision method
    """
    def txn_obj(row: Optional[TxnRow], decision_method: int) -> dict[str, Any]:
        if row is None:
            return {
                "fund_number": None,
                "fund_name": None,
                "security_name": None,
                "security_number": None,
                "quantity": None,
                "price": None,
                "date": None,
                "type": None,
                "decision_method": decision_method,
            }
        return {
            "fund_number": row.fund_no,
            "fund_name": row.fund_name,
            "security_name": row.security_name,
            "security_number": row.security_no,
            "quantity": row.quantity,
            "price": row.price,
            "date": row.tx_date.strftime("%d-%m-%Y") if row.tx_date else None,
            "type": row.tx_type,
            "decision_method": row.decision_method,
        }

    return [
        txn_obj(samples.decision_1, 1),
        txn_obj(samples.decision_2, 2),
    ]


# -----------------------------
# Output writing (XLSX)
# -----------------------------

def _rtl(ws) -> None:
    ws.sheet_view.rightToLeft = True


def _header(ws, headers: list[str]) -> None:
    for i, h in enumerate(headers, start=1):
        cell = ws.cell(1, i)
        cell.value = h
        cell.font = Font(bold=True)
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions


def _fmt_date(d: Optional[dt.date]) -> str:
    return d.strftime("%d-%m-%Y") if d else ""


def _fmt_time(t: Optional[dt.time]) -> str:
    return t.strftime("%H:%M:%S") if t else ""


def _txn_to_basic_list(r: TxnRow) -> list[Any]:
    return [
        r.fund_no,
        r.fund_name,
        r.security_name,
        r.security_no,
        r.quantity,
        r.price,
        _fmt_date(r.tx_date),
        _fmt_time(r.tx_time),
        r.tx_type,
        r.decision_method,
    ]


def write_output_xlsx(
    output_path: Path,
    *,
    report_month: str,
    summary: dict[str, Any],
    exceptions_duplicates: list[ExceptionRow],
    exceptions_date: list[ExceptionRow],
    exceptions_decision: list[ExceptionRow],
    samples: Samples,
    out_of_scope_funds: dict[int, dict[str, Any]],
    price_check_results: list[PriceCheckResult] = None,
    price_limit_results: list[PriceLimitResult] = None,
    problematic_security_results: list[ProblematicSecurityResult] = None,
) -> None:
    wb = openpyxl.Workbook()

    # Validation columns to add to each sheet
    VALIDATION_COLS = ["שם הבודק", "תוצאת בדיקה"]

    # Summary
    ws_sum = wb.active
    ws_sum.title = "סיכום"
    _rtl(ws_sum)
    _header(ws_sum, ["שדה", "ערך"])
    rr = 2
    for k, v in summary.items():
        ws_sum.cell(rr, 1).value = str(k)
        ws_sum.cell(rr, 2).value = v
        rr += 1

    # Out-of-scope funds (helps validate check #2)
    ws_oos = wb.create_sheet("קרנות מחוץ לתחום")
    _rtl(ws_oos)
    _header(ws_oos, ["מספר קרן", "שם קרן (מהקלט)", "מספר עסקאות", "סיבה"] + VALIDATION_COLS)
    for fid, info in sorted(out_of_scope_funds.items(), key=lambda x: x[0]):
        ws_oos.append([fid, info.get("fund_name"), info.get("count_rows"), info.get("reason"), "", ""])

    # Exceptions - duplicates (inter-fund transactions)
    ws_dup = wb.create_sheet("חריגות - עסקאות בין קרנות")
    _rtl(ws_dup)
    _header(
        ws_dup,
        [
            "בדיקה",
            "סיבה",
            "מפתח קבוצה",
            "מספר קרן",
            "שם קרן",
            "שם נייר",
            "מספר נייר",
            "כמות",
            "מחיר",
            "תאריך",
            "שעה",
            "סוג",
            "אופן החלטה",
            "שורה בקובץ",
        ] + VALIDATION_COLS,
    )
    for ex in exceptions_duplicates:
        ws_dup.append([ex.check_id, ex.reason, ex.group_key, *_txn_to_basic_list(ex.row), ex.row.row_num, "", ""])

    # Exceptions - date
    ws_date = wb.create_sheet("חריגות - תאריך")
    _rtl(ws_date)
    _header(
        ws_date,
        [
            "בדיקה",
            "סיבה",
            "מספר קרן",
            "שם קרן",
            "שם נייר",
            "מספר נייר",
            "כמות",
            "מחיר",
            "תאריך",
            "שעה",
            "סוג",
            "אופן החלטה",
            "שורה בקובץ",
        ] + VALIDATION_COLS,
    )
    for ex in exceptions_date:
        ws_date.append([ex.check_id, ex.reason, *_txn_to_basic_list(ex.row), ex.row.row_num, "", ""])

    # Exceptions - decision method
    ws_dm = wb.create_sheet("חריגות - אופן החלטה")
    _rtl(ws_dm)
    _header(
        ws_dm,
        [
            "בדיקה",
            "סיבה",
            "מספר קרן",
            "שם קרן",
            "שם נייר",
            "מספר נייר",
            "כמות",
            "מחיר",
            "תאריך",
            "שעה",
            "סוג",
            "אופן החלטה",
            "שורה בקובץ",
        ] + VALIDATION_COLS,
    )
    for ex in exceptions_decision:
        ws_dm.append([ex.check_id, ex.reason, *_txn_to_basic_list(ex.row), ex.row.row_num, "", ""])

    # Spec #6 Part 1: TASE price check results
    ws_price = wb.create_sheet("בדיקת מחירים - בורסה")
    _rtl(ws_price)
    _header(
        ws_price,
        [
            "מספר קרן",
            "שם קרן",
            "שם נייר",
            "מספר נייר",
            "כמות",
            "מחיר בעסקה",
            "תאריך",
            "שעה",
            "סוג",
            "מחיר סגירה בורסה",
            "סטייה באחוזים",
            "חריגה",
            "הערה",
        ] + VALIDATION_COLS,
    )
    if price_check_results:
        for r in price_check_results:
            ws_price.append([
                r.row.fund_no,
                r.row.fund_name,
                r.row.security_name,
                r.row.security_no,
                r.row.quantity,
                r.row.price,
                _fmt_date(r.row.tx_date),
                _fmt_time(r.row.tx_time),
                r.row.tx_type,
                r.tase_closing_price,
                f"{r.variance_pct:.2f}%" if r.variance_pct is not None else "",
                "כן" if r.is_exception else "לא",
                r.error_message or "",
                "", ""
            ])

    # Spec #6 Part 2: Price > 100 exceptions
    ws_price_limit = wb.create_sheet("חריגות - מחיר מעל 100")
    _rtl(ws_price_limit)
    _header(
        ws_price_limit,
        [
            "מספר קרן",
            "שם קרן",
            "שם נייר",
            "מספר נייר",
            "כמות",
            "מחיר",
            "תאריך",
            "שעה",
            "סוג",
            "אופן החלטה",
            "שורה בקובץ",
        ] + VALIDATION_COLS,
    )
    if price_limit_results:
        for r in price_limit_results:
            ws_price_limit.append([*_txn_to_basic_list(r.row), r.row.row_num, "", ""])

    # Spec #7: Problematic securities
    ws_prob = wb.create_sheet("חריגות - ניירות בעייתיים")
    _rtl(ws_prob)
    _header(
        ws_prob,
        [
            "מספר קרן",
            "שם קרן",
            "שם נייר",
            "מספר נייר",
            "כמות",
            "מחיר",
            "תאריך",
            "שעה",
            "סוג",
            "אופן החלטה",
            "רשימות בעייתיות",
            "שורה בקובץ",
        ] + VALIDATION_COLS,
    )
    if problematic_security_results:
        for r in problematic_security_results:
            ws_prob.append([
                *_txn_to_basic_list(r.row),
                ", ".join(r.matched_lists),
                r.row.row_num,
                "", ""
            ])

    # Samples
    ws_s = wb.create_sheet("דגימות לבדיקה")
    _rtl(ws_s)
    _header(
        ws_s,
        [
            "קבוצה",
            "מספר קרן",
            "שם קרן",
            "שם נייר",
            "מספר נייר",
            "כמות",
            "מחיר",
            "תאריך",
            "שעה",
            "סוג",
            "אופן החלטה",
            "תאריך החלטת דירקטוריון",
            "סבירות החלטה",
            "ציות לנוהל מנהל",
        ] + VALIDATION_COLS,
    )

    def add_sample(label: str, row: Optional[TxnRow], decision_method: int) -> None:
        if not row:
            ws_s.append([label, None, None, None, None, None, None, None, None, None, decision_method, "", "", "", "", ""])
            return
        ws_s.append([label, *_txn_to_basic_list(row), "", "", "", "", ""])

    add_sample("אופן החלטה = 1", samples.decision_1, 1)
    add_sample("אופן החלטה = 2", samples.decision_2, 2)

    wb.save(output_path)
    wb.close()


# -----------------------------
# CLI
# -----------------------------

def parse_args() -> argparse.Namespace:
    p = argparse.ArgumentParser(description="Mizrahi Special Transactions - CLI processor (checks #1–#7).")
    p.add_argument("--mutual-funds-list", required=True, type=Path, help="Path to 'Mutual Funds List.xlsx'")
    p.add_argument("--input-report", required=True, type=Path, help="Path to manager report CSV or XLSX (e.g., 1702431.csv)")
    p.add_argument("--output-xlsx", required=True, type=Path, help="Path to write output XLSX")
    p.add_argument("--email-json", required=True, type=Path, help="Path to write email JSON payload (two objects only)")
    p.add_argument("--report-month", type=str, default=None, help="Optional report month in YYYY-MM (otherwise inferred from ת.דוח)")
    p.add_argument("--seed", type=int, default=None, help="Optional RNG seed for sampling")
    p.add_argument("--trustee-name", type=str, default=MIZRAHI_TRUSTEE_NAME_DEFAULT, help="Trustee name filter (default: Mizrahi)")
    p.add_argument("--cache-lists", type=Path, default=None, help="Path to cache/load problematic securities lists JSON")
    p.add_argument("--skip-tase-prices", action="store_true", help="Skip TASE price checks (spec #6 part 1)")
    return p.parse_args()


def main() -> int:
    args = parse_args()

    for path in [args.mutual_funds_list, args.input_report]:
        if not path.exists():
            raise SystemExit(f"File not found: {path}")

    # Fund scope = Mizrahi trustee funds only (Maya report removed for now)
    in_scope_funds = load_mizrahi_fund_ids(args.mutual_funds_list, args.trustee_name)
    if not in_scope_funds:
        print("WARNING: Filtered Mizrahi fund list is empty. Check trustee name / input file.")

    rows, meta = load_manager_report(args.input_report)

    report_month = args.report_month or meta.get("report_month_inferred")
    if not report_month:
        raise SystemExit("Could not infer report month from ת.דוח. Provide --report-month YYYY-MM.")

    # Check #1: all rows (inter-fund transactions - same abs quantity, opposite signs)
    ex_dup = check_1_abs_quantity_pairs(rows)

    # Check #2: filter to in-scope rows (fund exists in Mizrahi filtered list)
    in_scope_rows = [r for r in rows if r.fund_no in in_scope_funds]
    out_scope_rows = [r for r in rows if r.fund_no is not None and r.fund_no not in in_scope_funds]

    out_of_scope_funds: dict[int, dict[str, Any]] = {}
    if out_scope_rows:
        counts = Counter([r.fund_no for r in out_scope_rows if r.fund_no is not None])
        for fid, cnt in counts.items():
            out_of_scope_funds[int(fid)] = {
                "count_rows": int(cnt),
                "fund_name": next((r.fund_name for r in out_scope_rows if r.fund_no == fid and r.fund_name), None),
                "reason": "לא ברשימת קרנות מזרחי",
            }

    # Check #3, #4: in-scope only
    ex_date = check_3_dates_in_report_month(in_scope_rows, report_month)
    ex_decision = check_4_decision_method_rules(in_scope_rows)

    # Valid lines for sampling: in-scope rows NOT present in any exception list (including duplicates)
    ex_row_nums = {e.row.row_num for e in (ex_dup + ex_date + ex_decision)}
    valid_rows = [r for r in in_scope_rows if r.row_num not in ex_row_nums]

    # Check #5: sampling
    samples = pick_samples(valid_rows, seed=args.seed)

    # Check #5.1: email JSON (two objects only)
    email_payload = build_email_json(samples)
    args.email_json.parent.mkdir(parents=True, exist_ok=True)
    args.email_json.write_text(json.dumps(email_payload, ensure_ascii=False, indent=2), encoding="utf-8")

    # Check #6 Part 1: TASE price checks (sampled)
    price_check_results: list[PriceCheckResult] = []
    if not args.skip_tase_prices:
        print("Running TASE price checks (spec #6 part 1)...")
        price_check_results = check_6_tase_prices(in_scope_rows, seed=args.seed)
        price_exceptions = [r for r in price_check_results if r.is_exception]
        print(f"  Sampled {len(price_check_results)} transactions, {len(price_exceptions)} exceptions (>5% variance)")

    # Check #6 Part 2: Price > 100 for types 31-36
    print("Running price > 100 check (spec #6 part 2)...")
    price_limit_results = check_6_price_limits(in_scope_rows)
    print(f"  Found {len(price_limit_results)} exceptions with price > 100")

    # Check #7: Problematic securities
    print("Running problematic securities check (spec #7)...")
    problematic_lists = fetch_problematic_lists(cache_path=args.cache_lists)
    for list_type, securities in problematic_lists.items():
        print(f"  {PROBLEMATIC_LISTS_CONFIG[list_type]['name_he']}: {len(securities)} ניירות")
    problematic_security_results = check_7_problematic_securities(in_scope_rows, problematic_lists)
    print(f"  Found {len(problematic_security_results)} transactions with problematic securities")

    summary = {
        "חודש דוח": report_month,
        "סיבת סינון": "קרנות מזרחי בלבד",
        "מספר קרנות מזרחי": len(in_scope_funds),
        "סה\"כ שורות": len(rows),
        "שורות בתחום": len(in_scope_rows),
        "שורות מחוץ לתחום": len(out_scope_rows),
        "חריגות עסקאות בין קרנות": len(ex_dup),
        "חריגות תאריך": len(ex_date),
        "חריגות אופן החלטה": len(ex_decision),
        "חריגות מחיר מעל 100": len(price_limit_results),
        "חריגות ניירות בעייתיים": len(problematic_security_results),
        "שורות תקינות לדגימה": len(valid_rows),
        "דגימה אופן החלטה 1 - שורה": samples.decision_1.row_num if samples.decision_1 else None,
        "דגימה אופן החלטה 2 - שורה": samples.decision_2.row_num if samples.decision_2 else None,
    }

    args.output_xlsx.parent.mkdir(parents=True, exist_ok=True)
    write_output_xlsx(
        args.output_xlsx,
        report_month=report_month,
        summary=summary,
        exceptions_duplicates=ex_dup,
        exceptions_date=ex_date,
        exceptions_decision=ex_decision,
        samples=samples,
        out_of_scope_funds=out_of_scope_funds,
        price_check_results=price_check_results,
        price_limit_results=price_limit_results,
        problematic_security_results=problematic_security_results,
    )

    print("\nסיום.")
    print(json.dumps(summary, ensure_ascii=False, indent=2))
    print(f"קובץ פלט: {args.output_xlsx}")
    print(f"קובץ JSON: {args.email_json}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
