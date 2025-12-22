#!/usr/bin/env python3
"""
Mizrahi Special Transactions - CLI Processor (single script)

Implements checks #1–#5.1 from the current specification.

Inputs:
  - Mutual Funds List (XLSX): filtered to Mizrahi trustee funds (by 'שם נאמן')
  - Manager special transactions report (XLSX): e.g., 1702431.xlsx

Outputs:
  - Output XLSX: summary + exceptions + samples (+ out-of-scope funds)
  - Email JSON: for n8n workflow - contains ONLY two JSON objects with transaction info (no full email body)

Dependencies:
  - Python 3.10+
  - openpyxl

Example:
  python mizrahi_special_transactions.py \
    --mutual-funds-list "Mutual Funds List.xlsx" \
    --input-report "1702431.xlsx" \
    --output-xlsx "output.xlsx" \
    --email-json "email.json" \
    --seed 123

Notes:
  - Hebrew column headers are expected (as in your provided files).
  - Dates in the manager report are often stored as numbers DDMMYYYY without leading zeros.
  - Times are often stored as numbers HHMMSS without leading zeros.
"""

from __future__ import annotations

import argparse
import datetime as dt
import json
import random
from collections import Counter, defaultdict
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Optional

import openpyxl
from openpyxl.styles import Font


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


def load_manager_report(input_report_path: Path) -> tuple[list[TxnRow], dict]:
    """Load manager special-transactions report. Expects headers in row 1."""
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
    """Spec #1: within unique_id(security+date), if there are two rows with different quantities but identical abs() -> flag."""
    by_uid: dict[str, list[TxnRow]] = defaultdict(list)
    for r in rows:
        by_uid[r.unique_id].append(r)

    out: list[ExceptionRow] = []
    for uid, group in by_uid.items():
        abs_map: dict[float, list[TxnRow]] = defaultdict(list)
        for r in group:
            if r.quantity is None:
                continue
            abs_map[abs(r.quantity)].append(r)

        for abs_qty, rs in abs_map.items():
            qty_values = {r.quantity for r in rs if r.quantity is not None}
            if len(qty_values) > 1:
                for r in rs:
                    out.append(ExceptionRow(check_id="CHK_1", reason="QTY_ABS_MATCH_DIFFERENT", row=r, group_key=f"{uid}|abs={abs_qty}"))

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
) -> None:
    wb = openpyxl.Workbook()

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
    _header(ws_oos, ["מספר קרן", "שם קרן (מהקלט)", "מספר עסקאות", "סיבה"])
    for fid, info in sorted(out_of_scope_funds.items(), key=lambda x: x[0]):
        ws_oos.append([fid, info.get("fund_name"), info.get("count_rows"), info.get("reason")])

    # Exceptions - duplicates
    ws_dup = wb.create_sheet("חריגות - כפילויות")
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
        ],
    )
    for ex in exceptions_duplicates:
        ws_dup.append([ex.check_id, ex.reason, ex.group_key, *_txn_to_basic_list(ex.row), ex.row.row_num])

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
        ],
    )
    for ex in exceptions_date:
        ws_date.append([ex.check_id, ex.reason, *_txn_to_basic_list(ex.row), ex.row.row_num])

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
        ],
    )
    for ex in exceptions_decision:
        ws_dm.append([ex.check_id, ex.reason, *_txn_to_basic_list(ex.row), ex.row.row_num])

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
        ],
    )

    def add_sample(label: str, row: Optional[TxnRow], decision_method: int) -> None:
        if not row:
            ws_s.append([label, None, None, None, None, None, None, None, None, None, decision_method, "", "", ""])
            return
        ws_s.append([label, *_txn_to_basic_list(row), "", "", ""])

    add_sample("Decision method = 1", samples.decision_1, 1)
    add_sample("Decision method = 2", samples.decision_2, 2)

    wb.save(output_path)
    wb.close()


# -----------------------------
# CLI
# -----------------------------

def parse_args() -> argparse.Namespace:
    p = argparse.ArgumentParser(description="Mizrahi Special Transactions - CLI processor (checks #1–#5.1).")
    p.add_argument("--mutual-funds-list", required=True, type=Path, help="Path to 'Mutual Funds List.xlsx'")
    p.add_argument("--input-report", required=True, type=Path, help="Path to manager report XLSX (e.g., 1702431.xlsx)")
    p.add_argument("--output-xlsx", required=True, type=Path, help="Path to write output XLSX")
    p.add_argument("--email-json", required=True, type=Path, help="Path to write email JSON payload (two objects only)")
    p.add_argument("--report-month", type=str, default=None, help="Optional report month in YYYY-MM (otherwise inferred from ת.דוח)")
    p.add_argument("--seed", type=int, default=None, help="Optional RNG seed for sampling")
    p.add_argument("--trustee-name", type=str, default=MIZRAHI_TRUSTEE_NAME_DEFAULT, help="Trustee name filter (default: Mizrahi)")
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

    # Check #1: all rows
    ex_dup = check_1_duplicates_exact(rows) + check_1_abs_quantity_pairs(rows)

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
                "reason": "not_in_mizrahi_filtered_list",
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

    summary = {
        "report_month": report_month,
        "scope_reason": "mizrahi_only",
        "mizrahi_funds_count": len(in_scope_funds),
        "rows_total": len(rows),
        "rows_in_scope": len(in_scope_rows),
        "rows_out_of_scope": len(out_scope_rows),
        "exceptions_duplicates": len(ex_dup),
        "exceptions_date": len(ex_date),
        "exceptions_decision_method": len(ex_decision),
        "valid_rows_for_sampling": len(valid_rows),
        "sample_dm1_row": samples.decision_1.row_num if samples.decision_1 else None,
        "sample_dm2_row": samples.decision_2.row_num if samples.decision_2 else None,
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
    )

    print("Done.")
    print(json.dumps(summary, ensure_ascii=False, indent=2))
    print(f"Output XLSX: {args.output_xlsx}")
    print(f"Email JSON:  {args.email_json}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
