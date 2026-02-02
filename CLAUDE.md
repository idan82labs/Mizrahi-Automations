# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

Mizrahi Fund Automation is a Python-based financial compliance system for Mizrahi Tefahot trustee services. It fetches mutual fund data from Maya (TASE - Tel Aviv Stock Exchange) via Apify actors and performs regulatory compliance checks on fund holdings.

## Commands

### Installation
```bash
pip install pandas openpyxl requests
```

### Running Scripts

**Fund completeness and compliance checks:**
```bash
python fund_automation_complete.py --fund-name "סיגמא"
python fund_automation_complete.py --fund-name "סיגמא" --output-dir ./reports --keep-temp
```

**Special transactions validation:**
```bash
python mizrahi_special_transactions.py \
    --mutual-funds-list "Mutual Funds List.xlsx" \
    --input-report "1702431.csv" \
    --output-xlsx "output.xlsx" \
    --email-json "email.json" \
    --manager-name "איילון" \
    --report-month "2025-12" \
    --max-exceptions 50
```

**K.303 disclosure report validation (standalone):**
```bash
python "k.303 validation (Mizrahi_5)/disclosure_k303_validator.py" \
    --mutual-funds-list "Mutual_Funds_List.xlsx - Worksheet.csv" \
    --current-report "disclosure_migdal_current_month.csv" \
    --previous-report "disclosure_migdal_previous_month.xlsx - Sheet1.csv" \
    --output-xlsx "disclosure_validation_output.xlsx" \
    --report-month "2025-12" \
    --manager-name "מגדל"
```

**K.303 full automation (fetches from Apify + validates):**
```bash
python "k.303 validation (Mizrahi_5)/k303_automation_complete.py" --fund-name "מגדל"
python "k.303 validation (Mizrahi_5)/k303_automation_complete.py" --fund-name "מגדל" --output-dir ./reports --keep-temp
```

**Apify integration test:**
```bash
python test_fund_automation.py
```

## Architecture

### Directory Structure

```
├── fund_automation_complete.py      # Fund holdings compliance (root level)
├── mizrahi_special_transactions.py  # Special transactions validation
├── test_fund_automation.py          # Apify integration test
└── k.303 validation (Mizrahi_5)/                       # K.303 disclosure module
    ├── disclosure_k303_validator.py # K.303 validation logic
    ├── k303_automation_complete.py  # End-to-end K.303 automation
    ├── k303_code_index.xlsx         # Hierarchical code definitions
    ├── k303_spec.xlsx               # Default spec file for checks
    └── k303_maya_downloader/        # Apify actor (Playwright/JS)
        └── main.js                  # Scrapes Maya for K.303 CSV files
```

### Data Flow Pattern
```
CSV/XLSX Input → Load & Parse → Normalize & Filter → Apply Checks → Aggregate Results → Excel Output
```

### Key Patterns

**Dataclass-based entities** - All scripts use Python dataclasses for structured data (`TxnRow`, `ExceptionRow`, `CheckStatus`, `ProcessingResult`).

**Apify integration** - REST API with polling loop (3s intervals, 180s timeout).

| Actor ID | Purpose |
|----------|---------|
| `K9WppTziYC3n2vxTu` | Funds list |
| `5lhI6O39Qbgv9O0gs` | Fund reports (holdings) |
| `iTpNz9ixbdQCmH43C` | K.303 disclosure reports (Maya scraper) |

**Hierarchical logging** - Each run creates a timestamped directory under `log/` with separate log files per check.

**Smart sampling** - When exceptions exceed threshold (default 100), stratified sampling preserves representation of all exception types.

**Configuration by constants** - Fund manager codes, asset types, and thresholds are hardcoded at module level (e.g., `FUND_MANAGER_CODES`, `UNUSUAL_ASSET_TYPES`, `REQUIRED_COMBINATIONS`).

### K.303 Code Index

The K.303 validator uses a hierarchical code system (levels 1-4) defined in `k.303 validation (Mizrahi_5)/k303_code_index.xlsx`. Codes range from 2-digit (e.g., `01` = מניות) to 8-digit detail levels (e.g., `03010101` = ממשלתי צמוד מדד).

### Supported Fund Managers

| Hebrew Name | Code |
|------------|------|
| מגדל (Migdal) | 10040 |
| איילון (Ayalon) | 10054 |
| קסם (Kesem) | 10047 |
| סיגמא (Sigma) | 10048 |
| פורסט (Forest) | 10082 |
| הראל (Harel) | 10031 |
| אנליסט (Analyst) | 10019 |
| מיטב (Meitav) | 10083 |
| איביאי (IBI) | 10068 |
| אלטשולר-שחם (Altshuler Shaham) | 10017 |

## Configuration

Set Apify API token via environment variable (preferred) or in script:
```bash
export APIFY_TOKEN="your_token_here"
```

## Language & Encoding

- Hebrew is the primary language for fund names, check names, and output labels
- Handle encoding carefully: UTF-8 BOM, cp1255 (Windows-1255), and potential 0x10 offset encoding quirks
- Use `fix_shifted_encoding()` for Hebrew text issues from Maya exports

## Output

- Excel files with multiple sheets (סיכום, סטטוס בדיקות, etc.)
- JSON output for n8n workflow integration (email notifications)
- Structured logs in `log/` directory with run-scoped subdirectories
