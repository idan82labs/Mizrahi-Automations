# Mizrahi Fund Automation

Automated fund report processing and validation system for Mizrahi Tefahot trustee services. This repository contains multiple automation tools for regulatory compliance checks on mutual fund data from Maya (TASE - Tel Aviv Stock Exchange).

## Automations

| Automation | Description | Script |
|------------|-------------|--------|
| **Fund Holdings Compliance** | Cross-references funds, detects unusual assets, validates regulatory clauses | `fund_automation_complete.py` |
| **Special Transactions** | Validates special transaction reports (checks 1-7) | `mizrahi_special_transactions.py` |
| **K.303 Disclosure Validation** | Validates K.303 disclosure reports from Maya | `k.303 validation (Mizrahi_5)/` |

---

## 1. Fund Holdings Compliance

Fetches fund data from Maya via Apify actors and performs comprehensive compliance checks.

### Features

- **Automated Data Fetching**: Retrieves fund lists and manager reports from Maya/TASE using Apify actors
- **Fund Completeness Check**: Cross-references funds between Magna list and manager reports
- **Unusual Asset Detection**: Flags holdings with unusual asset types (types 16, 21, 22, 23, 24, 52, 53, 57, 58, 99, 101, 112, 201, 207, 209)
- **Change Tracking**: Identifies new assets and quantity changes from previous month
- **Regulatory Compliance**: Validates Clause 214 (variable management fees) and Clause 328 (borrowed quantities)
- **Required Combinations**: Verifies required asset type pairings exist
- **Price Reasonableness**: Checks price ratios with 7.5% threshold

### Usage

```bash
python fund_automation_complete.py --fund-name "סיגמא"
python fund_automation_complete.py --fund-name "סיגמא" --output-dir ./reports
python fund_automation_complete.py --fund-name "סיגמא" --keep-temp
```

### Output Sheets

| Sheet | Description |
|-------|-------------|
| סיכום | Summary statistics |
| סטטוס בדיקות | Check status overview |
| קרנות חסרות | Missing funds (if any) |
| נכסים חריגים | Unusual assets |
| נכסים חדשים | New assets |
| שינויים בכמות | Quantity changes |
| סעיף 328 | Clause 328 issues |
| שילובים נדרשים | Required combinations (including Clause 214) |
| סבירות מחירים | Price reasonableness issues |

---

## 2. Special Transactions Validation

Validates special transaction reports with comprehensive checks for inter-fund transactions, date validation, decision methods, TASE price comparison, and problematic securities.

### Features

- **Check 1**: Inter-fund transaction detection
- **Check 3**: Date validation (transaction dates within reporting period)
- **Check 4**: Decision method rule validation
- **Check 4C**: Price/type consistency for same security+date+time
- **Check 6**: TASE price comparison and price limit checks (>100)
- **Check 7**: Problematic securities detection

### Usage

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

### Output

- **Excel Report**: Summary + check statuses + exceptions + samples
- **Email JSON**: For n8n workflow integration

---

## 3. K.303 Disclosure Validation

Validates K.303 disclosure reports ("גילוי נאות") from Maya with current vs. previous month comparison.

### Features

- **Check 1א**: Fund completeness - all Mizrahi funds present in report
- **Check 1ב**: Date validity - report dates match expected month
- **Check 2א**: Previous month comparison - significant changes detection
- **Check 2ב**: Exposure profile validation
- **Checks 3א-ח**: Various K.303-specific validations

### Directory Structure

```
k.303 validation (Mizrahi_5)/
├── disclosure_k303_validator.py  # Standalone validator
├── k303_automation_complete.py   # Full automation (Apify + validation)
├── k303_code_index.xlsx          # Hierarchical code definitions
├── k303_spec.xlsx                # Default spec file
├── k303_maya_downloader/         # Apify actor for Maya scraping
└── output/                       # Default output directory
```

### Usage

**Full Automation** (fetches from Maya via Apify + validates):
```bash
export APIFY_TOKEN="your_token_here"
python "k.303 validation (Mizrahi_5)/k303_automation_complete.py" --fund-name "מגדל"
python "k.303 validation (Mizrahi_5)/k303_automation_complete.py" --fund-name "מגדל" --keep-temp
```

**Standalone Validation** (with local files):
```bash
python "k.303 validation (Mizrahi_5)/disclosure_k303_validator.py" \
    --mutual-funds-list "Mutual_Funds_List.xlsx - Worksheet.csv" \
    --current-report "disclosure_current_month.csv" \
    --previous-report "disclosure_previous_month.csv" \
    --output-xlsx "validation_output.xlsx" \
    --report-month "2025-12" \
    --manager-name "מגדל"
```

---

## Supported Fund Managers

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

## Installation

```bash
pip install pandas openpyxl requests
```

## Configuration

Set Apify API token via environment variable:
```bash
export APIFY_TOKEN="your_token_here"
```

## Language & Encoding

- Hebrew is the primary language for fund names, check names, and output labels
- Files use UTF-8 BOM encoding
- Handle encoding carefully with Maya exports (cp1255, UTF-8)

## License

Proprietary - Mizrahi Tefahot
