# Mizrahi Fund Automation

Automated fund report processing and validation system for Mizrahi Tefahot trustee services. This tool fetches fund data from Maya (TASE) via Apify actors and performs comprehensive compliance checks.

## Features

- **Automated Data Fetching**: Retrieves fund lists and manager reports from Maya/TASE using Apify actors
- **Fund Completeness Check**: Cross-references funds between Magna list and manager reports
- **Unusual Asset Detection**: Flags holdings with unusual asset types (types 16, 21, 22, 23, 24, 52, 53, 57, 58, 99, 101, 112, 201, 207, 209)
- **Change Tracking**: Identifies new assets and quantity changes from previous month
- **Regulatory Compliance**: Validates Clause 214 (variable management fees) and Clause 328 (borrowed quantities)
- **Required Combinations**: Verifies required asset type pairings exist
- **Price Reasonableness**: Checks price ratios with 7.5% threshold

## Supported Fund Managers

- מגדל (Migdal)
- איילון (Ayalon)
- קסם (Kesem)
- סיגמא (Sigma)
- פורסט (Forest)
- הראל (Harel)
- אנליסט (Analyst)
- מיטב (Meitav)
- איביאי (IBI)
- אלטשולר-שחם (Altshuler Shaham)

## Installation

```bash
pip install pandas openpyxl requests
```

## Configuration

Before running, set your Apify API token in the script:

```python
APIFY_TOKEN = "YOUR_APIFY_TOKEN_HERE"
```

## Usage

### Full Automation (Complete Pipeline)

```bash
python fund_automation_complete.py --fund-name "סיגמא"
python fund_automation_complete.py --fund-name "סיגמא" --output-dir ./reports
python fund_automation_complete.py --fund-name "סיגמא" --keep-temp  # Keep temp files
```

### Test Script (Apify Integration Test)

```bash
python test_fund_automation.py
```

This will:
1. Fetch the master funds list from Apify
2. Fetch fund reports (current and previous month CSVs)
3. Save files to `test_output/` directory

## Output

The main script generates an Excel report with the following sheets:

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

## Checks Performed

1. **Fund Completeness** - Cross-reference Magna funds vs manager report
2. **Unusual Asset Types** - Flag holdings with non-standard asset types
3. **New Assets** - Identify assets added since previous month
4. **Quantity Changes** - Track changes in unusual asset quantities
5. **Clause 328** - Verify borrowed quantity consistency
6. **Required Combinations** - Validate asset type pairings (111, 212, 213, 208, 210)
7. **Price Reasonableness** - Check price ratios (300/301, 314/313, 316/315)

---

## Special Transactions Report Processor

`mizrahi_special_transactions.py` is a CLI tool for processing and validating manager special transactions reports. It implements compliance checks #1–#7 from the specification.

### Features

- **Inter-Fund Transaction Detection** (Check #1): Identifies transactions between funds
- **Date Validation** (Check #3): Validates transaction dates against report period
- **Decision Method Rules** (Check #4): Verifies decision method compliance for specific transaction types
- **TASE Price Comparison** (Check #6): Compares transaction prices against TASE prices with configurable threshold
- **Price Limit Validation** (Check #6): Flags transactions with price > 100 for specific types
- **Problematic Securities Detection** (Check #7): Cross-references transactions against TASE problematic lists (low liquidity, maintenance, suspended)

### Dependencies

```bash
pip install openpyxl selenium webdriver-manager
```

Note: Selenium is optional and only required for TASE price checks (spec #6).

### Command-Line Arguments

| Argument | Required | Description |
|----------|----------|-------------|
| `--mutual-funds-list` | Yes | Path to 'Mutual Funds List.xlsx' |
| `--input-report` | Yes | Path to manager report CSV or XLSX (e.g., 1702431.csv) |
| `--output-xlsx` | Yes | Path to write output XLSX report |
| `--email-json` | Yes | Path to write email JSON payload for n8n workflow |
| `--report-month` | No | Report month in YYYY-MM format (otherwise inferred from ת.דוח column) |
| `--seed` | No | RNG seed for reproducible sampling |
| `--trustee-name` | No | Trustee name filter (default: מזרחי טפחות חברה לנאמנות בע"מ) |
| `--manager-name` | No | Fund manager name for report header |
| `--spec-file` | No | Path to specification table Excel file (for פירוט בדיקות sheet) |
| `--cache-lists` | No | Path to cache/load problematic securities lists JSON |
| `--skip-tase-prices` | No | Skip TASE price checks (spec #6 part 1) |
| `--price-threshold` | No | Price variance threshold in percent for TASE price checks (default: 5.0%) |

### Usage Examples

Basic usage with all required arguments:
```bash
python mizrahi_special_transactions.py \
  --mutual-funds-list "Mutual Funds List.xlsx" \
  --input-report "1702431.csv" \
  --output-xlsx "output.xlsx" \
  --email-json "email.json"
```

Full usage with all options:
```bash
python mizrahi_special_transactions.py \
  --mutual-funds-list "Mutual Funds List.xlsx" \
  --input-report "1702431.csv" \
  --output-xlsx "output.xlsx" \
  --email-json "email.json" \
  --manager-name "איילון" \
  --spec-file "Special Transactions Report Testing Specifications.xlsx" \
  --price-threshold 5.0 \
  --seed 123
```

Skip TASE price checks (faster execution):
```bash
python mizrahi_special_transactions.py \
  --mutual-funds-list "Mutual Funds List.xlsx" \
  --input-report "1702431.csv" \
  --output-xlsx "output.xlsx" \
  --email-json "email.json" \
  --skip-tase-prices
```

### Output

The script generates an Excel report with the following sheets:

| Sheet | Description |
|-------|-------------|
| סיכום | Summary statistics |
| סטטוס בדיקות | Check status overview |
| קרנות מחוץ לתחום | Out-of-scope funds |
| חריגות - עסקאות בין קרנות | Inter-fund transaction exceptions |
| חריגות - תאריך | Date validation exceptions |
| חריגות - אופן החלטה | Decision method exceptions |
| בדיקת מחירים - בורסה | TASE price comparison results |
| חריגות - מחיר מעל 100 | Price > 100 exceptions |
| חריגות - ניירות בעייתיים | Problematic securities exceptions |
| דגימות לבדיקה | Samples for manual review |
| פירוט בדיקות | Specification details (if --spec-file provided) |

### Logging

The script creates detailed logs in the `log/` directory with a unique folder per run:
- `main.log` - General processing log
- `chk1_inter_fund.log` - Inter-fund transaction checks
- `chk3_date.log` - Date validation checks
- `chk4_decision.log` - Decision method rule checks
- `chk6_tase_price.log` - TASE price comparison checks
- `chk6_price_limit.log` - Price > 100 checks
- `chk7_problematic.log` - Problematic securities checks

---

## License

Proprietary - Mizrahi Tefahot
