# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

Mizrahi Fund Automation is a Python-based financial compliance system for Mizrahi Tefahot trustee services. It fetches mutual fund data from Maya (TASE - Tel Aviv Stock Exchange) via Apify actors and performs regulatory compliance checks on fund holdings.

## Commands

### Installation
```bash
pip install pandas openpyxl requests python-dotenv
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

**Daily tracking full automation (Mizrahi_4 - fetches from Apify + generates output):**
```bash
python Mizrahi_4/mizrahi_4_automation.py
python Mizrahi_4/mizrahi_4_automation.py --output-dir ./reports --keep-temp
python Mizrahi_4/mizrahi_4_automation.py --skip-indx  # Skip INDX data fetch
```

**Daily tracking data generation (Mizrahi_4 - standalone logic):**
```bash
python Mizrahi_4/mizrahi_4_logic.py \
    --mutual-funds-list "Mutual_Funds_List.xlsx - Worksheet.csv" \
    --fund-index-table "Mizrahi_4/input/fund-index-table.xlsx" \
    --bfix-prices "Mizrahi_4/input/BFIX PRICE BLOOMBERG.xlsx" \
    --bloomberg-index "Mizrahi_4/input/bloomberg-index.xlsx" \
    --output-xlsx "Mizrahi_4/output/output.xlsx"
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
├── Mizrahi_4/                       # Daily tracking data module
│   ├── mizrahi_4_automation.py      # Full automation (Apify + logic)
│   ├── mizrahi_4_logic.py           # Daily tracking generator
│   ├── input/                       # Input files
│   │   ├── fund-index-table.xlsx    # Fund-to-index mapping (86 funds)
│   │   ├── BFIX PRICE BLOOMBERG.xlsx # Bloomberg BFIX currency rates
│   │   └── bloomberg-index.xlsx     # Bloomberg index prices
│   ├── output/                      # Output files
│   └── indx_downloader/             # Apify actor (Playwright/JS)
│       └── main.js                  # Scrapes indx.co.il for historical index data
└── k.303 validation (Mizrahi_5)/    # K.303 disclosure module
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
| `P9tr210PVi6W8RvtU` | INDX historical index data (indx_downloader) |

### Mizrahi_4 Daily Tracking

The Mizrahi_4 module generates daily tracking workbooks for funds with variable management fees. Scope criteria:
- Trustee must be "מזרחי טפחות"
- Fund must have non-zero variable management fees (דמי ניהול משתנים)
- Fund must exist in `fund-index-table.xlsx` mapping

Output: One Excel sheet per fund with 31 base columns (date, index, fees, rates, etc.) from Jan 1 of current year to today.

**BFIX Price Integration**: When `--bfix-prices` is provided, BFIX currency rates from Bloomberg are appended as additional columns:
- Each fund's BFIX codes are read from column F ("סוג מטבע (קוד bfix)") in `fund-index-table.xlsx`
- All matching columns from `BFIX PRICE BLOOMBERG.xlsx` (sheet "ערכים") are appended
- Column headers use format: "Description (Code)" e.g., "שער 17:30 - כל הימים BFIX USD (ILS F103 Curncy)"
- Dates are matched between fund sheet and BFIX data
- Missing data is flagged as #N/A error with red font

**Bloomberg Index Integration**: When `--bloomberg-index` is provided, funds with "בלומברג" data source (column D in `fund-index-table.xlsx`) get actual index prices in column D:
- Index codes (מספר מדד) are matched to Bloomberg data using case-insensitive matching on the first part of the code
- Column D header changes from "מדד" to "מדד (FULL_INDEX_CODE)" e.g., "מדד (M1WO INDEX)"
- Values are populated from `bloomberg-index.xlsx` (sheet "ערכים") for each date
- Missing dates or values show #N/A error with red font
- Funds without "בלומברג" data source continue to show index ID as text placeholder

**INDX Index Integration**: For funds with "אינדקס" data source (column D in `fund-index-table.xlsx`):
- Historical index data is downloaded from indx.co.il using the `indx_downloader` Apify actor
- Index page URLs are stored in column C (e.g., `https://indx.co.il/index/2123-index/`)
- Downloaded files use pattern: `{index_id}_Historical_Data.xlsx`
- Actor input: `{ "indexUrls": ["https://indx.co.il/index/2123-index/", ...] }`

**TASE Index Integration**: For funds with "מאי"ה" data source (column D in `fund-index-table.xlsx`):
- Index data is fetched from TASE Data Hub API (`/v1/indices/eod/history/five-years/by-index`)
- TASE index IDs are stored in column B (e.g., `142` for TA-35, `724` for Tel Bond-Shekel 1-3)
- Index names are stored in column C (e.g., "תל בונד-שקלי 1-3") - used for column D header
- Uses `closingIndexPrice` field from API response
- Enabled by default; API key from `TASE_API_KEY` env var (set in `.env`)
- Skip with `--skip-tase-data` flag
- Rate limiting: 10 requests per 2 seconds (automatic 0.25s delay between requests)

**Hierarchical logging** - Each run creates a timestamped directory under `log/` with separate log files per check.

**Smart sampling** - When exceptions exceed threshold (default 100), stratified sampling preserves representation of all exception types.

**Configuration by constants** - Fund manager codes, asset types, and thresholds are hardcoded at module level (e.g., `FUND_MANAGER_CODES`, `UNUSUAL_ASSET_TYPES`, `REQUIRED_COMBINATIONS`).

### K.303 Code Index

The K.303 validator uses a hierarchical code system (levels 1-4) defined in `k.303 validation (Mizrahi_5)/k303_code_index.xlsx`. Codes range from 2-digit (e.g., `01` = מניות) to 8-digit detail levels (e.g., `03010101` = ממשלתי צמוד מדד).

### Supported Fund Managers

| Hebrew Name | Code |
|------------|------|
| מגדל (Migdal) | 10040 |
| קסם (Kesem) | 10047 |
| סיגמא (Sigma) | 10048 |
| הראל (Harel) | 10031 |
| אנליסט (Analyst) | 10019 |
| מיטב (Meitav) | 10083 |
| איביאי (IBI) | 10068 |
| אלטשולר-שחם (Altshuler Shaham) | 10017 |

## Configuration

### Environment Variables

Copy `.env.example` to `.env` and fill in your API keys:

```bash
cp .env.example .env
```

**.env file:**
```bash
# Apify API token (for Maya data fetching)
APIFY_TOKEN=your_apify_token_here

# TASE Data Hub API key (for index values)
TASE_API_KEY=your_tase_api_key_here
```

**Load in Python:**
```python
from dotenv import load_dotenv
import os

load_dotenv()
TASE_API_KEY = os.environ["TASE_API_KEY"]
APIFY_TOKEN = os.environ["APIFY_TOKEN"]
```

**Install python-dotenv:**
```bash
pip install python-dotenv
```

## TASE Data Hub API

The TASE Data Hub is a REST API provided by the Tel Aviv Stock Exchange for accessing capital market data including securities, indices, mutual funds, and public companies information.

### Quick Reference

| Item | Value |
|------|-------|
| Base URL | `https://datawise.tase.co.il` |
| Developers Portal | https://datahubportal.tase.co.il/login |
| API Documentation | https://datahubapi.tase.co.il/docs |
| Product Catalog | https://www.tase.co.il/he/content/products_lobby/datahub |
| API Guide (PDF) | `Mizrahi_4/docs/2000_api_guide_eng.pdf` |

### Authentication

API Key authentication via HTTP header:
```
apikey: YOUR_API_KEY
```

To obtain an API Key:
1. Register at the Developers Portal
2. Create an App in "My Apps" menu
3. Click "+Generate Credential" to create an API Key
4. Store the key securely (cannot be recovered once lost)

### Required Headers

```http
accept: application/json
accept-language: he-IL
apikey: YOUR_API_KEY
```

### Example Request (curl)

```bash
curl --request GET \
  --url 'https://datahubapi.tase.co.il/board-and-management/positions?issuerId=76' \
  --header 'accept: application/json' \
  --header 'accept-language: he-IL' \
  --header 'apikey: YOUR_API_KEY'
```

### Rate Limits

| Limit Type | Value |
|------------|-------|
| Rate Limit | 10 requests per 2 seconds |
| Burst Limit | 10 requests per 2 seconds |

Exceeding limits returns `HTTP 429 - Too Many Requests`.

### Contact

| Purpose | Email |
|---------|-------|
| API Support | apisupport@tase.co.il |
| Paid Products / Data Sales | marketdatateam@tase.co.il |

**Note**: Some data products require commercial approval prior to activation. Registration requests do not automatically grant access to paid products.

### Indices Data EoD - Five Years Back (Product)

This product provides EOD (End of Day) data for TASE indices for the last 5 years.

**Local Documentation**:
- `Mizrahi_4/docs/tase-indices-5years-openapi.yaml` - Raw OpenAPI spec

#### Endpoints

| Method | Endpoint | Description |
|--------|----------|-------------|
| GET | `/v1/indices/eod/history/five-years/by-date` | Index data for a specific date |
| GET | `/v1/indices/eod/history/five-years/by-index` | Index data for a specific index across dates |

---

#### GET by-index (Primary endpoint for index values)

**Endpoint**: `GET /v1/indices/eod/history/five-years/by-index`

Returns end-of-day prices and trading data for a specific TASE index over the past five years.

**Parameters**:

| Parameter | Type | Location | Required | Description |
|-----------|------|----------|----------|-------------|
| `indexId` | integer | query | **Yes** | Index Id, must be between 1 to 999 |
| `fromDate` | string | query | **Yes** | First date to filter by (format: `YYYY-MM-DD`), must be in range of last 5 years |
| `toDate` | string | query | No | Last date to filter by (format: `YYYY-MM-DD`) |

**Example Request**:
```bash
curl --request GET \
  --url 'https://datawise.tase.co.il/v1/indices/eod/history/five-years/by-index?indexId=142&fromDate=2024-01-01&toDate=2024-12-31' \
  --header 'accept: application/json' \
  --header 'accept-language: he-IL' \
  --header 'apikey: YOUR_API_KEY'
```

**Response Schema** (200 OK):
```json
{
  "indexEndOfDay": {
    "result": [
      {
        "tradeDate": "2023-03-01T00:00:00",
        "indexId": "142",
        "isin": "IL0004250176",
        "baseIndexPrice": 1260.33,
        "indexOpeningPrice": 1260.33,
        "high": 353.26,
        "low": 324.25,
        "overallMarketCap": 895057151.5,
        "closingIndexPrice": 348.76
      }
    ],
    "total": 50
  }
}
```

**Response Fields**:

| Field | Type | Description |
|-------|------|-------------|
| `tradeDate` | datetime | Trading date |
| `indexId` | string | TASE index identifier |
| `isin` | string | International Securities ID (IL...) |
| `baseIndexPrice` | number | Index base rate |
| `indexOpeningPrice` | number | Index opening rate |
| `high` | number | Highest index rate |
| `low` | number | Lowest index rate |
| `overallMarketCap` | number | Overall market cap in NIS |
| `closingIndexPrice` | number | **Index closing rate (use this for index values)** |

**Error Responses**:

| Code | Message |
|------|---------|
| 400 | Index is not a number |
| 429 | Too Many Requests (rate limit exceeded) |

---

#### GET by-date

**Endpoint**: `GET /v1/indices/eod/history/five-years/by-date`

Returns end-of-day prices and trading data for all TASE indices on a specific date.

**Parameters**:

| Parameter | Type | Location | Required | Description |
|-----------|------|----------|----------|-------------|
| `date` | string | query | **Yes** | Date to filter by (format: `YYYY-MM-DD`), must be in range of last 5 years |
| `indexId` | integer | query | No | Index Id to filter by, must be between 1 to 999 |

**Example Request**:
```bash
curl --request GET \
  --url 'https://datawise.tase.co.il/v1/indices/eod/history/five-years/by-date?date=2024-01-15' \
  --header 'accept: application/json' \
  --header 'accept-language: he-IL' \
  --header 'apikey: YOUR_API_KEY'
```

## Language & Encoding

- Hebrew is the primary language for fund names, check names, and output labels
- Handle encoding carefully: UTF-8 BOM, cp1255 (Windows-1255), and potential 0x10 offset encoding quirks
- Use `fix_shifted_encoding()` for Hebrew text issues from Maya exports

## Output

- Excel files with multiple sheets (סיכום, סטטוס בדיקות, etc.)
- JSON output for n8n workflow integration (email notifications)
- Structured logs in `log/` directory with run-scoped subdirectories
