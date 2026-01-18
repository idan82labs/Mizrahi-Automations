# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Overview

This is a validation system for special transactions (עסקאות מיוחדות) in mutual funds managed under Mizrahi Tefahot trusteeship. The system validates transaction reports from 10 fund managers against a specification of 7 validation checks, producing Excel reports and email notifications.

## Core Architecture

### Main Validation Script: `mizrahi_special_transactions.py`

Single-file CLI processor implementing 7 validation checks (CHK_1 through CHK_7), with CHK_4 split into 4 sub-checks:

1. **CHK_1**: Duplicate detection (exact duplicates + inter-fund transactions)
2. **CHK_2**: In-scope fund filtering (Mizrahi trustee funds only)
3. **CHK_3**: Date validation (transactions must be in report month)
4. **CHK_4**: Decision method controls (4 independent sub-checks):
   - **CHK_4D**: Missing data validation (transaction_type and decision_method presence)
   - **CHK_4A**: Type-decision compatibility (types 12,22 require decision=1; types 31-36 require decision=1 or 2)
   - **CHK_4B**: דח"צ opposition check (flag when any דחצ field = 2)
   - **CHK_4C**: דח"צ1 no decision check (flag when דחצ1 = 0, only if 4B didn't flag)
5. **CHK_5**: Random sampling for manual review (split into 5א and 5ב for decision methods 1 and 2)
6. **CHK_6**: Price validation (TASE price comparison via Selenium + price>100 checks)
7. **CHK_7**: Problematic securities (דלי סחירות, רשימת שימור, מושעים)

**Key Design Points:**
- Each check has a dedicated logger (`logger_chk1`, `logger_chk3`, etc.) writing to separate log files
- Logs are organized in timestamped run directories: `log/YYYYMMDD_HHMMSS_UUID/`
- Hebrew column headers expected in input files (CSV/XLSX)
- Dates stored as DDMMYYYY integers, times as HHMMSS integers (no leading zeros)
- CSV files must be UTF-8 encoded (BOM handled automatically)
- Unique transaction ID = security_no + date

### Batch Processing: `batch_special_transactions.py`

Orchestrates validation across all 10 fund managers:
1. Fetches Mutual Funds List from Apify (actor K9WppTziYC3n2vxTu)
2. For each manager, fetches special transactions report from Apify (actor 5lhI6O39Qbgv9O0gs)
3. Runs `mizrahi_special_transactions.py` validation
4. Generates timestamped output directory with manager-specific subdirectories

### Email Sender: `send_batch_email.py`

Sends batch processing results via Gmail SMTP with attachments.

### Testing Script: `test_special_transactions_all_managers.py`

Tests Apify actor for coordinated/off-exchange transactions (build nQh62mdhpUTM5l65l) across all managers.

### Helper Scripts

**Shell Scripts** (`run_batch_all.sh`, `run_batch_test.sh`):
- Full batch processing with optional email sending
- Single manager test runs
- Assumes virtual environment at `/root/mizrahi-venv/` (Linux paths hardcoded)
- Use Python scripts directly for Windows or non-standard setups

## Common Commands

### Single Manager Validation

```bash
python mizrahi_special_transactions.py \
  --mutual-funds-list "Mutual Funds List.xlsx" \
  --input-report "manager_report.csv" \
  --output-xlsx "output.xlsx" \
  --email-json "email.json" \
  --manager-name "סיגמא" \
  --price-threshold 5.0 \
  --seed 123
```

**Key arguments:**
- `--skip-tase-prices`: Skip Selenium price checks (faster processing)
- `--spec-file`: Path to specification Excel file for פירוט בדיקות sheet (default: spec-file.xlsx)
- `--price-threshold`: Variance threshold % for TASE price comparison (default: 5.0)
- `--seed`: RNG seed for reproducible sampling
- `--report-month`: YYYY-MM format (auto-inferred from ת.דוח column if omitted)

**Note**: The specification file (spec-file.xlsx) is included in the repository and contains validation rules details.

### Batch Processing All Managers

```bash
# Process all 10 managers
python batch_special_transactions.py \
  --apify-token "$APIFY_TOKEN" \
  --output-dir ./batch_output \
  --skip-tase-prices

# Process specific managers only
python batch_special_transactions.py \
  --apify-token "$APIFY_TOKEN" \
  --managers "מגדל,איילון,סיגמא" \
  --output-dir ./batch_output

# Test mode: Generate HTML email previews instead of sending
python batch_special_transactions.py \
  --apify-token "$APIFY_TOKEN" \
  --managers "מגדל" \
  --output-dir ./batch_output \
  --test-mode \
  --skip-tase-prices
```

**Test Mode Notes:**
- Automatically enabled if `RESEND_API_KEY` is not set
- Can be explicitly enabled with `--test-mode` flag
- Creates HTML email preview files in each manager's output directory
- Useful for reviewing email content before sending

### Send Emails for Single Manager

```bash
python send_special_txn_email.py \
  --manager-name "מגדל" \
  --xlsx-path "./output/מגדל_special_transactions_report.xlsx" \
  --email-json-path "./output/מגדל_email.json" \
  --report-month "דצמבר 2025" \
  --recipient "manager@example.com"
```

**Sends 2 emails per manager:**
1. Report email with Excel attachment
2. Transaction samples email (only if decision_1 or decision_2 samples exist)

**Test Mode (Preview Emails as HTML):**
```bash
# Option 1: Explicit test mode flag
python send_special_txn_email.py \
  --manager-name "מגדל" \
  --xlsx-path "./output/מגדל_special_transactions_report.xlsx" \
  --email-json-path "./output/מגדל_email.json" \
  --report-month "דצמבר 2025" \
  --recipient "manager@example.com" \
  --test-mode

# Option 2: Automatic test mode (no RESEND_API_KEY set)
# Simply run without RESEND_API_KEY environment variable
```

**Test mode behavior:**
- If `--test-mode` flag is set OR `RESEND_API_KEY` is not configured
- Creates HTML files instead of sending emails:
  - `{manager}_email_1_report.html` - Report email preview
  - `{manager}_email_2_samples.html` - Samples email preview (if samples exist)
- HTML files saved to output directory (same location as xlsx file)
- Open HTML files in browser to preview email formatting

### Shell Scripts

```bash
# Full batch run for all managers (emails sent automatically)
./run_batch_all.sh YOUR_APIFY_TOKEN

# Test run for single manager
./run_batch_test.sh YOUR_APIFY_TOKEN [MANAGER_NAME]
```

**Note**: Emails are sent automatically during batch processing if RESEND_API_KEY is set.

## Testing

Test suites organized under `tests/` directory. See `tests/README.md` for comprehensive documentation.

### Quick Test Commands

**Run דח"צ Validation Unit Test:**
```bash
cd tests/test_dachatz_validation
python run_test.py
```
Expected: 6 CHK_4 exceptions (1 Rule A, 5 Rule B)

**Single Manager Test:**
```bash
./run_batch_test.sh YOUR_APIFY_TOKEN [MANAGER_NAME]
```

**Test with Specific Manager:**
```bash
python batch_special_transactions.py \
  --apify-token "$APIFY_TOKEN" \
  --managers "סיגמא" \
  --output-dir ./test_output \
  --skip-tase-prices
```

### Test Directory Structure
Test runs create timestamped subdirectories (YYYYMMDD_HHMMSS) containing:
- Manager-specific output folders
- Validation reports (XLSX)
- Email configurations (JSON)
- Log files
- Summary data

### Available Test Suites
- `test_dachatz_validation/`: Unit test for CHK_4 דח"צ validation rules (Rule A: דחצ1=0, Rule B: any דחצ=2)
- `test_header_fix/`: Integration tests with real manager data
- `test_separate_sheets/`: Sheet formatting tests
- `test_separate_sheets_sigma/`: Sigma-specific tests

## Fund Managers

10 supported managers (name → code):
- מגדל: 10040
- איילון: 10054
- קסם: 10047
- סיגמא: 10048
- פורסט: 10082
- הראל: 10031
- אנליסט: 10019
- מיטב: 10083
- איביאי: 10068
- אלטשולר-שחם: 10017

## Environment Setup

### Required Dependencies
- Python 3.10+
- openpyxl
- requests
- resend (for email sending)

### Optional Dependencies
- selenium (for CHK_6 TASE price checks)
- webdriver-manager

If Selenium unavailable, CHK_6 price comparisons are skipped with warnings.

**Install all dependencies:**
```bash
pip install -r requirements.txt
```

### Environment Variables
Copy `.env.example` to `.env` and configure:

```bash
# Required for CHK_7 (Problematic Securities Check)
TASE_API_KEY=your_tase_api_key_here

# Required for batch processing
APIFY_TOKEN=your_apify_token_here

# Required for email sending
RESEND_API_KEY=re_xxxxxxxxxxxxxxxxxxxxxxxxxxxx
```

**API Key Sources:**
- TASE API Key: https://datawise.tase.co.il
- Apify Token: https://console.apify.com/account/integrations
- Resend API Key: https://resend.com/api-keys (requires verified sender domain)

## Email System

The system sends **two emails per manager report** using Resend API:

### Email 1: Report Email
**Rich HTML summary with Excel attachment**
- Contains validation check summary (בדיקות #1-#7)
- Lists all performed checks with sub-check details
- Includes complete Excel report as attachment
- Professional HTML styling with Hebrew RTL support
- Sent to fund manager

**Content includes:**
- Manager name and report month
- Report filename and creation date
- Complete list of all validation checks performed
- Excel attachment with full validation results

### Email 2: Transaction Samples Email (Conditional)
**Hebrew text email with transaction tables**
- Only sent if decision_1 or decision_2 samples exist
- Contains HTML tables with sampled transactions for manual review
- Requests manual verification from fund manager
- Follows spec-file.xlsx format requirements

**Format (from spec-file.xlsx section 5.1):**
```
שלום רב,

בהתאם לבדיקת עסקאות מיוחדות לחודש {report_month},
עלו לנו עסקאות שעליהן נבקש לקבל פרטים נוספים:

[Table with: מספר קרן, שם קרן, שם נייר, מספר נייר, כמות, מחיר, תאריך, סוג]

נשמח לקבל את התייחסותכם בהקדם האפשרי.

בברכה,
צוות מזרחי טפחות החברה לנאמנות
```

**Email Service**: Resend API (requires `RESEND_API_KEY` environment variable)

**Invocation**: Per-manager basis (2 emails sent automatically after each manager completes processing during batch runs)

**Sender Configuration**: Emails sent from `noreply@82labs.io` (must be verified domain in Resend)

### Test Mode - HTML Email Previews

**Automatic Test Mode:**
- If `RESEND_API_KEY` environment variable is not set, the system automatically enters test mode
- Instead of sending emails, HTML preview files are created in the output directory

**Manual Test Mode:**
- Add `--test-mode` flag to either script to force test mode even with API key configured
- Useful for previewing email content before sending

**Generated Files:**
- `{manager_name}_email_1_report.html` - Report email with full HTML styling
- `{manager_name}_email_2_samples.html` - Samples email (only if samples exist)
- Files saved in same directory as Excel report
- Open in browser to preview exactly how emails will appear

**Use Cases:**
- Development and testing without sending real emails
- Previewing email formatting and content
- Reviewing Hebrew text rendering
- Sharing email previews with stakeholders
- Running batch processing without email service configured

## Output Files

### Batch Processing Structure
```
batch_output/
└── YYYYMMDD_HHMMSS/
    ├── Mutual_Funds_List.xlsx
    ├── batch_summary.txt
    ├── batch_summary.json
    └── {manager_name}/
        ├── {manager}_special_transactions.csv
        ├── {manager}_special_transactions_report.xlsx
        ├── {manager}_email.json
        ├── {manager}_email_1_report.html         # Test mode only
        └── {manager}_email_2_samples.html        # Test mode only (if samples exist)
```

**Note**: HTML email preview files are created automatically when:
- `RESEND_API_KEY` is not set (automatic test mode), OR
- `--test-mode` flag is used (manual test mode)

### Validation Output XLSX Sheets
1. **סיכום** (Summary)
2. **סטטוס בדיקות** (Check Status) - includes 4 separate rows for CHK_4 sub-checks
3. **Exception Sheets** (created only if exceptions exist):
   - בדיקה #1 - עסקאות בין קרנות
   - בדיקה #3 - תאריך
   - בדיקה #4ד - אופן החלטה (Missing data)
   - בדיקה #4א - אופן החלטה (Type-decision compatibility)
   - בדיקה #4ב - אופן החלטה (דח"צ opposition)
   - בדיקה #4ג - אופן החלטה (דח"צ1 no decision)
   - בדיקה #6 - חריגות מחיר
   - בדיקה #7 - ניירות בעייתיים
4. **דגימות אקראיות** (Random Samples)
5. **פירוט בדיקות** (Specification details - if spec-file provided)
6. **קרנות מחוץ להיקף** (Out-of-scope funds)

### Email JSON Structure
Contains transactions grouped by decision method (אופן החלטה). Returns empty if no decision method 1 or 2 transactions exist.

## Logging

Each run creates a unique log directory: `log/YYYYMMDD_HHMMSS_UUID/`

**Log files:**
- `main.log`: General processing log
- `chk1_inter_fund.log`: Inter-fund transaction checks
- `chk3_date.log`: Date validation
- `chk4_decision.log`: Decision method rules (includes all 4 sub-checks: 4A, 4B, 4C, 4D)
- `chk6_tase_price.log`: TASE price comparisons
- `chk6_price_limit.log`: Price > 100 checks
- `chk6_failed_urls.log`: Failed URL fetches requiring manual verification
- `chk7_problematic.log`: Problematic securities checks

## Hebrew Field Names

**Mutual Funds List:**
- מספר בורסה (Fund ID)
- שם נאמן (Trustee name - filter for 'מזרחי טפחות חברה לנאמנות בע"מ')

**Manager Report:**
- מספר קרן, שם קרן (Fund number/name)
- מספר נייר, שם נייר (Security number/name)
- כמות (Quantity)
- מחיר (Price)
- תאריך (Date - DDMMYYYY)
- שעה (Time - HHMMSS)
- סוג (Type)
- אופן החלטה (Decision method)
- ת.דוח or ת. דוח (Report date)
- דחצ1-4 (Decision criteria fields for CHK_4)

## Apify Integration

**Actors:**
- `K9WppTziYC3n2vxTu`: Mutual Funds List fetcher
- `5lhI6O39Qbgv9O0gs`: Fund reports fetcher (also used for special transactions)
- Build `nQh62mdhpUTM5l65l`: Special build for עסקה מתואמת/מחוץ לבורסה

**API Usage:**
- Runs via Apify API with timeout=300s
- Polls run status every 3 seconds
- Downloads results as base64-decoded files

## Performance Notes

- Batch processing 10 managers with TASE checks: ~20-30 minutes
- With `--skip-tase-prices`: ~5-10 minutes
- Process specific managers only with `--managers` flag to speed up testing

## Important Implementation Details

### UTF-8 Handling
Windows output compatibility ensured via:
```python
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
```
This is critical for Hebrew text rendering in console output.

### CHK_7 TASE API Integration
CHK_7 (Problematic Securities) requires TASE_API_KEY environment variable. The system fetches three lists from TASE DataWise API:
- דלי סחירות (low liquidity securities)
- רשימת שימור (watchlist securities)
- מושעים (suspended trading securities)

Without the API key, CHK_7 will be skipped with warnings.

### Date and Time Format Conventions
- Dates: DDMMYYYY integers (no leading zeros), e.g., 1012024 = January 1, 2024
- Times: HHMMSS integers (no leading zeros), e.g., 93045 = 09:30:45
- Report month: YYYY-MM string format
- Unique transaction ID: security_no + date (used for duplicate detection)
