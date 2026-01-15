# Batch Special Transactions Processing - Quick Start Guide

This guide explains how to run the batch processor for all fund managers and send results to elay.g@82labs.io.

## Overview

The batch processing system consists of three main scripts:

1. **batch_special_transactions.py** - Main batch processor that:
   - Fetches Mutual Funds List from Apify
   - For each manager, fetches their special transactions report
   - Runs mizrahi_special_transactions.py validation
   - Generates output XLSX and email JSON files

2. **send_batch_email.py** - Email sender that:
   - Collects all results from batch processing
   - Sends consolidated email with attachments to specified recipient

3. **mizrahi_special_transactions.py** - Core validation script (already exists)

## Prerequisites

### 1. Python Packages

```bash
source /root/mizrahi-venv/bin/activate  # If using virtual environment
pip install requests openpyxl
```

### 2. Apify API Token

You need an Apify API token to fetch data from Apify actors.

### 3. Gmail App Password (for email sending)

To send emails via Gmail:
1. Enable 2-factor authentication on your Google account
2. Go to https://myaccount.google.com/apppasswords
3. Create an App Password for "Mail"
4. Save this password (you'll need it for the email sender)

## Quick Start - Run Everything

### Step 1: Set Your Apify Token

```bash
export APIFY_TOKEN="your_apify_token_here"
```

### Step 2: Run Batch Processor

Process all 10 managers (recommended):

```bash
python /root/batch_special_transactions.py \
  --apify-token "$APIFY_TOKEN" \
  --output-dir ./batch_output \
  --skip-tase-prices
```

Process specific managers only:

```bash
python /root/batch_special_transactions.py \
  --apify-token "$APIFY_TOKEN" \
  --managers "מגדל,איילון,סיגמא" \
  --output-dir ./batch_output
```

With all options:

```bash
python /root/batch_special_transactions.py \
  --apify-token "$APIFY_TOKEN" \
  --output-dir ./batch_output \
  --price-threshold 5.0 \
  --seed 123 \
  --email elay.g@82labs.io \
  --skip-tase-prices
```

### Step 3: Send Results Email

After batch processing completes, send the results:

```bash
python /root/send_batch_email.py \
  --batch-dir ./batch_output/20260114_120000 \
  --gmail-user your.email@gmail.com \
  --gmail-app-password "your_app_password" \
  --recipient elay.g@82labs.io
```

Replace `20260114_120000` with the actual timestamp directory created by the batch processor.

## Command-Line Arguments

### batch_special_transactions.py

| Argument | Required | Description |
|----------|----------|-------------|
| `--apify-token` | Yes | Apify API token for fetching data |
| `--managers` | No | Comma-separated list of managers (default: all 10) |
| `--output-dir` | No | Base output directory (default: ./batch_output) |
| `--skip-tase-prices` | No | Skip TASE price checks for faster processing |
| `--price-threshold` | No | Price variance threshold % (default: 5.0) |
| `--spec-file` | No | Path to specification table Excel file |
| `--seed` | No | RNG seed for reproducible sampling |
| `--email` | No | Email recipient (default: elay.g@82labs.io) |

### send_batch_email.py

| Argument | Required | Description |
|----------|----------|-------------|
| `--batch-dir` | Yes | Path to batch output directory with timestamp |
| `--gmail-user` | Yes | Gmail email address for sending |
| `--gmail-app-password` | Yes | Gmail App Password (not regular password!) |
| `--recipient` | No | Email recipient (default: elay.g@82labs.io) |

## Output Structure

After running the batch processor, you'll have:

```
batch_output/
└── 20260114_120000/              # Timestamp directory
    ├── Mutual_Funds_List.xlsx    # Fetched from Apify
    ├── batch_summary.txt          # Human-readable summary
    ├── batch_summary.json         # JSON summary
    ├── מגדל/
    │   ├── מגדל_special_transactions.csv
    │   ├── מגדל_special_transactions_report.xlsx
    │   └── מגדל_email.json
    ├── איילון/
    │   ├── איילון_special_transactions.csv
    │   ├── איילון_special_transactions_report.xlsx
    │   └── איילון_email.json
    └── ... (one directory per manager)
```

## Fund Managers Supported

The batch processor supports all 10 fund managers:

1. מגדל (Migdal) - 10040
2. איילון (Ayalon) - 10054
3. קסם (Kesem) - 10047
4. סיגמא (Sigma) - 10048
5. פורסט (Forest) - 10082
6. הראל (Harel) - 10031
7. אנליסט (Analyst) - 10019
8. מיטב (Meitav) - 10083
9. איביאי (IBI) - 10068
10. אלטשולר-שחם (Altshuler Shaham) - 10017

## Troubleshooting

### "No report available" for a manager

This means the Apify actor couldn't fetch the special transactions report for that manager. This could be due to:
- No reports available on Maya for that manager
- Different URL format needed for special transactions
- Network/timeout issues

### Email sending fails

Common issues:
1. **Not using App Password**: You must use a Gmail App Password, not your regular password
2. **2FA not enabled**: Enable 2-factor authentication first
3. **Wrong email**: Make sure you're using a Gmail address

### Batch processing is slow

The batch processor processes each manager sequentially. For 10 managers with TASE price checks, expect:
- With `--skip-tase-prices`: ~5-10 minutes total
- Without skipping: ~20-30 minutes total

To speed up:
- Use `--skip-tase-prices` flag
- Process only specific managers with `--managers`

## Test Mode - Single Manager

To test with just one manager first:

```bash
python /root/batch_special_transactions.py \
  --apify-token "$APIFY_TOKEN" \
  --managers "סיגמא" \
  --output-dir ./test_output \
  --skip-tase-prices
```

## Logs

The batch processor creates detailed logs in the output directory. Each manager's processing logs are stored in:
- Main script output: captured in batch_summary.json
- Individual script logs: in `log/` directories created by mizrahi_special_transactions.py

## Support

For issues or questions, contact: elay.g@82labs.io
