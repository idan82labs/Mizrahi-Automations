# K.303 Report Downloader - Apify Puppeteer Scraper

This configuration downloads K.303 disclosure reports from the ISA Magna website.

## Setup Instructions

### 1. Go to Apify Console
- Navigate to: https://console.apify.com/actors/apify~puppeteer-scraper

### 2. Configure the Actor Input
- Copy the contents of `input.json` into the Input JSON editor
- Or use the form-based input editor with these settings

### 3. Input Parameters

| Parameter | Description | Example |
|-----------|-------------|---------|
| `managerName` | Hebrew name of the fund manager | מגדל קרנות נאמנות |
| `fromDate` | Start date (DD/MM/YYYY) | 01/01/2025 |
| `toDate` | End date (DD/MM/YYYY) | 01/02/2026 |

### 4. Run the Actor
- Click "Start" to run
- Downloaded files will be in the Key-Value Store
- Metadata will be in the Dataset

## URL Parameters

The Magna ISA search URL supports these query parameters:
- `form=ק303` - Filter to K.303 forms
- `q=<manager name>` - Search query (URL encoded Hebrew)

Example:
```
https://www.magna.isa.gov.il/?form=%D7%A7303&q=%D7%9E%D7%92%D7%93%D7%9C%20%D7%A7%D7%A8%D7%A0%D7%95%D7%AA%20%D7%A0%D7%90%D7%9E%D7%A0%D7%95%D7%AA
```

## Output

The actor outputs:
1. **Key-Value Store**: Downloaded XLSX files
   - `k303_current_<manager>.xlsx` - Current month report
   - `k303_previous_<manager>.xlsx` - Previous month report

2. **Dataset**: Metadata about downloads
   - Report dates
   - File URLs
   - Download status
