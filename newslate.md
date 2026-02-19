# Specification: Disclosure Report K.303 Variant (Migdal)

This document defines the **new project variant** that validates the K.303 disclosure report using a similar Excel output template to the existing `mizrahi_special_transactions.py` tool, but with **different inputs, checks, and report logic**. It is intended to be handed to a code-generation agent to build the new baseline implementation.

---

## 1. Baseline Reference (Current Project Behavior)

The current project (`mizrahi_special_transactions.py`) processes **special transactions** and produces an Excel report with:

- **Inputs**: Mutual funds list (XLSX), manager special-transactions report (CSV/XLSX), optional spec sheet (XLSX).
- **Checks**: #1–#7 (duplicates/inter-fund transactions, report month dates, decision-method rules, price/type consistency, sampling, TASE price variance, price limit, problematic securities).
- **Outputs**: XLSX workbook with `סיכום`, `פירוט בדיקות`, `סטטוס בדיקות`, and per-check exception sheets, plus optional email JSON.  

This new variant **does not reuse** those checks. The new implementation should **keep the report format and Excel styling template** but **replace the inputs, checks, and sheet contents** as described below.

---

## 2. New Inputs (K.303 Disclosure Report)

### 2.1 Current Month Disclosure Report  
File: `disclosure_migdal_previous_month.xlsx - Sheet1.csv` (Excel Sheet1 exported to CSV for reference).  
Expected format (columns in order):  
1. **מספר קרן**  
2. **שם קרן**  
3. **רמה 1**  
4. **רמה 2**  
5. **רמה 3**  
6. **רמה 4**  
7. **%מקרן**  
8. **נתונים נוספים**  
9. **תאריך דוח**  
10. **מס.רשומה**  
11. **סהכ רשומות**  
12. **מס.מנהל ברשם** (note: in the CSV header it includes a stray `\r`, so the parser should trim whitespace).  

**Key parsing expectations:**
- All four “רמה” fields are numeric codes but may be empty; use the **most granular non-empty code** as the “effective code”.  
- `%מקרן` is numeric and is the value used in reasonableness checks.  
- `תאריך דוח` is in **DDMMYYYY** numeric form (e.g., `30112025`) without leading zeros; convert to a date and derive the report month.  

### 2.2 Previous Month Disclosure Report  
Same format as above. Required for check **2א** (month-over-month reasonableness).

### 2.3 Mutual Funds List (Static Reference)
File: `Mutual_Funds_List.xlsx - Worksheet.csv`.  
Key columns used:
- **מספר בורסה** (fund identifier; join key with `מספר קרן` from the disclosure report).  
- **שם נאמן** (filter to Mizrahi trustee funds).  
- **שם מנהל** (used for summary metadata).  
- **פרופיל החשיפה** (column P in the CSV header; used for exposure sanity checks).  

### 2.4 Checking Articles / Spec Sheet  
File: `בדיקת דוח גילוי נאות ק.303.xlsx - גיליון1.csv`.  
This file provides the testing checklist and should be copied into the output workbook as the `פירוט בדיקות` sheet (styling preserved if possible).

---

## 3. Data Scoping Rules

1. **In-scope funds** are those where `שם נאמן == "מזרחי טפחות חברה לנאמנות בע\"מ"` in the Mutual Funds list.  
2. Join the disclosure report rows on **`מספר קרן` ↔ `מספר בורסה`** to determine whether each row is in scope.  
3. Out-of-scope rows are still counted for summary statistics but are excluded from checks unless explicitly stated otherwise.  

---

## 4. New Checks (Based on Draft Checklist)

### 4.1 Check 1א — Fund Completeness (Mutual Funds ↔ Disclosure Report)
**Purpose:** Validate that the current-month disclosure report includes **all funds in scope** and does not include unexpected funds.

- Build `funds_in_mutual_list` (filtered to Mizrahi trustee).  
- Build `funds_in_report` (unique `מספר קרן` in disclosure report).  
**Exceptions:**  
- Funds in mutual list but missing from report.  
- Funds in report but missing from mutual list.  
**Output sheet:** `בדיקה #1א - שלמות קרנות`  
Include two categories (missing-from-report, extra-in-report) with fund number + name.

### 4.2 Check 1ב — Report Month Validity
**Purpose:** The `תאריך דוח` for each row must match the expected report month.

- Determine `report_month` from CLI input or infer from the most common `תאריך דוח` in the report.  
- Flag any row where the month/year does not match.  
**Output sheet:** `בדיקה #1ב - תקינות תאריכים`

### 4.3 Check 2א — Reasonableness vs Previous Month
**Purpose:** Compare disclosure codes between the current and previous month reports.

**Rules (per checklist):**
- For each fund + effective code, compare `%מקרן`.  
- Flag **absolute change > 10%** (percentage points).  
- Flag any effective code **added** or **removed** between months.

**Output sheet:** `בדיקה #2א - סבירות מול דוח קודם`

### 4.4 Check 2ב — Reasonableness vs Fund Exposure Profile
**Purpose:** Cross-check disclosure exposure codes against the fund’s exposure profile (`פרופיל החשיפה`).

**Requirements from checklist:**
1. **Exposure to equities (code 01)**  
   - If exposure profile indicates **0 exposure to equities**, then any `01` disclosure code must have `%מקרן == 0`.  
2. **Exposure to FX (code 06)**  
   - If exposure profile indicates **0 exposure to FX**, then any `06` disclosure code must have `%מקרן == 0`.  
3. **Non-investment-grade bonds (code 070201)**  
   - Marked **“לא לפתח”** in the checklist → **do not implement**.  
4. **Fund name ↔ code mapping**  
   - Marked **“בשלב הבא”** → **do not implement**.  

**Implementation note:**  
The mutual funds profile format is encoded (e.g., `0A`, `3B`). The new implementation must include a **mapping table** (“מקרא לפרופיל חשיפה”) to determine whether equities/FX exposure is zero.  

**Output sheet:** `בדיקה #2ב - סבירות מול מאפייני הקרן`

### 4.5 Check 3 — Within-Month Combinations and Cross-Checks
**Purpose:** Validate the presence of required disclosure codes **within the same month**.

Use the **effective code** per row (deepest non-empty “רמה” code).  
When the checklist references a broader category (e.g., `03` or `07`), treat it as a **prefix match** (any effective code that starts with the given code).

#### Check 3א — FX Exposure
If any of the following codes exist:  
`0102` (equities abroad), `0302` (bonds abroad), `0502` (cash in FX),  
then code `06` must also exist.  
Also, if code `06` exists, at least one of the three FX-related codes must exist.  

**Output sheet:** `בדיקה #3א - חשיפה למט"ח`

#### Check 3ב — Bond Exposure
If code `03` (bonds) exists, then both:
- code `07` (ratings) **and**  
- code `08` (duration / מח״מ)  
must also exist.  

Also, if either `07` or `08` exists, code `03` must exist.  

**Output sheet:** `בדיקה #3ב - חשיפה לאג"ח`

#### Check 3ג — Government Bonds (Shekel)
If code `03010101` exists, then code `080201` must also exist.  
If `080201` exists, then `03010101` must exist.  

**Output sheet:** `בדיקה #3ג - אג"ח ממשלתי שקלי`

#### Check 3ד — Government Bonds (Linked)
If code `03010102` exists, then code `080202` must also exist.  
If `080202` exists, then `03010102` must exist.  

**Output sheet:** `בדיקה #3ד - אג"ח ממשלתי צמוד`

#### Check 3ה — Government Bonds (Linked FX)
If code `03010103` exists, then code `080203` must also exist.  
If `080203` exists, then `03010103` must exist.  

**Output sheet:** `בדיקה #3ה - אג"ח ממשלתי צמוד מט"ח`

#### Check 3ו — Corporate Bonds (Shekel)
If code `03010202` or `03010203` exists, then code `080204` must also exist.  
If `080204` exists, then at least one of `03010202`/`03010203` must exist.  

**Output sheet:** `בדיקה #3ו - אג"ח קונצרני שקלי`

#### Check 3ז — Corporate Bonds (Linked)
If code `03010201` exists, then code `080205` must also exist.  
If `080205` exists, then `03010201` must exist.  

**Output sheet:** `בדיקה #3ז - אג"ח קונצרני צמוד`

#### Check 3ח — Corporate Bonds (Linked FX)
If code `03010204` exists, then code `080206` must also exist.  
If `080206` exists, then `03010204` must exist.  

**Output sheet:** `בדיקה #3ח - אג"ח קונצרני צמוד מט"ח`

---

## 5. Output Report (Excel)

The output should keep the **same structure and styling** as the baseline Excel report:

### Required Sheets
1. **סיכום**  
   - Include: manager name, trustee name, report month, number of funds in report, number of in-scope funds, number of out-of-scope funds, total rows.  
2. **פירוט בדיקות**  
   - Copy from the K.303 checklist file (`בדיקת דוח גילוי נאות ק.303.xlsx` → `גיליון1`).  
3. **סטטוס בדיקות**  
   - One row per check (#1א, #1ב, #2א, #2ב, #3א–#3ח), with pass/fail status and exception count.  
4. **Per-check exception sheets** (created only when exceptions exist):  
   - `בדיקה #1א - שלמות קרנות`  
   - `בדיקה #1ב - תקינות תאריכים`  
   - `בדיקה #2א - סבירות מול דוח קודם`  
   - `בדיקה #2ב - סבירות מול מאפייני הקרן`  
   - `בדיקה #3א - חשיפה למט"ח`  
   - `בדיקה #3ב - חשיפה לאג"ח`  
   - `בדיקה #3ג - אג"ח ממשלתי שקלי`  
   - `בדיקה #3ד - אג"ח ממשלתי צמוד`  
   - `בדיקה #3ה - אג"ח ממשלתי צמוד מט"ח`  
   - `בדיקה #3ו - אג"ח קונצרני שקלי`  
   - `בדיקה #3ז - אג"ח קונצרני צמוד`  
   - `בדיקה #3ח - אג"ח קונצרני צמוד מט"ח`

### Exception Sheet Columns (suggested)
For consistency with the baseline:
- **בדיקה**, **סיבה**, **מספר קרן**, **שם קרן**, **קוד חשיפה**, **%מקרן**, **תאריך דוח**, **שורה בקובץ**, plus validation columns `טופל?`, `שם הבודק`.  
- Check 1א (completeness) should include **category** (missing vs extra) in the reason field.  
- Check 2א should include previous-month `%מקרן`, delta, and flag type (added/removed vs delta > 10%).  

---

## 6. CLI / Workflow Changes

Replace the current CLI parameters with:
- `--mutual-funds-list` (same as baseline)  
- `--current-report` (current-month disclosure report)  
- `--previous-report` (previous-month disclosure report)  
- `--output-xlsx`  
- `--spec-file` (K.303 checklist, optional but recommended)  
- Optional: `--report-month` (YYYY-MM) override

**Remove**:
- Email JSON output.  
- Selenium/TASE price checks.  
- Sampling logic.  
- Decision method validations (not relevant).  

---

## 7. Reusable Components from Baseline

The new implementation should reuse:
- Excel workbook creation / styling utilities (RTL formatting, auto-fit).  
- Summary and status-sheet scaffolding (adapt field names and checks).  
- Spec sheet copy helper (for `פירוט בדיקות`).  
- Logging infrastructure (main log + per-check logs).

---

## 8. Items Explicitly Out of Scope

From the draft checklist:
- **Check 2ב**: non-investment-grade bonds (`070201`) → **do not implement**.  
- **Check 2ב**: fund-name to code mapping → **future phase only**.  

These should be documented in code comments and excluded from the current release.
