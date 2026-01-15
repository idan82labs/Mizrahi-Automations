# Tests Directory

This directory contains all test suites for the Mizrahi Special Transactions validation system.

## Test Suites

### 1. test_dachatz_validation/
Unit test for the new דחצ (decision criteria) validation rules in CHK_4.

**Purpose**: Validates Rules A and B for דחצ fields (דחצ1-4)
- Rule A: Flag when דחצ1 = 0
- Rule B: Flag when any דחצ = 2 (takes priority over Rule A)

**Run**: `python tests/test_dachatz_validation/run_test.py`

**Files**:
- `run_test.py` - Test runner
- `test_data.xlsx` - 10 test transactions covering all scenarios
- `test_funds_list.xlsx` - Test mutual funds list
- `test_email.json` - Email configuration
- `README.md` - Detailed test documentation

### 2. test_header_fix/
Integration tests for header parsing and validation with real manager data.

### 3. test_separate_sheets/
Tests for separate sheet output formatting.

### 4. test_separate_sheets_sigma/
Tests specific to Sigma fund manager data.

## Test Data Organization

All test runs are stored in timestamped subfolders:
```
tests/
├── test_dachatz_validation/     # Unit test for דחצ validation
├── test_header_fix/              # Integration tests
│   └── YYYYMMDD_HHMMSS/         # Timestamped test runs
├── test_separate_sheets/         # Sheet formatting tests
│   └── YYYYMMDD_HHMMSS/
└── test_separate_sheets_sigma/   # Sigma-specific tests
    └── YYYYMMDD_HHMMSS/
```

Each timestamped test run contains:
- Manager-specific output folders
- Validation reports (XLSX)
- Email configurations (JSON)
- Log files
- Summary data

## Running Tests

### דחצ Validation Test
```bash
cd tests/test_dachatz_validation
python run_test.py
```

Expected output: 6 CHK_4 exceptions (1 Rule A, 5 Rule B)
