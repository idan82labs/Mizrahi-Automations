#!/usr/bin/env python3
"""Shared constants for Mizrahi automations."""

# Fund Manager Codes - maps Hebrew names to TASE codes
FUND_MANAGER_CODES = {
    "מגדל": "10040",
    "קסם": "10047",
    "סיגמא": "10048",
    "הראל": "10031",
    "אנליסט": "10019",
    "מיטב": "10083",
    "איביאי": "10068",
    "אלטשולר-שחם": "10017",
}

# Apify Actor IDs
APIFY_ACTORS = {
    "FUNDS_LIST": "K9WppTziYC3n2vxTu",
    "FUND_REPORTS": "5lhI6O39Qbgv9O0gs",
    "K303_REPORTS": "iTpNz9ixbdQCmH43C",
    "INDX_HISTORICAL": "P9tr210PVi6W8RvtU",
}

# Default trustee name
MIZRAHI_TRUSTEE_NAME = 'מזרחי טפחות חברה לנאמנות בע"מ'

# Hebrew month names (number to name)
HEBREW_MONTHS = {
    1: "ינואר",
    2: "פברואר",
    3: "מרץ",
    4: "אפריל",
    5: "מאי",
    6: "יוני",
    7: "יולי",
    8: "אוגוסט",
    9: "ספטמבר",
    10: "אוקטובר",
    11: "נובמבר",
    12: "דצמבר",
}

# Reverse mapping (name to number)
HEBREW_MONTHS_REVERSE = {v: k for k, v in HEBREW_MONTHS.items()}

# Asset types flagged as unusual/requires attention
UNUSUAL_ASSET_TYPES = {16, 21, 22, 23, 24, 52, 53, 57, 58, 99, 101, 112, 201, 207, 209}

# Required asset type combinations
REQUIRED_COMBINATIONS = {
    111: [38, 42, 45, 47, 49, 56],
    212: [326, 327],
    213: [319],
    208: [307],
    210: [310],
}

# Threshold for combination 111 check
COMBINATION_111_THRESHOLD = 100000
