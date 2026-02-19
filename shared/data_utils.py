#!/usr/bin/env python3
"""Shared data conversion utilities."""

import datetime as dt
from typing import Optional, Any

try:
    import pandas as pd
    HAS_PANDAS = True
except ImportError:
    HAS_PANDAS = False


def to_str(value: Any) -> Optional[str]:
    """
    Convert value to string, handling None and pandas NaN.

    Args:
        value: Any value to convert

    Returns:
        String representation or None if empty/NaN
    """
    if value is None:
        return None
    if HAS_PANDAS and isinstance(value, float) and pd.isna(value):
        return None
    s = str(value).strip()
    if s.lower() == "nan" or not s:
        return None
    return s


def to_int(value: Any) -> Optional[int]:
    """
    Convert value to int, handling None and non-numeric values.

    Args:
        value: Any value to convert

    Returns:
        Integer value or None if not convertible
    """
    if value is None or value == "":
        return None
    if HAS_PANDAS and isinstance(value, float) and pd.isna(value):
        return None
    try:
        return int(float(str(value).strip()))
    except (ValueError, TypeError):
        return None


def to_float(value: Any) -> Optional[float]:
    """
    Convert value to float, handling None and non-numeric values.

    Args:
        value: Any value to convert

    Returns:
        Float value or None if not convertible
    """
    if value is None or value == "":
        return None
    if HAS_PANDAS and isinstance(value, float) and pd.isna(value):
        return None
    try:
        return float(str(value).strip())
    except (ValueError, TypeError):
        return None


def parse_date_ddmmyyyy(value: Any) -> Optional[dt.date]:
    """
    Parse date from DD/MM/YYYY or DDMMYYYY format.

    Args:
        value: Date value (string, datetime, or date)

    Returns:
        date object or None if not parseable
    """
    if value is None:
        return None

    # Handle datetime objects
    if isinstance(value, dt.datetime):
        return value.date()
    if isinstance(value, dt.date):
        return value

    s = str(value).strip()
    if not s:
        return None

    # Try DD/MM/YYYY
    if '/' in s:
        parts = s.split('/')
        if len(parts) == 3:
            try:
                return dt.date(int(parts[2]), int(parts[1]), int(parts[0]))
            except (ValueError, IndexError):
                pass

    # Try DDMMYYYY (no separators)
    if len(s) == 8 and s.isdigit():
        try:
            return dt.date(int(s[4:8]), int(s[2:4]), int(s[0:2]))
        except ValueError:
            pass

    # Try YYYY-MM-DD (ISO format)
    if '-' in s and len(s) == 10:
        parts = s.split('-')
        if len(parts) == 3:
            try:
                return dt.date(int(parts[0]), int(parts[1]), int(parts[2]))
            except (ValueError, IndexError):
                pass

    return None


def fix_shifted_encoding(content: bytes) -> bytes:
    """
    Fix files with shifted Hebrew encoding (0x10 offset).

    Some Maya exports have a peculiar encoding issue where Hebrew characters
    are shifted by 0x10 from their correct cp1255 positions.

    Args:
        content: Raw file content as bytes

    Returns:
        Fixed content as UTF-8 encoded bytes
    """
    if not content or content[0] != 0xff:
        return content
    content = content[1:]
    fixed = bytearray()
    for b in content:
        if 0xce <= b <= 0xea:
            fixed.append(b + 0x10)
        else:
            fixed.append(b)
    return bytes(fixed).decode('cp1255').encode('utf-8-sig')


def normalize_spaces(text: str) -> str:
    """
    Normalize whitespace in text.

    Args:
        text: Input text

    Returns:
        Text with normalized whitespace
    """
    if not text:
        return ""
    return " ".join(text.split())


def clean_excel_string(value: Any) -> str:
    """
    Clean a string for safe Excel output.

    Removes illegal XML characters that would cause openpyxl to fail.

    Args:
        value: Any value to clean

    Returns:
        Cleaned string safe for Excel
    """
    if value is None:
        return ""
    s = str(value)
    # Remove illegal XML characters (control chars except tab, newline, carriage return)
    return ''.join(c for c in s if c >= ' ' or c in '\t\n\r')
