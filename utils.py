"""
utils.py
--------
Shared helper functions for the ASL Inventory Parameter Optimization Tool.

Everything here is intentionally small, dependency-light and side-effect free so
that it is easy to reuse and easy to unit-test.
"""

import math


def to_number(value):
    """
    Safely convert an arbitrary cell value to a float.

    Returns ``None`` when the value is blank, ``NaN`` or cannot be parsed as a
    number.  This keeps the rule code free of repetitive try/except blocks and
    avoids accidental division-by-zero / type errors.
    """
    if value is None:
        return None

    # pandas uses float('nan') for blank numeric cells.
    if isinstance(value, float) and math.isnan(value):
        return None

    if isinstance(value, (int, float)):
        return float(value)

    # Strings such as "1,234.0", " 12 " or "" need cleaning.
    text = str(value).strip()
    if text == "" or text.lower() in ("nan", "none", "n/a", "na"):
        return None

    text = text.replace(",", "")
    try:
        return float(text)
    except (ValueError, TypeError):
        return None


def clean_text(value):
    """Return a trimmed string for a cell value ('' when blank/NaN)."""
    if value is None:
        return ""
    if isinstance(value, float) and math.isnan(value):
        return ""
    return str(value).strip()


def norm_upper(value):
    """Trimmed, upper-cased text used for case-insensitive comparisons."""
    return clean_text(value).upper()


def ceil_positive(value):
    """Ceiling that is safe for ``None`` and negative numbers."""
    if value is None:
        return 0
    return int(math.ceil(value))
