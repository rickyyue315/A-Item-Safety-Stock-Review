"""Buyer-related utilities."""
import re


def extract_buyer_number(buyer_str):
    """Extract numeric part from buyer string for sorting."""
    m = re.match(r"(\d+)", str(buyer_str))
    return int(m.group(1)) if m else 0


def get_col_idx(raw_headers):
    """Build column name to index mapping for O(1) lookups."""
    return {name: i for i, name in enumerate(raw_headers)}


def normalize_row(row, width):
    """Pad a worksheet row to ``width`` values.

    openpyxl's read-only mode yields short tuples when a row has fewer cells
    than the sheet's widest row (trailing empty cells are not included), so
    indexing by header position can raise IndexError. Padding with None keeps
    positional access safe.
    """
    if row is None:
        return (None,) * width
    row = tuple(row)
    if len(row) >= width:
        return row
    return row + (None,) * (width - len(row))
