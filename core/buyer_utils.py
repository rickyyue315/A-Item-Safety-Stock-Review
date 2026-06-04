"""Buyer-related utilities."""
import re


def extract_buyer_number(buyer_str):
    """Extract numeric part from buyer string for sorting."""
    m = re.match(r"(\d+)", str(buyer_str))
    return int(m.group(1)) if m else 0


def get_col_idx(raw_headers):
    """Build column name to index mapping for O(1) lookups."""
    return {name: i for i, name in enumerate(raw_headers)}