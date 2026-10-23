"""Overstock Shop Ratio logic: No. of Overstock Shop (P90D 0 Sales) ÷ No. of Shop Carry."""
from .buyer_utils import get_col_idx

OVERSTOCK_RATIO_MIN = 0.001
OVERSTOCK_RATIO_MAX = 0.20
OVERSTOCK_RATIO_HEADER = "Overstock Shop Ratio"
OVERSTOCK_COUNT_HEADER = "No. of Overstock Shop (P90D 0 Sales)"
SHOP_CARRY_HEADER = "No. of Shop Carry"


def compute_overstock_ratio(no_of_overstock_shop, no_of_shop_carry):
    """Return the overstock shop ratio, or None when No. of Shop Carry is blank or 0."""
    if not no_of_shop_carry or no_of_shop_carry <= 0:
        return None
    return (no_of_overstock_shop or 0) / no_of_shop_carry


def filter_overstock_ratio_data(ws):
    """Return (raw_headers, matched_rows, total_sku, no_carry_sku).

    matched_rows is a list of (raw_values, ratio) where the ratio is within
    OVERSTOCK_RATIO_MIN..OVERSTOCK_RATIO_MAX (both ends inclusive).
    total_sku counts data rows (footer rows excluded); no_carry_sku counts rows
    whose ratio cannot be calculated because No. of Shop Carry is blank or 0.
    """
    raw_headers = [cell.value for cell in ws[1]]
    col_idx = get_col_idx(raw_headers)
    overstock_idx = col_idx[OVERSTOCK_COUNT_HEADER]
    carry_idx = col_idx[SHOP_CARRY_HEADER]

    total_sku = 0
    no_carry_sku = 0
    matched_rows = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row or row[0] is None:
            continue
        if row[0] == "Total" or (isinstance(row[0], str) and row[0].startswith("Applied filters:")):
            continue

        total_sku += 1
        ratio = compute_overstock_ratio(row[overstock_idx], row[carry_idx])
        if ratio is None:
            no_carry_sku += 1
            continue
        if OVERSTOCK_RATIO_MIN <= ratio <= OVERSTOCK_RATIO_MAX:
            matched_rows.append((list(row), ratio))

    return raw_headers, matched_rows, total_sku, no_carry_sku
