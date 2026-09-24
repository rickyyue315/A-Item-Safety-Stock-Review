"""ND Ratio checking logic."""
import datetime
from zoneinfo import ZoneInfo
import openpyxl
from .buyer_utils import get_col_idx

HK_TZ = ZoneInfo("Asia/Hong_Kong")

# Default number of shops used for ND ratio calculation.
DEFAULT_TOTAL_SHOPS = 92


def get_default_reference_date():
    """Return yesterday (T-1) in Hong Kong timezone as the default reference date.

    Because uploaded report data is usually T-1 (i.e. up to the previous day),
    using today's Hong Kong date minus 1 day as the default reference date aligns
    with the data's effective date.
    """
    return (datetime.datetime.now(HK_TZ) - datetime.timedelta(days=1)).date()


def compute_nd_checking(no_of_rf_shop, total_shops):
    """Compute ND ratio and checking category.

    Returns (nd_ratios, nd_checking) where nd_checking is one of:
      - "ND Ratio > 80%"
      - "ND Ratio 70%-80%"
      - "Check"        (data anomaly: RF shops exceed total shops)
      - None           (ratio within acceptable range, < 70%)
    """
    if no_of_rf_shop is None:
        nd_ratios = 1.0
    elif no_of_rf_shop > total_shops:
        # Data anomaly: more RF shops than total shops. Flag for review
        # instead of silently treating it as a healthy 0% ratio.
        return 0.0, "Check"
    else:
        nd_ratios = (total_shops - no_of_rf_shop) / total_shops if total_shops > 0 else 0

    if nd_ratios > 0.8:
        nd_checking = "ND Ratio > 80%"
    elif nd_ratios >= 0.7:
        nd_checking = "ND Ratio 70%-80%"
    else:
        nd_checking = None

    return nd_ratios, nd_checking


def _parse_launch_date(launch_date_val):
    """Parse Launch Date cell value to datetime.date or None if invalid."""
    if launch_date_val is None:
        return None
    if isinstance(launch_date_val, datetime.datetime):
        return launch_date_val.date()
    if isinstance(launch_date_val, datetime.date):
        return launch_date_val
    try:
        v = int(launch_date_val)
        return datetime.date(v // 10000, (v // 100) % 100, v % 100)
    except (ValueError, TypeError):
        return None


def get_new_launch_flag(launch_date_val, reference_date=None):
    """Determine if SKU is a new launch within last 3 months."""
    if reference_date is None:
        reference_date = datetime.date.today()
    launch_date = _parse_launch_date(launch_date_val)
    if launch_date is None:
        return ""

    cutoff = reference_date - datetime.timedelta(days=90)

    return "New SKU" if launch_date >= cutoff else ""


def get_launch_days(launch_date_val, reference_date=None):
    """Return days since launch (reference_date - launch_date).

    Returns an int (may be negative if launch_date is in the future) or
    "" if launch_date_val is missing/invalid, to keep Excel cells blank.
    """
    if reference_date is None:
        reference_date = datetime.date.today()
    launch_date = _parse_launch_date(launch_date_val)
    if launch_date is None:
        return ""
    return (reference_date - launch_date).days


def filter_ndrf_data(ws, progress_callback=None):
    """Filter ND RF data based on criteria."""
    raw_headers = [cell.value for cell in ws[1]]
    col_idx = get_col_idx(raw_headers)

    buyer_col_idx = col_idx["Buyer"]
    supply_source_col_idx = col_idx["Supply Source"]
    launch_date_col_idx = col_idx["Launch Date"]
    total_stock_col_idx = col_idx["Total Stock Qty (CM)"]
    adj_safety_stock_col_idx = col_idx["Adj. Safety Stock"]
    sku_desc_col_idx = col_idx["SKU Desc"]

    total_rows = ws.max_row - 1 if ws.max_row is not None else 0
    data_rows = []
    for r, row in enumerate(ws.iter_rows(min_row=2, values_only=True), 2):
        if not row:
            if progress_callback:
                progress_callback(r - 1, total_rows)
            continue
        col_a_val = row[0]
        if col_a_val == "Total" or (isinstance(col_a_val, str) and col_a_val.startswith("Applied filters:")):
            if progress_callback:
                progress_callback(r - 1, total_rows)
            continue

        buyer = str(row[buyer_col_idx]) if row[buyer_col_idx] else ""
        if any(x in buyer for x in ["Marketing", "SSDC", "Dummy", "DotCom"]):
            if progress_callback:
                progress_callback(r - 1, total_rows)
            continue

        sku_desc_val = str(row[sku_desc_col_idx]) if row[sku_desc_col_idx] else ""
        if "(STAFF)" in sku_desc_val or "(MAC)" in sku_desc_val:
            if progress_callback:
                progress_callback(r - 1, total_rows)
            continue

        supply_source = str(row[supply_source_col_idx]) if row[supply_source_col_idx] else ""
        if supply_source == "Warehouse":
            stock_threshold = 100
        elif supply_source == "Vendor":
            stock_threshold = 200
        else:
            if progress_callback:
                progress_callback(r - 1, total_rows)
            continue

        if row[launch_date_col_idx] is None:
            if progress_callback:
                progress_callback(r - 1, total_rows)
            continue

        total_stock = row[total_stock_col_idx]
        if total_stock is None or total_stock < stock_threshold:
            if progress_callback:
                progress_callback(r - 1, total_rows)
            continue

        if row[adj_safety_stock_col_idx] is None:
            if progress_callback:
                progress_callback(r - 1, total_rows)
            continue

        data_rows.append(list(row))
        if progress_callback:
            progress_callback(r - 1, total_rows)

    return raw_headers, data_rows