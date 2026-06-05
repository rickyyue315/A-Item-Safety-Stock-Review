"""ND Ratio checking logic."""
import datetime
import openpyxl
from .buyer_utils import get_col_idx


def compute_nd_checking(no_of_rf_shop, total_shops):
    """Compute ND ratio and checking category."""
    if no_of_rf_shop is None:
        nd_ratios = 1.0
    elif no_of_rf_shop > total_shops:
        nd_ratios = 0.0
    else:
        nd_ratios = (total_shops - no_of_rf_shop) / total_shops if total_shops > 0 else 0

    if nd_ratios > 0.8:
        nd_checking = "ND Ratio > 80%"
    elif nd_ratios >= 0.7:
        nd_checking = "ND Ratio 70%-80%"
    else:
        nd_checking = None

    return nd_ratios, nd_checking


def get_new_launch_flag(launch_date_val, reference_date=None):
    """Determine if SKU is a new launch within last 3 months."""
    if reference_date is None:
        reference_date = datetime.date.today()
    if launch_date_val is None:
        return ""
    try:
        launch_date_val = int(launch_date_val)
        launch_date = datetime.date(
            launch_date_val // 10000,
            (launch_date_val // 100) % 100,
            launch_date_val % 100
        )
    except (ValueError, TypeError):
        return ""

    past_month = reference_date.month - 3
    past_year = reference_date.year
    if past_month <= 0:
        past_month += 12
        past_year -= 1
    cutoff = datetime.date(past_year, past_month, 1)

    return "New SKU" if launch_date >= cutoff else ""


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

    total_rows = ws.max_row - 1
    data_rows = []
    for r, row in enumerate(ws.iter_rows(min_row=2, values_only=True), 2):
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
        if supply_source != "Warehouse":
            if progress_callback:
                progress_callback(r - 1, total_rows)
            continue

        if row[launch_date_col_idx] is None:
            if progress_callback:
                progress_callback(r - 1, total_rows)
            continue

        total_stock = row[total_stock_col_idx]
        if total_stock is None or total_stock < 100:
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