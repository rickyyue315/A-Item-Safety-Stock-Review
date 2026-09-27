"""Excel processing utilities."""
import io
import datetime
import openpyxl
from openpyxl.cell import WriteOnlyCell
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

from .safety_logic import (
    SAFETY_ORDER,
    YELLOW_FILL,
    AMBER_FILL,
    DARK_RED_FONT,
    PCT_FORMAT,
    YELLOW_COLUMNS,
    is_aitem_overstock_short_d001,
    is_aitem_overstock_short_d001_tier2,
    is_aitem_overstock_short_d001_tier3,
)
from .buyer_utils import extract_buyer_number, get_col_idx
from .nd_logic import get_new_launch_flag, get_launch_days, DEFAULT_TOTAL_SHOPS


# Sensitive substrings removed from Safety SKU output before external sharing.
# Any header containing "Value", "GP$" or "COGS" (case-insensitive) is dropped.
SAFETY_SENSITIVE_SUBSTRINGS = ("value", "gp$", "cogs")


def is_sensitive_safety_column(header):
    """Return True if a Safety output header is sensitive and must be dropped."""
    if header is None:
        return False
    h = str(header).lower()
    return any(s in h for s in SAFETY_SENSITIVE_SUBSTRINGS)


def filter_safety_headers(new_headers):
    """Return filtered headers with sensitive (Value/GP$/COGS) columns removed."""
    return [h for h in new_headers if not is_sensitive_safety_column(h)]


def build_row(raw_values, safety_val, raw_headers, col_idx=None):
    """Build output row with SASA ABC and Safety Checking columns inserted.

    Pass a pre-built ``col_idx`` (from :func:`get_col_idx`) when calling in a
    loop over many rows; otherwise it is rebuilt on every call.
    """
    if col_idx is None:
        col_idx = get_col_idx(raw_headers)
    row = list(raw_values)
    item_class_pos = col_idx["Item Class"]
    item_class_val = raw_values[item_class_pos]
    sasa_abc = str(item_class_val)[0] if item_class_val else ""
    row.insert(item_class_pos + 1, sasa_abc)

    safety_pos = col_idx["Safety VS Monthly Sold Qty"]
    offset = 1 if item_class_pos < safety_pos else 0
    row.insert(safety_pos + 1 + offset, safety_val)
    return row


def build_intermediate_wb(new_headers, data_rows, raw_headers):
    """Build intermediate workbook with SKU data."""
    col_idx = get_col_idx(raw_headers)
    keep_idx = [i for i, h in enumerate(new_headers) if not is_sensitive_safety_column(h)]
    out_headers = [new_headers[i] for i in keep_idx]
    wb = openpyxl.Workbook(write_only=True)
    ws = wb.create_sheet(title="Export")

    ws.append(out_headers)

    for raw_values, safety_val in data_rows:
        full_row = build_row(raw_values, safety_val, raw_headers, col_idx=col_idx)
        ws.append([full_row[i] for i in keep_idx])

    return wb


def build_final_wb(new_headers, data_rows, raw_headers):
    """Build final workbook with SKU and Summary sheets."""
    col_idx = get_col_idx(raw_headers)
    buyer_col_idx = col_idx["Buyer"]

    categories = sorted(set(sv for _, sv in data_rows), key=lambda x: SAFETY_ORDER.get(x, 99))
    n = len(categories)

    buyer_data = {}
    for raw_values, safety_val in data_rows:
        buyer = raw_values[buyer_col_idx]
        if buyer not in buyer_data:
            buyer_data[buyer] = {cat: 0 for cat in categories}
        buyer_data[buyer][safety_val] += 1

    sorted_buyers = sorted(buyer_data.keys(), key=extract_buyer_number)

    sorted_data = sorted(
        data_rows,
        key=lambda x: (SAFETY_ORDER[x[1]], extract_buyer_number(x[0][buyer_col_idx]))
    )

    # --- Collect problem data for both tiers ---
    ratio_col = col_idx["Safety VS Monthly Sold Qty"]
    item_class_col = col_idx["Item Class"]
    supply_source_col = col_idx["Supply Source"]
    d001_col = col_idx["D001 Net Stock Qty (CM)"]
    sku_col = col_idx["SKU"]
    sku_desc_col = col_idx["SKU Desc"]
    brand_col = col_idx["Brand"]
    safety_stock_col = col_idx["Adj. Safety Stock"]
    shop_monthly_sold_col = col_idx["Shop Monthly Sold Qty (P90D)"]
    shop_stock_col = col_idx["Shop Stock Qty (CM)"]

    problem_data_low = []  # Aitem<25% (understocked A items)
    problem_data_tier1 = []  # Aitem>200% & D001<20pcs
    problem_data_tier2 = []  # Aitem>150% & D001<50pcs
    problem_data_tier3 = []  # Aitem>100% & D001<100pcs

    for raw_values, safety_val in data_rows:
        ratio = raw_values[ratio_col]
        item_class_val = raw_values[item_class_col]
        supply_source = raw_values[supply_source_col]
        d001_net_stock = raw_values[d001_col]

        is_low_aitem = safety_val == "Aitem<25%"
        tier1 = is_aitem_overstock_short_d001(ratio, item_class_val, supply_source, d001_net_stock)
        tier2 = is_aitem_overstock_short_d001_tier2(ratio, item_class_val, supply_source, d001_net_stock)
        tier3 = is_aitem_overstock_short_d001_tier3(ratio, item_class_val, supply_source, d001_net_stock)

        if is_low_aitem or tier1 or tier2 or tier3:
            row_data = {
                "Buyer": raw_values[buyer_col_idx],
                "SKU": raw_values[sku_col],
                "SKU Desc": raw_values[sku_desc_col],
                "Brand": raw_values[brand_col],
                "Supply Source": str(supply_source) if supply_source else "",
                "Safety Stock Qty": raw_values[safety_stock_col],
                "Shop Monthly Sold Qty (P90D)": int(raw_values[shop_monthly_sold_col]) if raw_values[shop_monthly_sold_col] is not None else None,
                "Safety/Monthly %": f"{ratio*100:.1f}%" if ratio is not None else "N/A",
                "D001 Net Stock Qty (CM)": d001_net_stock,
                "Shop Stock Qty (CM)": raw_values[shop_stock_col],
            }
            if is_low_aitem:
                problem_data_low.append(row_data)
            if tier1:
                problem_data_tier1.append(row_data)
            if tier2:
                problem_data_tier2.append(row_data)
            if tier3:
                problem_data_tier3.append(row_data)

    SAFETY_PROBLEM_HEADERS = [
        "Buyer", "SKU", "SKU Desc", "Brand", "Supply Source",
        "Safety Stock Qty", "Shop Monthly Sold Qty (P90D)",
        "Safety/Monthly %", "D001 Net Stock Qty (CM)", "Shop Stock Qty (CM)",
    ]

    wb = openpyxl.Workbook(write_only=True)

    # --- SKU sheet (sensitive Value/GP$ columns removed) ---
    ws_sku = wb.create_sheet(title="SKU")
    keep_idx = [i for i, h in enumerate(new_headers) if not is_sensitive_safety_column(h)]
    out_headers = [new_headers[i] for i in keep_idx]
    ws_sku.append(out_headers)

    for raw_values, safety_val in sorted_data:
        row_data = build_row(raw_values, safety_val, raw_headers, col_idx=col_idx)
        row_data = [row_data[i] for i in keep_idx]
        styled_row = []
        for c, val in enumerate(row_data, 1):
            col_name = out_headers[c - 1]
            if col_name == "Safety Checking":
                cell = WriteOnlyCell(ws_sku, value=val)
                cell.fill = AMBER_FILL
                if safety_val == "Aitem<25%":
                    cell.font = DARK_RED_FONT
                styled_row.append(cell)
            elif col_name in YELLOW_COLUMNS:
                cell = WriteOnlyCell(ws_sku, value=val)
                cell.fill = YELLOW_FILL
                styled_row.append(cell)
            else:
                styled_row.append(val)
        ws_sku.append(styled_row)

    sku_widths = {
        "Item": 65.86, "Valuation Class": 25.71, "Buyer": 19.14,
        "Brand": 31.71, "SKU": 13.43, "SKU Desc": 46.29
    }
    for c, h in enumerate(out_headers, 1):
        if h in sku_widths:
            ws_sku.column_dimensions[get_column_letter(c)].width = sku_widths[h]

    ws_sku.freeze_panes = "A2"

    # --- Summary sheet ---
    ws_sum = wb.create_sheet("Summary", 0)

    cnt_buyer = 1
    cnt_start = 2
    cnt_total = cnt_start + n
    pct_buyer = cnt_total + 2
    pct_start = pct_buyer + 1
    pct_total = pct_start + n
    max_col = pct_total

    first_cat = categories[0] if categories else ""
    group_label = first_cat[0] if first_cat else "A"

    # Row 1: group header
    row1 = [None] * max_col
    row1[cnt_buyer - 1] = "SASA ABC"
    row1[cnt_start - 1] = group_label
    row1[pct_buyer - 1] = "SASA ABC"
    row1[pct_start - 1] = group_label
    ws_sum.append(row1)

    # Row 2: blank
    ws_sum.append([])

    # Row 3: sub-header
    row3 = [None] * max_col
    row3[cnt_buyer - 1] = "No. of SKU"
    row3[cnt_start - 1] = "Safety VS Monthly Sold Qty"
    row3[pct_buyer - 1] = "No. of SKU"
    row3[pct_start - 1] = "Safety VS Monthly Sold Qty"
    ws_sum.append(row3)

    # Row 4: column headers
    header_labels = ["Buyer"] + categories + ["\u7e3d\u8a08"]
    row4 = [None] * max_col
    for i, label in enumerate(header_labels):
        row4[cnt_buyer - 1 + i] = label
        row4[pct_buyer - 1 + i] = label
    for col in [cnt_start, pct_start]:
        cell = WriteOnlyCell(ws_sum, value=row4[col - 1])
        cell.fill = YELLOW_FILL
        row4[col - 1] = cell
    ws_sum.append(row4)

    # Data rows
    for b_idx, buyer in enumerate(sorted_buyers):
        row = [None] * max_col
        row[cnt_buyer - 1] = buyer
        row[pct_buyer - 1] = buyer

        total = sum(buyer_data[buyer][cat] for cat in categories)
        for c_idx, cat in enumerate(categories):
            count = buyer_data[buyer][cat]
            col_count = cnt_start + c_idx
            col_pct = pct_start + c_idx

            if count > 0:
                if c_idx == 0:
                    cell = WriteOnlyCell(ws_sum, value=count)
                    cell.fill = YELLOW_FILL
                    row[col_count - 1] = cell
                else:
                    row[col_count - 1] = count

            pct = count / total if total > 0 else 0
            pct_cell = WriteOnlyCell(ws_sum, value=pct)
            pct_cell.number_format = PCT_FORMAT
            if c_idx == 0:
                pct_cell.fill = YELLOW_FILL
            row[col_pct - 1] = pct_cell

        row[cnt_total - 1] = total
        total_pct_cell = WriteOnlyCell(ws_sum, value=1.0)
        total_pct_cell.number_format = PCT_FORMAT
        row[pct_total - 1] = total_pct_cell
        ws_sum.append(row)

    # Grand total row
    total_row = [None] * max_col
    total_row[cnt_buyer - 1] = "\u7e3d\u8a08"
    total_row[pct_buyer - 1] = "\u7e3d\u8a08"

    grand_total = sum(sum(buyer_data[b][cat] for cat in categories) for b in sorted_buyers)
    for c_idx, cat in enumerate(categories):
        cat_total = sum(buyer_data[b][cat] for b in sorted_buyers)
        col_count = cnt_start + c_idx
        col_pct = pct_start + c_idx

        if c_idx == 0:
            cell = WriteOnlyCell(ws_sum, value=cat_total)
            cell.fill = YELLOW_FILL
            total_row[col_count - 1] = cell
        else:
            total_row[col_count - 1] = cat_total

        pct = cat_total / grand_total if grand_total > 0 else 0
        pct_cell = WriteOnlyCell(ws_sum, value=pct)
        pct_cell.number_format = PCT_FORMAT
        if c_idx == 0:
            pct_cell.fill = YELLOW_FILL
        total_row[col_pct - 1] = pct_cell

    total_row[cnt_total - 1] = grand_total
    gt_pct = WriteOnlyCell(ws_sum, value=1.0)
    gt_pct.number_format = PCT_FORMAT
    total_row[pct_total - 1] = gt_pct
    ws_sum.append(total_row)

    sum_widths = {get_column_letter(cnt_buyer): 18.43}
    for i in range(n):
        sum_widths[get_column_letter(cnt_start + i)] = 15.14 if i == 0 else 11.0
    sum_widths[get_column_letter(cnt_total)] = 11.0
    sum_widths[get_column_letter(pct_buyer)] = 18.43
    for i in range(n):
        sum_widths[get_column_letter(pct_start + i)] = 15.14 if i == 0 else 11.0
    sum_widths[get_column_letter(pct_total)] = 8.14
    for col_letter, w in sum_widths.items():
        ws_sum.column_dimensions[col_letter].width = w

    # --- Safety Stock Check Tier sheets ---
    safety_problem_widths = {
        "Buyer": 18.43, "SKU": 13.43, "SKU Desc": 46.29, "Brand": 31.71,
        "Supply Source": 18.0, "Safety Stock Qty": 16.0,
        "Shop Monthly Sold Qty (P90D)": 22.0, "Safety/Monthly %": 16.0,
        "D001 Net Stock Qty (CM)": 20.0, "Shop Stock Qty (CM)": 20.0,
    }

    for sheet_title, problem_data in [
        ("Aitem<25%", problem_data_low),
        ("Aitem>200% & D001<20pcs", problem_data_tier1),
        ("Aitem>150% & D001<50pcs", problem_data_tier2),
        ("Aitem>100% & D001<100pcs", problem_data_tier3),
    ]:
        ws_prob = wb.create_sheet(sheet_title)
        ws_prob.append(SAFETY_PROBLEM_HEADERS)
        for row_dict in problem_data:
            ws_prob.append([row_dict.get(h, "") for h in SAFETY_PROBLEM_HEADERS])
        for c, h in enumerate(SAFETY_PROBLEM_HEADERS, 1):
            if h in safety_problem_widths:
                ws_prob.column_dimensions[get_column_letter(c)].width = safety_problem_widths[h]
        ws_prob.freeze_panes = "A2"

    return wb


def wb_to_bytes(wb):
    """Convert workbook to bytes."""
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


# Columns dropped / reordered when producing the ND Ratio Review output file.
ND_COLUMNS_TO_DELETE = {
    "WH Stock Value (CM)", "Shop Stock Value (CM)", "Total Stock Value (CM)",
    "P90D Daily COGS", "GP$ (P90D)", "GP$ (M-2)", "GP$ (M-1)", "GP$ (CM - MTD)",
}
ND_COLUMNS_TO_MOVE = ["D001 Net Stock Qty (CM)", "Shop Stock Qty (CM)", "Total Stock Qty (CM)"]

# Header row for the two ND problem sheets.
ND_PROBLEM_HEADERS = [
    "Buyer", "SKU", "SKU Desc", "Brand", "No. of RF Shop",
    "No. of ND Shop with Stock", "ND ratios", "ND Checking", "New Launch Flag (L3M)",
    "Days Since Launch",
]

ND_PROBLEM_WIDTHS = {
    "Buyer": 18.43, "SKU": 13.43, "SKU Desc": 46.29, "Brand": 31.71,
    "No. of RF Shop": 16, "No. of ND Shop with Stock": 22,
    "ND ratios": 12, "ND Checking": 18,
    "New Launch Flag (L3M)": 18, "Days Since Launch": 18,
}


def build_nd_wb(enriched_data, buyer_counts, raw_headers, total_shops=DEFAULT_TOTAL_SHOPS, reference_date=None):
    """Build the ND Ratio Review output workbook.

    Produces 4 sheets:
      - Summary     : buyer-level ND Ratio counts/percentages (placed first)
      - ND Ratio > 80%       : problem SKUs
      - ND Ratio 70%-80%     : problem SKUs
      - RAW          : all enriched rows with ND ratios / checking inserted

    ``enriched_data`` is a list of
      (raw_values, nd_ratios, nd_checking, new_launch_flag, launch_days).
    ``buyer_counts`` maps buyer -> {"ND Ratio > 80%", "ND Ratio 70%-80%", "Total"}.

    ``launch_days`` is an int (days since launch) or "" if unknown; it is
    inserted as "Days Since Launch" immediately after "New Launch Flag (L3M)"
    so both columns sit next to the original Launch Date.

    ``reference_date`` is used only as fallback when ``enriched_data`` is a
    legacy 4-tuple without launch_days; otherwise days are taken from the tuple.
    """
    col_idx = get_col_idx(raw_headers)
    buyer_col_idx = col_idx["Buyer"]
    launch_date_idx = col_idx["Launch Date"]
    nd_shop_stock_idx = col_idx["No. of ND Shop with Stock"]

    intermediate_headers = list(raw_headers)
    intermediate_headers.insert(launch_date_idx + 1, "New Launch Flag (L3M)")
    intermediate_headers.insert(launch_date_idx + 2, "Days Since Launch")
    adjusted_nd_idx = nd_shop_stock_idx + (2 if nd_shop_stock_idx > launch_date_idx else 0)
    intermediate_headers.insert(adjusted_nd_idx + 1, "ND ratios")
    intermediate_headers.insert(adjusted_nd_idx + 2, "ND Checking")

    final_headers = [h for h in intermediate_headers if h not in ND_COLUMNS_TO_DELETE]
    for col in ND_COLUMNS_TO_MOVE:
        final_headers.remove(col)
    nd_checking_pos = final_headers.index("ND Checking")
    for i, col in enumerate(ND_COLUMNS_TO_MOVE):
        final_headers.insert(nd_checking_pos + 1 + i, col)

    wb = openpyxl.Workbook(write_only=True)

    # --- RAW sheet ---
    ws_raw = wb.create_sheet(title="RAW")
    ws_raw.append(final_headers)

    nd_ratios_col = final_headers.index("ND ratios") + 1
    # Support both 4-tuple (legacy) and 5-tuple enriched_data for back-compat.
    for item in enriched_data:
        if len(item) == 5:
            raw_values, nd_ratios, nd_checking, new_launch_flag, launch_days = item
        else:
            raw_values, nd_ratios, nd_checking, new_launch_flag = item
            # Fallback: compute from launch date if days not supplied
            launch_days = get_launch_days(raw_values[launch_date_idx], reference_date)
        intermediate_row = list(raw_values)
        intermediate_row.insert(launch_date_idx + 1, new_launch_flag)
        intermediate_row.insert(launch_date_idx + 2, launch_days)
        adjusted = nd_shop_stock_idx + (2 if nd_shop_stock_idx > launch_date_idx else 0)
        intermediate_row.insert(adjusted + 1, nd_ratios)
        intermediate_row.insert(adjusted + 2, nd_checking)
        row_dict = dict(zip(intermediate_headers, intermediate_row))
        final_row = [row_dict[h] for h in final_headers]
        nd_cell = WriteOnlyCell(ws_raw, value=final_row[nd_ratios_col - 1])
        nd_cell.number_format = "0.0%"
        final_row[nd_ratios_col - 1] = nd_cell
        ws_raw.append(final_row)

    # optional: RAW column widths for new columns (write_only still respects widths)
    for c, h in enumerate(final_headers, 1):
        if h in ("New Launch Flag (L3M)", "Days Since Launch"):
            ws_raw.column_dimensions[get_column_letter(c)].width = 18
        elif h in ("Launch Date",):
            ws_raw.column_dimensions[get_column_letter(c)].width = 14
    ws_raw.freeze_panes = "A2"

    # --- Summary sheet (placed first) ---
    sorted_buyers = sorted(buyer_counts.keys(), key=extract_buyer_number)
    ws_sum = wb.create_sheet("Summary", 0)
    ws_sum.append(["Note: Warehouse supply source, Total Stock Qty > 100 units; Vendor supply source, Total Stock Qty > 200 units"])
    ws_sum.append(["'Total' = all filtered SKUs (including ND Ratio < 70%)"])
    ws_sum.append([f"Current Total Shops = {total_shops}"])
    ws_sum.append([])
    summary_cols = ["Buyer", "ND Ratio > 80%", "ND Ratio 70%-80%",
                    "ND Ratio > 80% %", "ND Ratio 70%-80% %", "Total Filtered SKU"]
    ws_sum.append(summary_cols)

    for buyer in sorted_buyers:
        bc = buyer_counts[buyer]
        total = bc["Total"]
        gt80 = bc["ND Ratio > 80%"]
        btw = bc["ND Ratio 70%-80%"]
        gt80_pct = WriteOnlyCell(ws_sum, value=gt80 / total if total > 0 else 0)
        gt80_pct.number_format = "0.0%"
        btw_pct = WriteOnlyCell(ws_sum, value=btw / total if total > 0 else 0)
        btw_pct.number_format = "0.0%"
        ws_sum.append([buyer, gt80, btw, gt80_pct, btw_pct, total])

    grand_gt80 = sum(b["ND Ratio > 80%"] for b in buyer_counts.values())
    grand_btw = sum(b["ND Ratio 70%-80%"] for b in buyer_counts.values())
    grand_total = sum(b["Total"] for b in buyer_counts.values())
    gt80_pct_t = WriteOnlyCell(ws_sum, value=grand_gt80 / grand_total if grand_total > 0 else 0)
    gt80_pct_t.number_format = "0.0%"
    btw_pct_t = WriteOnlyCell(ws_sum, value=grand_btw / grand_total if grand_total > 0 else 0)
    btw_pct_t.number_format = "0.0%"
    ws_sum.append(["Grand Total", grand_gt80, grand_btw, gt80_pct_t, btw_pct_t, grand_total])

    # --- Two ND problem sheets ---
    nd_field_indices = [
        col_idx["Buyer"], col_idx["SKU"], col_idx["SKU Desc"],
        col_idx["Brand"], col_idx["No. of RF Shop"],
        col_idx["No. of ND Shop with Stock"],
    ]
    # launch_date index for fallback when enriched_data is legacy 4-tuple
    _launch_idx_for_days = launch_date_idx
    problem_sheets = []
    for sheet_title, nd_condition in [
        ("ND Ratio > 80%", "ND Ratio > 80%"),
        ("ND Ratio 70%-80%", "ND Ratio 70%-80%"),
    ]:
        ws = wb.create_sheet(sheet_title, 1 if sheet_title == "ND Ratio > 80%" else 2)
        ws.append(ND_PROBLEM_HEADERS)
        for item in enriched_data:
            if len(item) == 5:
                raw_values, nd_ratios, nd_checking, new_launch_flag, launch_days = item
            else:
                raw_values, nd_ratios, nd_checking, new_launch_flag = item
                launch_days = get_launch_days(raw_values[_launch_idx_for_days], reference_date)
            if nd_checking == nd_condition:
                row_data = [raw_values[fi] for fi in nd_field_indices]
                nd_cell = WriteOnlyCell(ws, value=nd_ratios)
                nd_cell.number_format = "0.0%"
                row_data.extend([nd_cell, nd_checking, new_launch_flag, launch_days])
                ws.append(row_data)
        # set column widths & freeze
        for c, h in enumerate(ND_PROBLEM_HEADERS, 1):
            if h in ND_PROBLEM_WIDTHS:
                ws.column_dimensions[get_column_letter(c)].width = ND_PROBLEM_WIDTHS[h]
        ws.freeze_panes = "A2"
        problem_sheets.append(ws)

    # optional: set widths for Summary (already handled) and ensure RAW freeze already set

    return wb