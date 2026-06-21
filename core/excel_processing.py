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
)
from .buyer_utils import extract_buyer_number, get_col_idx
from .nd_logic import get_new_launch_flag


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
    wb = openpyxl.Workbook(write_only=True)
    ws = wb.create_sheet(title="Export")

    ws.append(new_headers)

    for raw_values, safety_val in data_rows:
        ws.append(build_row(raw_values, safety_val, raw_headers, col_idx=col_idx))

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

    wb = openpyxl.Workbook(write_only=True)

    # --- SKU sheet ---
    ws_sku = wb.create_sheet(title="SKU")
    ws_sku.append(new_headers)

    for raw_values, safety_val in sorted_data:
        row_data = build_row(raw_values, safety_val, raw_headers, col_idx=col_idx)
        styled_row = []
        for c, val in enumerate(row_data, 1):
            col_name = new_headers[c - 1]
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
    for c, h in enumerate(new_headers, 1):
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
]


def build_nd_wb(enriched_data, buyer_counts, raw_headers):
    """Build the ND Ratio Review output workbook.

    Produces 4 sheets:
      - Summary     : buyer-level ND Ratio counts/percentages (placed first)
      - ND Ratio > 80%       : problem SKUs
      - ND Ratio 70%-80%     : problem SKUs
      - RAW          : all enriched rows with ND ratios / checking inserted

    ``enriched_data`` is a list of (raw_values, nd_ratios, nd_checking, new_launch_flag).
    ``buyer_counts`` maps buyer -> {"ND Ratio > 80%", "ND Ratio 70%-80%", "Total"}.
    """
    col_idx = get_col_idx(raw_headers)
    buyer_col_idx = col_idx["Buyer"]
    launch_date_idx = col_idx["Launch Date"]
    nd_shop_stock_idx = col_idx["No. of ND Shop with Stock"]

    intermediate_headers = list(raw_headers)
    intermediate_headers.insert(launch_date_idx + 1, "New Launch Flag (L3M)")
    adjusted_nd_idx = nd_shop_stock_idx + (1 if nd_shop_stock_idx > launch_date_idx else 0)
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
    for raw_values, nd_ratios, nd_checking, new_launch_flag in enriched_data:
        intermediate_row = list(raw_values)
        intermediate_row.insert(launch_date_idx + 1, new_launch_flag)
        adjusted = nd_shop_stock_idx + (1 if nd_shop_stock_idx > launch_date_idx else 0)
        intermediate_row.insert(adjusted + 1, nd_ratios)
        intermediate_row.insert(adjusted + 2, nd_checking)
        row_dict = dict(zip(intermediate_headers, intermediate_row))
        final_row = [row_dict[h] for h in final_headers]
        nd_cell = WriteOnlyCell(ws_raw, value=final_row[nd_ratios_col - 1])
        nd_cell.number_format = "0.0%"
        final_row[nd_ratios_col - 1] = nd_cell
        ws_raw.append(final_row)

    # --- Summary sheet (placed first) ---
    sorted_buyers = sorted(buyer_counts.keys(), key=extract_buyer_number)
    ws_sum = wb.create_sheet("Summary", 0)
    ws_sum.append(["Note: Warehouse supply source, Total Stock Qty > 100 units; Vendor supply source, Total Stock Qty > 200 units"])
    ws_sum.append(["'Total' = all filtered SKUs (including ND Ratio < 70%)"])
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
    for sheet_title, nd_condition in [
        ("ND Ratio > 80%", "ND Ratio > 80%"),
        ("ND Ratio 70%-80%", "ND Ratio 70%-80%"),
    ]:
        ws = wb.create_sheet(sheet_title, 1 if sheet_title == "ND Ratio > 80%" else 2)
        ws.append(ND_PROBLEM_HEADERS)
        for raw_values, nd_ratios, nd_checking, new_launch_flag in enriched_data:
            if nd_checking == nd_condition:
                row_data = [raw_values[fi] for fi in nd_field_indices]
                nd_cell = WriteOnlyCell(ws, value=nd_ratios)
                nd_cell.number_format = "0.0%"
                row_data.extend([nd_cell, nd_checking, new_launch_flag])
                ws.append(row_data)

    return wb