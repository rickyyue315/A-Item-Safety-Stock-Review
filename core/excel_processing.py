"""Excel processing utilities."""
import io
import datetime
import openpyxl
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
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Export"

    for c, h in enumerate(new_headers, 1):
        ws.cell(row=1, column=c, value=h)

    for r_idx, (raw_values, safety_val) in enumerate(data_rows, 2):
        row = build_row(raw_values, safety_val, raw_headers, col_idx=col_idx)
        for c, val in enumerate(row, 1):
            ws.cell(row=r_idx, column=c, value=val)

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

    wb = openpyxl.Workbook()

    # --- SKU sheet ---
    ws_sku = wb.active
    ws_sku.title = "SKU"

    for c, h in enumerate(new_headers, 1):
        ws_sku.cell(row=1, column=c, value=h)

    for r_idx, (raw_values, safety_val) in enumerate(sorted_data, 2):
        row = build_row(raw_values, safety_val, raw_headers, col_idx=col_idx)
        for c, val in enumerate(row, 1):
            cell = ws_sku.cell(row=r_idx, column=c, value=val)
            col_name = new_headers[c - 1]
            if col_name == "Safety Checking":
                cell.fill = AMBER_FILL
                if safety_val == "Aitem<25%":
                    cell.font = DARK_RED_FONT
            elif col_name in YELLOW_COLUMNS:
                cell.fill = YELLOW_FILL

    sku_widths = {
        "Item": 65.86, "Valuation Class": 25.71, "Buyer": 19.14,
        "Brand": 31.71, "SKU": 13.43, "SKU Desc": 46.29
    }
    for c, h in enumerate(new_headers, 1):
        if h in sku_widths:
            ws_sku.column_dimensions[get_column_letter(c)].width = sku_widths[h]

    ws_sku.freeze_panes = "A2"

    # --- Summary sheet ---
    # Layout (rows are fixed by design; columns are derived from n = #categories):
    #   row 1  : "SASA ABC" group header
    #   row 3  : "No. of SKU" / "Safety VS Monthly Sold Qty" sub-header
    #   row 4  : per-category column headers ("Buyer", <categories...>, "總計")
    #   row 5+ : one row per buyer, then a grand-total row at the end.
    GROUP_HEADER_ROW = 1
    SUBHEADER_ROW = 3
    COLUMN_HEADER_ROW = 4
    DATA_START_ROW = 5

    ws_sum = wb.create_sheet("Summary", 0)

    cnt_buyer = 1
    cnt_start = 2
    cnt_total = cnt_start + n
    pct_buyer = cnt_total + 2
    pct_start = pct_buyer + 1
    pct_total = pct_start + n

    ws_sum.cell(row=GROUP_HEADER_ROW, column=cnt_buyer, value="SASA ABC")
    ws_sum.cell(row=GROUP_HEADER_ROW, column=cnt_start, value="A")
    ws_sum.cell(row=GROUP_HEADER_ROW, column=pct_buyer, value="SASA ABC")
    ws_sum.cell(row=GROUP_HEADER_ROW, column=pct_start, value="A")

    ws_sum.cell(row=SUBHEADER_ROW, column=cnt_buyer, value="No. of SKU")
    ws_sum.cell(row=SUBHEADER_ROW, column=cnt_start, value="Safety VS Monthly Sold Qty")
    ws_sum.cell(row=SUBHEADER_ROW, column=pct_buyer, value="No. of SKU")
    ws_sum.cell(row=SUBHEADER_ROW, column=pct_start, value="Safety VS Monthly Sold Qty")

    header_labels = ["Buyer"] + categories + ["\u7e3d\u8a08"]
    for i, label in enumerate(header_labels):
        ws_sum.cell(row=COLUMN_HEADER_ROW, column=cnt_buyer + i, value=label)
        ws_sum.cell(row=COLUMN_HEADER_ROW, column=pct_buyer + i, value=label)

    ws_sum.cell(row=COLUMN_HEADER_ROW, column=cnt_start).fill = YELLOW_FILL
    ws_sum.cell(row=COLUMN_HEADER_ROW, column=pct_start).fill = YELLOW_FILL

    data_start_row = DATA_START_ROW
    for b_idx, buyer in enumerate(sorted_buyers):
        r = data_start_row + b_idx
        ws_sum.cell(row=r, column=cnt_buyer, value=buyer)
        ws_sum.cell(row=r, column=pct_buyer, value=buyer)

        total = sum(buyer_data[buyer][cat] for cat in categories)
        for c_idx, cat in enumerate(categories):
            count = buyer_data[buyer][cat]
            col_count = cnt_start + c_idx
            col_pct = pct_start + c_idx

            if count > 0:
                ws_sum.cell(row=r, column=col_count, value=count)
            else:
                ws_sum.cell(row=r, column=col_count, value=None)

            pct = count / total if total > 0 else 0
            pct_cell = ws_sum.cell(row=r, column=col_pct, value=pct)
            pct_cell.number_format = PCT_FORMAT

            if c_idx == 0:
                ws_sum.cell(row=r, column=col_count).fill = YELLOW_FILL
                ws_sum.cell(row=r, column=col_pct).fill = YELLOW_FILL

        ws_sum.cell(row=r, column=cnt_total, value=total)
        total_pct_cell = ws_sum.cell(row=r, column=pct_total, value=1.0)
        total_pct_cell.number_format = PCT_FORMAT

    total_row = data_start_row + len(sorted_buyers)
    ws_sum.cell(row=total_row, column=cnt_buyer, value="\u7e3d\u8a08")
    ws_sum.cell(row=total_row, column=pct_buyer, value="\u7e3d\u8a08")

    grand_total = sum(sum(buyer_data[b][cat] for cat in categories) for b in sorted_buyers)
    for c_idx, cat in enumerate(categories):
        cat_total = sum(buyer_data[b][cat] for b in sorted_buyers)
        col_count = cnt_start + c_idx
        col_pct = pct_start + c_idx

        ws_sum.cell(row=total_row, column=col_count, value=cat_total)
        if c_idx == 0:
            ws_sum.cell(row=total_row, column=col_count).fill = YELLOW_FILL

        pct = cat_total / grand_total if grand_total > 0 else 0
        pct_cell = ws_sum.cell(row=total_row, column=col_pct, value=pct)
        pct_cell.number_format = PCT_FORMAT
        if c_idx == 0:
            ws_sum.cell(row=total_row, column=col_pct).fill = YELLOW_FILL

    ws_sum.cell(row=total_row, column=cnt_total, value=grand_total)
    gt_pct = ws_sum.cell(row=total_row, column=pct_total, value=1.0)
    gt_pct.number_format = PCT_FORMAT

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

    wb = openpyxl.Workbook()

    # --- RAW sheet ---
    ws_raw = wb.active
    ws_raw.title = "RAW"
    for c, h in enumerate(final_headers, 1):
        ws_raw.cell(row=1, column=c, value=h)

    nd_ratios_col = final_headers.index("ND ratios") + 1
    for r_idx, (raw_values, nd_ratios, nd_checking, new_launch_flag) in enumerate(enriched_data, 2):
        intermediate_row = list(raw_values)
        intermediate_row.insert(launch_date_idx + 1, new_launch_flag)
        adjusted = nd_shop_stock_idx + (1 if nd_shop_stock_idx > launch_date_idx else 0)
        intermediate_row.insert(adjusted + 1, nd_ratios)
        intermediate_row.insert(adjusted + 2, nd_checking)
        row_dict = dict(zip(intermediate_headers, intermediate_row))
        final_row = [row_dict[h] for h in final_headers]
        for c, val in enumerate(final_row, 1):
            cell = ws_raw.cell(row=r_idx, column=c, value=val)
            if c == nd_ratios_col:
                cell.number_format = "0.0%"

    # --- Summary sheet (placed first) ---
    sorted_buyers = sorted(buyer_counts.keys(), key=extract_buyer_number)
    ws_sum = wb.create_sheet("Summary", 0)
    ws_sum.cell(row=1, column=1, value="Note: Warehouse supply source, Total Stock Qty > 100 units; Vendor supply source, Total Stock Qty > 200 units")
    ws_sum.cell(row=2, column=1, value="'Total' = all filtered SKUs (including ND Ratio < 70%)")

    header_row = 4
    summary_cols = ["Buyer", "ND Ratio > 80%", "ND Ratio 70%-80%",
                    "ND Ratio > 80% %", "ND Ratio 70%-80% %", "Total Filtered SKU"]
    for c, h in enumerate(summary_cols, 1):
        ws_sum.cell(row=header_row, column=c, value=h)

    data_start = header_row + 1
    for i, buyer in enumerate(sorted_buyers):
        r = data_start + i
        bc = buyer_counts[buyer]
        total = bc["Total"]
        gt80 = bc["ND Ratio > 80%"]
        btw = bc["ND Ratio 70%-80%"]
        ws_sum.cell(row=r, column=1, value=buyer)
        ws_sum.cell(row=r, column=2, value=gt80)
        ws_sum.cell(row=r, column=3, value=btw)
        gt80_pct = ws_sum.cell(row=r, column=4, value=gt80 / total if total > 0 else 0)
        gt80_pct.number_format = "0.0%"
        btw_pct = ws_sum.cell(row=r, column=5, value=btw / total if total > 0 else 0)
        btw_pct.number_format = "0.0%"
        ws_sum.cell(row=r, column=6, value=total)

    total_row = data_start + len(sorted_buyers)
    grand_gt80 = sum(b["ND Ratio > 80%"] for b in buyer_counts.values())
    grand_btw = sum(b["ND Ratio 70%-80%"] for b in buyer_counts.values())
    grand_total = sum(b["Total"] for b in buyer_counts.values())
    ws_sum.cell(row=total_row, column=1, value="Grand Total")
    ws_sum.cell(row=total_row, column=2, value=grand_gt80)
    ws_sum.cell(row=total_row, column=3, value=grand_btw)
    gt80_pct_t = ws_sum.cell(row=total_row, column=4, value=grand_gt80 / grand_total if grand_total > 0 else 0)
    gt80_pct_t.number_format = "0.0%"
    btw_pct_t = ws_sum.cell(row=total_row, column=5, value=grand_btw / grand_total if grand_total > 0 else 0)
    btw_pct_t.number_format = "0.0%"
    ws_sum.cell(row=total_row, column=6, value=grand_total)

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
        for c, h in enumerate(ND_PROBLEM_HEADERS, 1):
            ws.cell(row=1, column=c, value=h)
        r = 2
        for raw_values, nd_ratios, nd_checking, new_launch_flag in enriched_data:
            if nd_checking == nd_condition:
                for ci, fi in enumerate(nd_field_indices):
                    ws.cell(row=r, column=ci + 1, value=raw_values[fi])
                nd_cell = ws.cell(row=r, column=len(nd_field_indices) + 1, value=nd_ratios)
                nd_cell.number_format = "0.0%"
                ws.cell(row=r, column=len(nd_field_indices) + 2, value=nd_checking)
                ws.cell(row=r, column=len(nd_field_indices) + 3, value=new_launch_flag)
                r += 1

    return wb