import datetime
import re
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
from copy import copy

INPUT_FILE = "By SKU_RAW.xlsx"
INTERMEDIATE_FILE = "By SKU.xlsx"
SAFETY_ORDER = {"Aitem<25%": 0, "Aitem>75%": 1, "OK": 2}

YELLOW_FILL = PatternFill(start_color="FFFFFF00", end_color="FFFFFF00", fill_type="solid")
AMBER_FILL = PatternFill(start_color="FFFFC000", end_color="FFFFC000", fill_type="solid")
PCT_FORMAT = "0.00%"


def get_safety_checking(ratio):
    if ratio is None:
        return "Aitem<25%"
    if ratio < 0.25:
        return "Aitem<25%"
    if ratio >= 0.75:
        return "Aitem>75%"
    return "OK"


def extract_buyer_number(buyer_str):
    m = re.match(r"(\d+)", str(buyer_str))
    return int(m.group(1)) if m else 0


def copy_cell(src, dst):
    dst.value = src.value
    if src.has_style:
        dst.font = copy(src.font)
        dst.border = copy(src.border)
        dst.fill = copy(src.fill)
        dst.number_format = copy(src.number_format)
        dst.protection = copy(src.protection)
        dst.alignment = copy(src.alignment)


def read_raw_data():
    wb = openpyxl.load_workbook(INPUT_FILE, data_only=True)
    ws = wb["Export"]
    headers = [cell.value for cell in ws[1]]

    item_col = headers.index("Item") + 1
    buyer_col = headers.index("Buyer") + 1
    ratio_col = headers.index("Safety VS Monthly Sold Qty") + 1

    data_rows = []
    for r in range(2, ws.max_row + 1):
        col_a_val = ws.cell(row=r, column=1).value
        if col_a_val == "Total" or (isinstance(col_a_val, str) and col_a_val.startswith("Applied filters:")):
            continue
        item_val = ws.cell(row=r, column=item_col).value
        buyer_val = ws.cell(row=r, column=buyer_col).value
        if buyer_val is None:
            continue
        row_cells = []
        for c in range(1, ws.max_column + 1):
            row_cells.append(ws.cell(row=r, column=c))
        ratio_val = ws.cell(row=r, column=ratio_col).value
        safety_val = get_safety_checking(ratio_val)
        data_rows.append((row_cells, safety_val))

    wb.close()
    return headers, data_rows


def build_new_headers(raw_headers):
    new_headers = list(raw_headers)
    item_class_idx = raw_headers.index("Item Class")
    new_headers.insert(item_class_idx + 1, "SASA ABC")
    updated_headers = list(new_headers)
    safety_vs_idx = updated_headers.index("Safety VS Monthly Sold Qty")
    updated_headers.insert(safety_vs_idx + 1, "Safety Checking")
    return updated_headers


def build_row_data(raw_row_cells, safety_val, raw_headers):
    sasa_idx = raw_headers.index("Item Class")
    safety_vs_idx_raw = raw_headers.index("Safety VS Monthly Sold Qty")

    values = []
    for i, cell in enumerate(raw_row_cells):
        col_name = raw_headers[i]
        values.append((col_name, cell))

    sasa_idx_v = raw_headers.index("Item Class")
    values.insert(sasa_idx_v + 1, ("SASA ABC", "A"))

    updated_col_names = [v[0] for v in values]
    safety_vs_idx_v = updated_col_names.index("Safety VS Monthly Sold Qty")
    values.insert(safety_vs_idx_v + 1, ("Safety Checking", safety_val))

    return values


def write_intermediate(headers, data_rows, raw_headers):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Export"

    for c, h in enumerate(headers, 1):
        ws.cell(row=1, column=c, value=h)

    for r_idx, (raw_cells, safety_val) in enumerate(data_rows, 2):
        row_values = build_row_data(raw_cells, safety_val, raw_headers)
        for c, (col_name, val_or_cell) in enumerate(row_values, 1):
            if col_name == "SASA ABC":
                ws.cell(row=r_idx, column=c, value="A")
            elif col_name == "Safety Checking":
                cell = ws.cell(row=r_idx, column=c, value=safety_val)
                cell.fill = AMBER_FILL
            else:
                cell_src = val_or_cell
                ws.cell(row=r_idx, column=c, value=cell_src.value)

    wb.save(INTERMEDIATE_FILE)
    wb.close()
    print(f"Saved {INTERMEDIATE_FILE}")


def write_final(headers, data_rows, raw_headers):
    safety_check_col = headers.index("Safety Checking") + 1
    buyer_col_name = "Buyer"

    buyer_data = {}
    for raw_cells, safety_val in data_rows:
        row_dict = {}
        for i, cell in enumerate(raw_cells):
            row_dict[raw_headers[i]] = cell.value
        buyer = row_dict.get(buyer_col_name, "")
        if buyer not in buyer_data:
            buyer_data[buyer] = {"Aitem<25%": 0, "Aitem>75%": 0, "OK": 0}
        buyer_data[buyer][safety_val] += 1

    sorted_buyers = sorted(buyer_data.keys(), key=extract_buyer_number)

    sorted_data = sorted(
        data_rows,
        key=lambda x: (SAFETY_ORDER[x[1]], extract_buyer_number(
            next(c.value for i, c in enumerate(x[0]) if raw_headers[i] == "Buyer")
        ))
    )

    wb = openpyxl.Workbook()

    # --- SKU sheet ---
    ws_sku = wb.active
    ws_sku.title = "SKU"

    for c, h in enumerate(headers, 1):
        ws_sku.cell(row=1, column=c, value=h)

    for r_idx, (raw_cells, safety_val) in enumerate(sorted_data, 2):
        row_values = build_row_data(raw_cells, safety_val, raw_headers)
        for c, (col_name, val_or_cell) in enumerate(row_values, 1):
            if col_name == "SASA ABC":
                ws_sku.cell(row=r_idx, column=c, value="A")
            elif col_name == "Safety Checking":
                cell = ws_sku.cell(row=r_idx, column=c, value=safety_val)
                cell.fill = AMBER_FILL
            else:
                cell_src = val_or_cell
                ws_sku.cell(row=r_idx, column=c, value=cell_src.value)

    sku_widths = {
        "Item": 65.86, "Valuation Class": 25.71, "Buyer": 19.14,
        "Brand": 31.71, "SKU": 13.43, "SKU Desc": 46.29
    }
    for c, h in enumerate(headers, 1):
        if h in sku_widths:
            ws_sku.column_dimensions[get_column_letter(c)].width = sku_widths[h]

    ws_sku.freeze_panes = "A2"

    # --- Summary sheet ---
    ws_sum = wb.create_sheet("Summary", 0)

    ws_sum.cell(row=1, column=1, value="SASA ABC")
    ws_sum.cell(row=1, column=2, value="A")
    ws_sum.cell(row=1, column=8, value="SASA ABC")
    ws_sum.cell(row=1, column=9, value="A")

    ws_sum.cell(row=3, column=1, value="No. of SKU")
    ws_sum.cell(row=3, column=2, value="Safety VS Monthly Sold Qty")
    ws_sum.cell(row=3, column=8, value="No. of SKU")
    ws_sum.cell(row=3, column=9, value="Safety VS Monthly Sold Qty")

    header_labels = ["Buyer", "Aitem<25%", "Aitem>75%", "OK", "\u7e3d\u8a08"]
    for i, label in enumerate(header_labels):
        ws_sum.cell(row=4, column=1 + i, value=label)
        ws_sum.cell(row=4, column=8 + i, value=label)

    ws_sum.cell(row=4, column=2).fill = YELLOW_FILL
    ws_sum.cell(row=4, column=9).fill = YELLOW_FILL

    categories = ["Aitem<25%", "Aitem>75%", "OK"]

    data_start_row = 5
    for b_idx, buyer in enumerate(sorted_buyers):
        r = data_start_row + b_idx
        ws_sum.cell(row=r, column=1, value=buyer)
        ws_sum.cell(row=r, column=8, value=buyer)

        total = sum(buyer_data[buyer][cat] for cat in categories)
        for c_idx, cat in enumerate(categories):
            count = buyer_data[buyer][cat]
            col_count = 2 + c_idx
            col_pct = 9 + c_idx

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

        ws_sum.cell(row=r, column=5, value=total)
        total_pct_cell = ws_sum.cell(row=r, column=12, value=1.0)
        total_pct_cell.number_format = PCT_FORMAT

    total_row = data_start_row + len(sorted_buyers)
    ws_sum.cell(row=total_row, column=1, value="\u7e3d\u8a08")
    ws_sum.cell(row=total_row, column=8, value="\u7e3d\u8a08")

    grand_total = sum(sum(buyer_data[b][cat] for cat in categories) for b in sorted_buyers)
    for c_idx, cat in enumerate(categories):
        cat_total = sum(buyer_data[b][cat] for b in sorted_buyers)
        col_count = 2 + c_idx
        col_pct = 9 + c_idx

        ws_sum.cell(row=total_row, column=col_count, value=cat_total)
        if c_idx == 0:
            ws_sum.cell(row=total_row, column=col_count).fill = YELLOW_FILL

        pct = cat_total / grand_total if grand_total > 0 else 0
        pct_cell = ws_sum.cell(row=total_row, column=col_pct, value=pct)
        pct_cell.number_format = PCT_FORMAT
        if c_idx == 0:
            ws_sum.cell(row=total_row, column=col_pct).fill = YELLOW_FILL

    ws_sum.cell(row=total_row, column=5, value=grand_total)
    gt_pct = ws_sum.cell(row=total_row, column=12, value=1.0)
    gt_pct.number_format = PCT_FORMAT

    sum_widths = {"A": 18.43, "B": 15.14, "C": 11.0, "H": 18.43, "I": 15.14, "J": 11.0, "K": 7.14, "L": 8.14}
    for col_letter, w in sum_widths.items():
        ws_sum.column_dimensions[col_letter].width = w

    today = datetime.date.today()
    date_str = today.strftime("%d%b%Y")
    output_file = f"Check Safety A Item by SKU_To Buyer_{date_str}.xlsx"
    wb.save(output_file)
    wb.close()
    print(f"Saved {output_file}")


def main():
    print(f"Reading {INPUT_FILE}...")
    raw_headers, data_rows = read_raw_data()
    print(f"Found {len(data_rows)} data rows")

    new_headers = build_new_headers(raw_headers)
    print(f"New headers: {len(new_headers)} columns")

    write_intermediate(new_headers, data_rows, raw_headers)
    write_final(new_headers, data_rows, raw_headers)
    print("Done!")


if __name__ == "__main__":
    main()
