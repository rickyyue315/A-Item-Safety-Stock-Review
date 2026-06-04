import datetime
import re
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
from copy import copy

INPUT_FILE = "By SKU_RAW.xlsx"
INTERMEDIATE_FILE = "By SKU.xlsx"
SAFETY_ORDER = {
    "Aitem<25%": 0, "New item<15%": 1, "B/C/D item<15%": 2,
    "Aitem>75%": 3, "New item>100%": 4, "B/C/D item>100%": 5,
    "OK": 6, "Check": 7
}

YELLOW_FILL = PatternFill(start_color="FFFFFF00", end_color="FFFFFF00", fill_type="solid")
AMBER_FILL = PatternFill(start_color="FFFFC000", end_color="FFFFC000", fill_type="solid")
DARK_RED_FONT = Font(color="8B0000")
PCT_FORMAT = "0.00%"
YELLOW_COLUMNS = {"Adj. Safety Stock", "Shop Monthly Sold Qty (P90D)", "Safety VS Monthly Sold Qty"}


def get_safety_checking(ratio, sasa_abc):
    if sasa_abc == "A":
        if ratio is None or ratio < 0.25:
            return "Aitem<25%"
        if ratio > 0.75:
            return "Aitem>75%"
        return "OK"
    elif sasa_abc in ("B", "C", "D", "N", ""):
        if sasa_abc in ("N", ""):
            if ratio is None or ratio < 0.15:
                return "New item<15%"
            if ratio > 1.0:
                return "New item>100%"
            return "OK"
        else:
            if ratio is None or ratio < 0.15:
                return "B/C/D item<15%"
            if ratio > 1.0:
                return "B/C/D item>100%"
            return "OK"
    else:
        return "Check"


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
    item_class_col = headers.index("Item Class") + 1
    sku_desc_col = headers.index("SKU Desc") + 1

    data_rows = []
    for r in range(2, ws.max_row + 1):
        col_a_val = ws.cell(row=r, column=1).value
        if col_a_val == "Total" or (isinstance(col_a_val, str) and col_a_val.startswith("Applied filters:")):
            continue
        item_val = ws.cell(row=r, column=item_col).value
        buyer_val = ws.cell(row=r, column=buyer_col).value
        if buyer_val is None or item_val is None:
            continue
        sku_desc_val = str(ws.cell(row=r, column=sku_desc_col).value or "")
        if "(STAFF)" in sku_desc_val or "(MAC)" in sku_desc_val:
            continue
        row_cells = []
        for c in range(1, ws.max_column + 1):
            row_cells.append(ws.cell(row=r, column=c))
        ratio_val = ws.cell(row=r, column=ratio_col).value
        item_class_cell = ws.cell(row=r, column=item_class_col)
        item_class_val = item_class_cell.value
        sasa_abc = str(item_class_val)[0] if item_class_val else ""
        safety_val = get_safety_checking(ratio_val, sasa_abc)
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
    item_class_cell = raw_row_cells[sasa_idx_v]
    item_class_val = str(item_class_cell.value) if item_class_cell.value else ""
    sasa_abc = item_class_val[0] if item_class_val else ""
    values.insert(sasa_idx_v + 1, ("SASA ABC", sasa_abc))

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
                ws.cell(row=r_idx, column=c, value=val_or_cell)
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
    buyer_col_name = "Buyer"
    buyer_col_idx = raw_headers.index(buyer_col_name)

    categories = sorted(set(sv for _, sv in data_rows), key=lambda x: SAFETY_ORDER.get(x, 99))
    n = len(categories)

    buyer_data = {}
    for raw_cells, safety_val in data_rows:
        buyer = raw_cells[buyer_col_idx].value
        if buyer not in buyer_data:
            buyer_data[buyer] = {cat: 0 for cat in categories}
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
                cell = ws_sku.cell(row=r_idx, column=c, value=val_or_cell)
            elif col_name == "Safety Checking":
                cell = ws_sku.cell(row=r_idx, column=c, value=safety_val)
                cell.fill = AMBER_FILL
                if safety_val == "Aitem<25%":
                    cell.font = DARK_RED_FONT
            else:
                cell_src = val_or_cell
                cell = ws_sku.cell(row=r_idx, column=c, value=cell_src.value)
                if col_name in YELLOW_COLUMNS:
                    cell.fill = YELLOW_FILL

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

    cnt_buyer = 1
    cnt_start = 2
    cnt_total = cnt_start + n
    pct_buyer = cnt_total + 2
    pct_start = pct_buyer + 1
    pct_total = pct_start + n

    ws_sum.cell(row=1, column=cnt_buyer, value="SASA ABC")
    ws_sum.cell(row=1, column=cnt_start, value="A")
    ws_sum.cell(row=1, column=pct_buyer, value="SASA ABC")
    ws_sum.cell(row=1, column=pct_start, value="A")

    ws_sum.cell(row=3, column=cnt_buyer, value="No. of SKU")
    ws_sum.cell(row=3, column=cnt_start, value="Safety VS Monthly Sold Qty")
    ws_sum.cell(row=3, column=pct_buyer, value="No. of SKU")
    ws_sum.cell(row=3, column=pct_start, value="Safety VS Monthly Sold Qty")

    header_labels = ["Buyer"] + categories + ["\u7e3d\u8a08"]
    for i, label in enumerate(header_labels):
        ws_sum.cell(row=4, column=cnt_buyer + i, value=label)
        ws_sum.cell(row=4, column=pct_buyer + i, value=label)

    ws_sum.cell(row=4, column=cnt_start).fill = YELLOW_FILL
    ws_sum.cell(row=4, column=pct_start).fill = YELLOW_FILL

    data_start_row = 5
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
