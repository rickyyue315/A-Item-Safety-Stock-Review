import datetime
import openpyxl

from core import (
    SAFETY_ORDER,
    YELLOW_FILL,
    AMBER_FILL,
    DARK_RED_FONT,
    PCT_FORMAT,
    YELLOW_COLUMNS,
    get_safety_checking,
    build_intermediate_wb,
    build_final_wb,
)

INPUT_FILE = "By SKU_RAW.xlsx"
INTERMEDIATE_FILE = "By SKU.xlsx"


def read_raw_data():
    wb = openpyxl.load_workbook(INPUT_FILE, data_only=True)
    ws = wb["Export"]
    headers = [cell.value for cell in ws[1]]

    item_col = headers.index("Item")
    buyer_col = headers.index("Buyer")
    ratio_col = headers.index("Safety VS Monthly Sold Qty")
    item_class_col = headers.index("Item Class")
    sku_desc_col = headers.index("SKU Desc")

    data_rows = []
    for row_values in ws.iter_rows(min_row=2, values_only=True):
        col_a_val = row_values[0]
        if col_a_val == "Total" or (isinstance(col_a_val, str) and col_a_val.startswith("Applied filters:")):
            continue
        item_val = row_values[item_col]
        buyer_val = row_values[buyer_col]
        if buyer_val is None or item_val is None:
            continue
        sku_desc_val = str(row_values[sku_desc_col] or "")
        if "(STAFF)" in sku_desc_val or "(MAC)" in sku_desc_val:
            continue
        ratio_val = row_values[ratio_col]
        item_class_val = row_values[item_class_col]
        sasa_abc = str(item_class_val)[0] if item_class_val else ""
        safety_val = get_safety_checking(ratio_val, sasa_abc)
        data_rows.append((list(row_values), safety_val))

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


def write_intermediate(headers, data_rows, raw_headers):
    wb = build_intermediate_wb(headers, data_rows, raw_headers)
    wb.save(INTERMEDIATE_FILE)
    wb.close()
    print(f"Saved {INTERMEDIATE_FILE}")


def write_final(headers, data_rows, raw_headers):
    wb = build_final_wb(headers, data_rows, raw_headers)
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
