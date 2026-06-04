import io
import datetime
import re
import time
import traceback

import streamlit as st
import openpyxl
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter
import pandas as pd

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

DARK_THEME_CSS = """
<style>
.stApp, .main {
    background: linear-gradient(135deg, #0a0a0f 0%, #1a1a2e 50%, #0f0f1a 100%) !important;
    color: #ffffff !important;
}

h1 {
    color: #00f5ff !important;
    text-shadow: 0 0 20px rgba(0,245,255,0.6), 0 0 40px rgba(0,245,255,0.3);
    font-weight: 900 !important;
}
h2, h3 {
    color: #ffbe0b !important;
    text-shadow: 0 0 15px rgba(255,190,11,0.5);
}

hr {
    border: none;
    height: 2px;
    background: linear-gradient(90deg, transparent, #00f5ff, transparent) !important;
    margin: 1.5rem 0 !important;
}

[data-testid="stDataFrame"] {
    background: rgba(10,10,20,0.95) !important;
    border-radius: 12px !important;
    border: 1px solid rgba(0,245,255,0.3) !important;
}
[data-testid="stDataFrame"] th {
    background: linear-gradient(135deg, #1a1a2e 0%, #2d2d44 100%) !important;
    color: #00f5ff !important;
    font-weight: 700 !important;
    text-shadow: 0 0 10px rgba(0,245,255,0.5);
    border-bottom: 2px solid #ff006e !important;
}
[data-testid="stDataFrame"] td {
    color: #e8e8e8 !important;
    border-color: rgba(0,245,255,0.15) !important;
}
[data-testid="stDataFrame"] tr:hover {
    background: rgba(255,0,110,0.2) !important;
}

.stButton > button {
    background: linear-gradient(135deg, #ff006e 0%, #c4116a 100%) !important;
    color: #fff !important;
    border: none !important;
    border-radius: 12px !important;
    font-weight: 700 !important;
    padding: 0.8rem 2.5rem !important;
    box-shadow: 0 0 25px rgba(255,0,110,0.5) !important;
    transition: all 0.3s ease !important;
}
.stButton > button:hover {
    transform: scale(1.05);
    box-shadow: 0 0 40px rgba(255,0,110,0.8) !important;
}

.stDownloadButton > button {
    background: linear-gradient(135deg, #39ff14 0%, #00c853 100%) !important;
    color: #0a0a0f !important;
    font-weight: 700 !important;
}

.streamlit-expanderHeader {
    background: linear-gradient(90deg, rgba(0,245,255,0.15) 0%, rgba(255,0,110,0.1) 100%) !important;
    border-radius: 12px !important;
    border: 1px solid rgba(0,245,255,0.3) !important;
    color: #00f5ff !important;
    font-weight: 700 !important;
}

[data-testid="stFileUploadDropzone"] {
    background: rgba(26,26,46,0.9) !important;
    border: 2px dashed rgba(0,245,255,0.5) !important;
    border-radius: 16px !important;
}

.stSuccess {
    background: rgba(57,255,20,0.15) !important;
    border: 1px solid #39ff14 !important;
    color: #39ff14 !important;
}
.stError {
    background: rgba(255,0,110,0.15) !important;
    border: 1px solid #ff006e !important;
    color: #ff006e !important;
}

.stTabs [data-baseweb="tab"] {
    background: rgba(26,26,46,0.8) !important;
    border-radius: 8px 8px 0 0 !important;
    color: #e8e8e8 !important;
}
.stTabs [aria-selected="true"] {
    background: rgba(0,245,255,0.2) !important;
    color: #00f5ff !important;
    font-weight: 700 !important;
}

.metric-grid {
    display: grid;
    grid-template-columns: repeat(4, 1fr);
    gap: 16px;
    margin: 1rem 0 1.5rem 0;
}

.metric-card {
    background: linear-gradient(135deg, #13151c 0%, #1a1d28 100%);
    border-radius: 14px;
    border: 1px solid rgba(255,255,255,0.06);
    padding: 1.25rem 1.5rem;
    position: relative;
    overflow: hidden;
    transition: all 0.3s ease;
    min-height: 130px;
    display: flex;
    align-items: stretch;
}

.metric-card:hover {
    transform: translateY(-3px);
    box-shadow: 0 12px 40px rgba(0,0,0,0.5);
    border-color: rgba(255,255,255,0.15);
}

.metric-card__bar {
    position: absolute;
    left: 0;
    top: 8px;
    bottom: 8px;
    width: 4px;
    border-radius: 0 4px 4px 0;
}

.metric-card__content {
    display: flex;
    flex-direction: column;
    justify-content: center;
    gap: 6px;
    margin-left: 16px;
    width: 100%;
}

.metric-card__title {
    font-size: 1.17rem;
    font-weight: 700;
    color: #b8e0ff;
    letter-spacing: 0.5px;
    text-shadow: 0 0 8px rgba(184,224,255,0.3);
}

.metric-card__value {
    font-size: 2.1rem;
    font-weight: 900;
    color: #ffffff;
    line-height: 1.1;
    letter-spacing: -0.5px;
    text-shadow: 0 0 14px rgba(255,255,255,0.35);
}

.metric-card__description {
    font-size: 1.05rem;
    color: #7ac0ff;
    margin-top: 1px;
    line-height: 1.3;
    text-shadow: 0 0 6px rgba(122,192,255,0.25);
}

@media (max-width: 1100px) {
    .metric-grid {
        grid-template-columns: repeat(3, 1fr);
    }
}

@media (max-width: 768px) {
    .metric-grid {
        grid-template-columns: repeat(2, 1fr);
    }
}

@media (max-width: 480px) {
    .metric-grid {
        grid-template-columns: 1fr;
    }
}
</style>
"""


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


def get_new_launch_flag(launch_date_val, reference_date=None):
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


def process_uploaded_file(uploaded_bytes, progress_callback=None):
    start_time = time.time()
    wb_in = openpyxl.load_workbook(io.BytesIO(uploaded_bytes), data_only=True)
    ws = wb_in["Export"]
    raw_headers = [cell.value for cell in ws[1]]

    item_col_idx = raw_headers.index("Item")
    buyer_col_idx = raw_headers.index("Buyer")
    ratio_col_idx = raw_headers.index("Safety VS Monthly Sold Qty")
    item_class_idx = raw_headers.index("Item Class")
    sku_desc_col_idx = raw_headers.index("SKU Desc")

    total_rows = ws.max_row - 1
    data_rows = []
    for r, row in enumerate(ws.iter_rows(min_row=2, values_only=True), 2):
        col_a_val = row[0]
        if col_a_val == "Total" or (isinstance(col_a_val, str) and col_a_val.startswith("Applied filters:")):
            if progress_callback:
                progress_callback(r - 1, total_rows, time.time() - start_time, skipped=True)
            continue
        item_val = row[item_col_idx]
        buyer_val = row[buyer_col_idx]
        if buyer_val is None or item_val is None:
            if progress_callback:
                progress_callback(r - 1, total_rows, time.time() - start_time, skipped=True)
            continue
        sku_desc_val = str(row[sku_desc_col_idx]) if row[sku_desc_col_idx] else ""
        if "(STAFF)" in sku_desc_val or "(MAC)" in sku_desc_val:
            if progress_callback:
                progress_callback(r - 1, total_rows, time.time() - start_time, skipped=True)
            continue
        ratio_val = row[ratio_col_idx]
        item_class_val = row[item_class_idx]
        sasa_abc = str(item_class_val)[0] if item_class_val else ""
        safety_val = get_safety_checking(ratio_val, sasa_abc)
        data_rows.append((list(row), safety_val))
        if progress_callback:
            progress_callback(r - 1, total_rows, time.time() - start_time)

    wb_in.close()

    new_headers = list(raw_headers)
    new_headers.insert(raw_headers.index("Item Class") + 1, "SASA ABC")
    new_headers.insert(new_headers.index("Safety VS Monthly Sold Qty") + 1, "Safety Checking")

    elapsed = time.time() - start_time
    return raw_headers, new_headers, data_rows, elapsed


def build_row(raw_values, safety_val, raw_headers):
    row = list(raw_values)
    item_class_pos = raw_headers.index("Item Class")
    item_class_val = raw_values[item_class_pos]
    sasa_abc = str(item_class_val)[0] if item_class_val else ""
    row.insert(item_class_pos + 1, sasa_abc)

    safety_pos = raw_headers.index("Safety VS Monthly Sold Qty")
    offset = 1 if item_class_pos < safety_pos else 0
    row.insert(safety_pos + 1 + offset, safety_val)
    return row


def build_intermediate_wb(new_headers, data_rows, raw_headers):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Export"

    for c, h in enumerate(new_headers, 1):
        ws.cell(row=1, column=c, value=h)

    for r_idx, (raw_values, safety_val) in enumerate(data_rows, 2):
        row = build_row(raw_values, safety_val, raw_headers)
        for c, val in enumerate(row, 1):
            ws.cell(row=r_idx, column=c, value=val)

    return wb


def build_final_wb(new_headers, data_rows, raw_headers):
    buyer_col_idx = raw_headers.index("Buyer")

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

    ws_sku = wb.active
    ws_sku.title = "SKU"

    for c, h in enumerate(new_headers, 1):
        ws_sku.cell(row=1, column=c, value=h)

    for r_idx, (raw_values, safety_val) in enumerate(sorted_data, 2):
        row = build_row(raw_values, safety_val, raw_headers)
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

    return wb


def wb_to_bytes(wb):
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


def filter_ndrf_data(ws, progress_callback=None):
    raw_headers = [cell.value for cell in ws[1]]
    buyer_col_idx = raw_headers.index("Buyer")
    supply_source_col_idx = raw_headers.index("Supply Source")
    launch_date_col_idx = raw_headers.index("Launch Date")
    total_stock_col_idx = raw_headers.index("Total Stock Qty (CM)")
    adj_safety_stock_col_idx = raw_headers.index("Adj. Safety Stock")
    sku_desc_col_idx = raw_headers.index("SKU Desc")

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


def compute_nd_checking(no_of_rf_shop, total_shops):
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

def process_ndrf_mode(file_bytes, total_shops, sub_mode="顯示所有 SKU"):
    progress_bar = st.progress(0)
    status_text = st.empty()
    elapsed_text = st.empty()

    wb_in = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)
    ws = wb_in["Export"]

    total_rows = ws.max_row - 1
    status_text.markdown(f"🔄 **Processing ND data...** `0` / `{total_rows:,}` lines (0.0%)")
    elapsed_text.markdown(f"⏱️  **Elapsed:** `00:00`")

    def update_progress(current, total):
        pct = current / total if total > 0 else 0
        progress_bar.progress(min(pct, 1.0))
        status_text.markdown(
            f"🔄 **Processing ND data...** `{current:,}` / `{total:,}` lines "
            f"({pct*100:.1f}%)"
        )

    raw_headers, data_rows = filter_ndrf_data(ws, progress_callback=update_progress)
    wb_in.close()

    progress_bar.progress(1.0)
    status_text.markdown(f"🔄 **Computing ND ratios...** `{len(data_rows):,}` rows filtered")

    buyer_col_idx = raw_headers.index("Buyer")
    launch_date_col_idx = raw_headers.index("Launch Date")
    today = datetime.date.today()

    total_sku = len(data_rows)

    gt_80 = 0
    btw_50_80 = 0
    ok_nd = 0
    enriched_data = []
    buyer_counts = {}
    for i, raw_values in enumerate(data_rows):
        nd_ratios, nd_checking = compute_nd_checking(
            raw_values[raw_headers.index("No. of RF Shop")], total_shops
        )
        new_launch_flag = get_new_launch_flag(raw_values[launch_date_col_idx], today)
        if nd_checking == "ND Ratio > 80%":
            gt_80 += 1
        elif nd_checking == "ND Ratio 70%-80%":
            btw_50_80 += 1
        else:
            ok_nd += 1
        enriched_data.append((raw_values, nd_ratios, nd_checking, new_launch_flag))

        buyer = raw_values[buyer_col_idx]
        if buyer not in buyer_counts:
            buyer_counts[buyer] = {"ND Ratio > 80%": 0, "ND Ratio 70%-80%": 0, "Total": 0}
        if nd_checking == "ND Ratio > 80%":
            buyer_counts[buyer]["ND Ratio > 80%"] += 1
        elif nd_checking == "ND Ratio 70%-80%":
            buyer_counts[buyer]["ND Ratio 70%-80%"] += 1
        buyer_counts[buyer]["Total"] += 1

        pct = (i + 1) / total_sku if total_sku > 0 else 0
        progress_bar.progress(min(pct, 1.0))
        status_text.markdown(
            f"🔄 **Computing ND ratios...** `{i + 1:,}` / `{total_sku:,}` rows "
            f"({pct*100:.1f}%)"
        )

    progress_bar.progress(1.0)
    status_text.markdown(f"✅ **Processing complete!** `{total_sku:,}` SKUs analyzed")
    progress_bar.empty()

    if sub_mode == "只顯示 New SKU":
        enriched_data = [(rv, nr, nc, nlf) for rv, nr, nc, nlf in enriched_data if nlf == "New SKU"]
        gt_80 = sum(1 for _, _, nc, _ in enriched_data if nc == "ND Ratio > 80%")
        btw_50_80 = sum(1 for _, _, nc, _ in enriched_data if nc == "ND Ratio 70%-80%")
        ok_nd = sum(1 for _, _, nc, _ in enriched_data if nc is None)
        total_sku = len(enriched_data)
        buyer_counts = {}
        for rv, _, nc, _ in enriched_data:
            buyer = rv[buyer_col_idx]
            if buyer not in buyer_counts:
                buyer_counts[buyer] = {"ND Ratio > 80%": 0, "ND Ratio 70%-80%": 0, "Total": 0}
            if nc == "ND Ratio > 80%":
                buyer_counts[buyer]["ND Ratio > 80%"] += 1
            elif nc == "ND Ratio 70%-80%":
                buyer_counts[buyer]["ND Ratio 70%-80%"] += 1
            buyer_counts[buyer]["Total"] += 1

    METRIC_ACCENT_ND = {
        "Total Filtered SKU": "#00f5ff",
        "ND Ratio > 80%": "#ff006e",
        "ND Ratio 70%-80%": "#ff7b00",
        "ND Ratio < 70%": "#39ff14",
    }
    METRIC_DESC_ND = {
        "Total Filtered SKU": "SKUs after first-layer filtering",
        "ND Ratio > 80%": "ND ratio exceeds 80% threshold",
        "ND Ratio 70%-80%": "ND ratio between 70%-80%",
        "ND Ratio < 70%": "ND ratio within acceptable range",
    }

    st.subheader("Key Metrics")
    metric_items = [
        ("Total Filtered SKU", total_sku),
        ("ND Ratio > 80%", gt_80),
        ("ND Ratio 70%-80%", btw_50_80),
        ("ND Ratio < 70%", ok_nd),
    ]
    cards = []
    for key, val in metric_items:
        accent = METRIC_ACCENT_ND.get(key, "#00f5ff")
        desc = METRIC_DESC_ND.get(key, "")
        cards.append(
            f'<div class="metric-card">'
            f'<div class="metric-card__bar" style="background:{accent};box-shadow:0 0 10px {accent}66;"></div>'
            f'<div class="metric-card__content">'
            f'<div class="metric-card__title">{key}</div>'
            f'<div class="metric-card__value">{val:,}</div>'
            f'<div class="metric-card__description">{desc}</div>'
            f'</div>'
            f'</div>'
        )
    st.markdown(
        f'<div class="metric-grid">{"".join(cards)}</div>',
        unsafe_allow_html=True
    )

    st.markdown("---")

    sorted_buyers = sorted(buyer_counts.keys(), key=extract_buyer_number)

    st.subheader("Summary by Buyer")
    st.caption("Filtered: Warehouse only, Total Stock Qty > 100 units. 'Total' = all filtered SKUs (incl. ND Ratio < 70%).")
    summary_rows = []
    for buyer in sorted_buyers:
        bc = buyer_counts[buyer]
        total = bc["Total"]
        gt80 = bc["ND Ratio > 80%"]
        btw = bc["ND Ratio 70%-80%"]
        summary_rows.append({
            "Buyer": buyer,
            "ND Ratio > 80%": gt80,
            "ND Ratio 70%-80%": btw,
            "ND Ratio > 80% %": f"{gt80/total*100:.1f}%" if total > 0 else "0%",
            "ND Ratio 70%-80% %": f"{btw/total*100:.1f}%" if total > 0 else "0%",
            "Total": total,
        })

    grand_gt80 = sum(bc["ND Ratio > 80%"] for bc in buyer_counts.values())
    grand_btw = sum(bc["ND Ratio 70%-80%"] for bc in buyer_counts.values())
    grand_total = sum(bc["Total"] for bc in buyer_counts.values())
    summary_rows.append({
        "Buyer": "Grand Total",
        "ND Ratio > 80%": grand_gt80,
        "ND Ratio 70%-80%": grand_btw,
        "ND Ratio > 80% %": f"{grand_gt80/grand_total*100:.1f}%" if grand_total > 0 else "0%",
        "ND Ratio 70%-80% %": f"{grand_btw/grand_total*100:.1f}%" if grand_total > 0 else "0%",
        "Total": grand_total,
    })

    summary_df = pd.DataFrame(summary_rows)
    st.dataframe(summary_df, use_container_width=True, hide_index=True)

    st.markdown("---")

    sku_col_idx = raw_headers.index("SKU")
    sku_desc_col_idx = raw_headers.index("SKU Desc")
    brand_col_idx = raw_headers.index("Brand")
    rf_shop_idx = raw_headers.index("No. of RF Shop")
    nd_shop_stock_idx = raw_headers.index("No. of ND Shop with Stock")

    gt80_data = []
    btw50_80_data = []
    for raw_values, nd_ratios, nd_checking, new_launch_flag in enriched_data:
        if nd_checking == "ND Ratio > 80%":
            gt80_data.append({
                "Buyer": raw_values[buyer_col_idx],
                "SKU": raw_values[sku_col_idx],
                "SKU Desc": raw_values[sku_desc_col_idx],
                "Brand": raw_values[brand_col_idx],
                "No. of RF Shop": raw_values[rf_shop_idx],
                "No. of ND Shop with Stock": raw_values[nd_shop_stock_idx],
                "ND ratios": f"{nd_ratios*100:.1f}%",
                "ND Checking": nd_checking,
                "New Launch Flag (L3M)": new_launch_flag,
            })
        elif nd_checking == "ND Ratio 70%-80%":
            btw50_80_data.append({
                "Buyer": raw_values[buyer_col_idx],
                "SKU": raw_values[sku_col_idx],
                "SKU Desc": raw_values[sku_desc_col_idx],
                "Brand": raw_values[brand_col_idx],
                "No. of RF Shop": raw_values[rf_shop_idx],
                "No. of ND Shop with Stock": raw_values[nd_shop_stock_idx],
                "ND ratios": f"{nd_ratios*100:.1f}%",
                "ND Checking": nd_checking,
                "New Launch Flag (L3M)": new_launch_flag,
            })

    if gt80_data:
        with st.expander(f"ND Ratio > 80% Problem SKUs ({len(gt80_data)} SKUs)"):
            gt80_df = pd.DataFrame(gt80_data)
            st.dataframe(gt80_df, use_container_width=True, hide_index=True)

    if btw50_80_data:
        with st.expander(f"ND Ratio 70%-80% Problem SKUs ({len(btw50_80_data)} SKUs)"):
            btw50_80_df = pd.DataFrame(btw50_80_data)
            st.dataframe(btw50_80_df, use_container_width=True, hide_index=True)

    st.markdown("---")

    if st.button("Generate Output Files", type="primary"):
        gen_progress = st.progress(0)
        gen_status = st.empty()
        gen_start = time.time()

        gen_status.markdown("📝 Building workbook (RAW + Summary sheets)...")
        gen_progress.progress(0.3)

        COLUMNS_TO_DELETE = {
            "WH Stock Value (CM)", "Shop Stock Value (CM)", "Total Stock Value (CM)",
            "P90D Daily COGS", "GP$ (P90D)", "GP$ (M-2)", "GP$ (M-1)", "GP$ (CM - MTD)"
        }
        COLUMNS_TO_MOVE = ["D001 Net Stock Qty (CM)", "Shop Stock Qty (CM)", "Total Stock Qty (CM)"]

        intermediate_headers = list(raw_headers)
        launch_date_idx = raw_headers.index("Launch Date")
        intermediate_headers.insert(launch_date_idx + 1, "New Launch Flag (L3M)")
        nd_shop_stock_idx = raw_headers.index("No. of ND Shop with Stock")
        adjusted_nd_idx = nd_shop_stock_idx + (1 if nd_shop_stock_idx > launch_date_idx else 0)
        intermediate_headers.insert(adjusted_nd_idx + 1, "ND ratios")
        intermediate_headers.insert(adjusted_nd_idx + 2, "ND Checking")

        final_headers = [h for h in intermediate_headers if h not in COLUMNS_TO_DELETE]
        for col in COLUMNS_TO_MOVE:
            final_headers.remove(col)
        nd_checking_pos = final_headers.index("ND Checking")
        for i, col in enumerate(COLUMNS_TO_MOVE):
            final_headers.insert(nd_checking_pos + 1 + i, col)

        wb = openpyxl.Workbook()
        ws_raw = wb.active
        ws_raw.title = "RAW"

        for c, h in enumerate(final_headers, 1):
            ws_raw.cell(row=1, column=c, value=h)

        nd_ratios_col = final_headers.index("ND ratios") + 1
        for r_idx, (raw_values, nd_ratios, nd_checking, new_launch_flag) in enumerate(enriched_data, 2):
            intermediate_row = list(raw_values)
            intermediate_row.insert(launch_date_idx + 1, new_launch_flag)
            adjusted_nd_idx = nd_shop_stock_idx + (1 if nd_shop_stock_idx > launch_date_idx else 0)
            intermediate_row.insert(adjusted_nd_idx + 1, nd_ratios)
            intermediate_row.insert(adjusted_nd_idx + 2, nd_checking)
            row_dict = dict(zip(intermediate_headers, intermediate_row))
            final_row = [row_dict[h] for h in final_headers]
            for c, val in enumerate(final_row, 1):
                cell = ws_raw.cell(row=r_idx, column=c, value=val)
                if c == nd_ratios_col:
                    cell.number_format = "0.0%"

        sorted_buyers = sorted(buyer_counts.keys(), key=extract_buyer_number)

        ws_sum = wb.create_sheet("Summary", 0)
        ws_sum.cell(row=1, column=1, value="Note: Only Warehouse supply source, Total Stock Qty > 100 units")
        ws_sum.cell(row=2, column=1, value="'Total' = all filtered SKUs (including ND Ratio < 70%)")

        header_row = 4
        ws_sum.cell(row=header_row, column=1, value="Buyer")
        ws_sum.cell(row=header_row, column=2, value="ND Ratio > 80%")
        ws_sum.cell(row=header_row, column=3, value="ND Ratio 70%-80%")
        ws_sum.cell(row=header_row, column=4, value="ND Ratio > 80% %")
        ws_sum.cell(row=header_row, column=5, value="ND Ratio 70%-80% %")
        ws_sum.cell(row=header_row, column=6, value="Total Filtered SKU")

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

        nd_problem_headers = ["Buyer", "SKU", "SKU Desc", "Brand", "No. of RF Shop",
                              "No. of ND Shop with Stock", "ND ratios", "ND Checking",
                              "New Launch Flag (L3M)"]
        nd_field_indices = [
            raw_headers.index("Buyer"), raw_headers.index("SKU"), raw_headers.index("SKU Desc"),
            raw_headers.index("Brand"), raw_headers.index("No. of RF Shop"),
            raw_headers.index("No. of ND Shop with Stock"),
        ]

        for sheet_title, nd_condition, rows_list in [
            ("ND Ratio > 80%", "ND Ratio > 80%", []),
            ("ND Ratio 70%-80%", "ND Ratio 70%-80%", []),
        ]:
            ws = wb.create_sheet(sheet_title, 1 if sheet_title == "ND Ratio > 80%" else 2)
            for c, h in enumerate(nd_problem_headers, 1):
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

        gen_progress.progress(0.8)
        gen_status.markdown("💾 Converting to bytes...")
        wb_bytes = wb_to_bytes(wb)

        gen_progress.progress(1.0)
        today = datetime.date.today()
        date_str = today.strftime("%d%b%Y")
        filename = f"Check NDRF_To Buyer_{date_str}.xlsx"
        gen_m, gen_s = divmod(int(time.time() - gen_start), 60)
        gen_status.markdown(f"✅ **File generated!** (`{gen_m:02d}:{gen_s:02d}`)")

        st.download_button(
            label=f"Download {filename}",
            data=wb_bytes,
            file_name=filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )


st.set_page_config(page_title="Safety Stock Review", page_icon="\U0001f4ca", layout="wide")
st.markdown(DARK_THEME_CSS, unsafe_allow_html=True)

st.title("A Item Safety Stock Review vs Monthly Sold Qty")
st.markdown("---")

mode = st.radio("Mode", ["Review Safety", "ND Ratio Review"], index=0, horizontal=True)

uploaded_file = st.file_uploader(
    "\U0001f4c1  Upload By SKU_RAW.xlsx",
    type=["xlsx"],
    help="Select the RAW export file with 'Export' sheet"
)

if mode == "ND Ratio Review":
    total_shops = st.number_input(
        "Total Shops",
        min_value=1, max_value=500, value=86, step=1,
        help="Total number of shops used for ND ratio calculation"
    )
    nd_sub_mode = st.radio(
        "ND Ratio 顯示模式",
        ["顯示所有 SKU", "只顯示 New SKU"],
        index=0,
        horizontal=True
    )

if uploaded_file is not None:
    try:
        file_bytes = uploaded_file.read()

        if mode == "ND Ratio Review":
            try:
                process_ndrf_mode(file_bytes, total_shops, nd_sub_mode)
            except Exception as e:
                st.error(f"ND Mode Error: {e}")
                st.exception(e)

        else:
            progress_bar = st.progress(0)
            status_text = st.empty()
            elapsed_text = st.empty()
            detail_text = st.empty()
            processing_start = time.time()
            last_update = [0.0]
    
            def update_progress(current, total, elapsed, skipped=False):
                if elapsed - last_update[0] < 0.3 and current != total:
                    return
                last_update[0] = elapsed
                pct = current / total if total > 0 else 0
                progress_bar.progress(min(pct, 1.0))
                mins, secs = divmod(int(elapsed), 60)
                status_text.markdown(
                    f"🔄 **Processing...** `{current:,}` / `{total:,}` lines  "
                    f"({pct*100:.1f}%)"
                )
                elapsed_text.markdown(
                    f"⏱️ **Elapsed:** `{mins:02d}:{secs:02d}`  |  "
                    f"**Valid rows:** `{current}`"
                )
                if skipped:
                    detail_text.caption(f"⏭️ Skipped row {current:,}")
    
            raw_headers, new_headers, data_rows, file_elapsed = process_uploaded_file(
                file_bytes, progress_callback=update_progress
            )
    
            progress_bar.progress(1.0)
            total_elapsed = time.time() - processing_start
            mins, secs = divmod(int(total_elapsed), 60)
            status_text.markdown(
                f"✅ **Read complete!** `{len(data_rows):,}` valid rows processed"
            )
            elapsed_text.markdown(
                f"⏱️ **Total read time:** `{mins:02d}:{secs:02d}`"
            )
            detail_text.empty()
    
            categories_display = sorted(set(sv for _, sv in data_rows), key=lambda x: SAFETY_ORDER.get(x, 99))
            safety_counts = {cat: 0 for cat in categories_display}
            for _, sv in data_rows:
                safety_counts[sv] += 1
    
            st.subheader("Key Metrics")
    
            METRIC_LABELS = {
                "Total SKU": "Total SKU",
                "Aitem<25%": "A Item < 25%",
                "New item<15%": "New < 15%",
                "B/C/D item<15%": "B/C/D < 15%",
                "Aitem>75%": "A Item > 75%",
                "New item>100%": "New > 100%",
                "B/C/D item>100%": "B/C/D > 100%",
                "OK": "OK",
                "Check": "Check",
            }
            METRIC_DESC = {
                "Total SKU": "Total SKUs analyzed",
                "Aitem<25%": "Safety stock below 25% threshold",
                "New item<15%": "New items below 15% threshold",
                "B/C/D item<15%": "B/C/D items below 15% threshold",
                "Aitem>75%": "Safety stock above 75% threshold",
                "New item>100%": "New items exceed 100% of sold qty",
                "B/C/D item>100%": "B/C/D items exceed 100% of sold qty",
                "OK": "Within normal safety stock range",
                "Check": "Requires manual review",
            }
            METRIC_ACCENT = {
                "Total SKU": "#00f5ff",
                "Aitem<25%": "#ff006e",
                "New item<15%": "#ff7b00",
                "B/C/D item<15%": "#ffbe0b",
                "Aitem>75%": "#b100e8",
                "New item>100%": "#ff00aa",
                "B/C/D item>100%": "#ff5e7d",
                "OK": "#39ff14",
                "Check": "#ffd60a",
            }
    
            metric_items = [("Total SKU", len(data_rows))] + [
                (cat, safety_counts[cat]) for cat in categories_display
            ]
            cards = []
            for key, val in metric_items:
                label = METRIC_LABELS.get(key, key)
                desc = METRIC_DESC.get(key, "")
                accent = METRIC_ACCENT.get(key, "#00f5ff")
                cards.append(
                    f'<div class="metric-card">'
                    f'<div class="metric-card__bar" style="background:{accent};box-shadow:0 0 10px {accent}66;"></div>'
                    f'<div class="metric-card__content">'
                    f'<div class="metric-card__title">{label}</div>'
                    f'<div class="metric-card__value">{val:,}</div>'
                    f'<div class="metric-card__description">{desc}</div>'
                    f'</div>'
                    f'</div>'
                )
            st.markdown(
                f'<div class="metric-grid">{"".join(cards)}</div>',
                unsafe_allow_html=True
            )
    
            st.markdown("---")
    
            buyer_col_idx = raw_headers.index("Buyer")
            categories = sorted(set(sv for _, sv in data_rows), key=lambda x: SAFETY_ORDER.get(x, 99))
    
            buyer_data = {}
            for raw_values, safety_val in data_rows:
                buyer = raw_values[buyer_col_idx]
                if buyer not in buyer_data:
                    buyer_data[buyer] = {cat: 0 for cat in categories}
                buyer_data[buyer][safety_val] += 1
    
            sorted_buyers = sorted(buyer_data.keys(), key=extract_buyer_number)
    
            summary_data = []
            for buyer in sorted_buyers:
                total = sum(buyer_data[buyer][cat] for cat in categories)
                row = {"Buyer": buyer}
                for cat in categories:
                    row[cat] = buyer_data[buyer][cat]
                    row[f"{cat} %"] = f"{buyer_data[buyer][cat]/total*100:.1f}%" if total > 0 else "0%"
                row["Total"] = total
                summary_data.append(row)
    
            grand_total = sum(sum(buyer_data[b][cat] for cat in categories) for b in sorted_buyers)
            total_row = {"Buyer": "Grand Total"}
            for cat in categories:
                cat_total = sum(buyer_data[b][cat] for b in sorted_buyers)
                total_row[cat] = cat_total
                total_row[f"{cat} %"] = f"{cat_total/grand_total*100:.1f}%" if grand_total > 0 else "0%"
            total_row["Total"] = grand_total
            summary_data.append(total_row)

            st.subheader("Summary by Buyer")
            summary_df = pd.DataFrame(summary_data)
            st.dataframe(summary_df, use_container_width=True, hide_index=True)
    
            st.markdown("---")
    
            st.subheader("Aitem<25% Problem SKUs & Buyer Info")
    
            supply_source_col_idx = raw_headers.index("Supply Source")
            shop_stock_col_idx = raw_headers.index("Shop Stock Qty (CM)")
            safety_stock_col_idx = raw_headers.index("Adj. Safety Stock")
            shop_monthly_sold_col_idx = raw_headers.index("Shop Monthly Sold Qty (P90D)")
            d001_net_stock_col_idx = raw_headers.index("D001 Net Stock Qty (CM)")
            sku_col_idx = raw_headers.index("SKU")
            sku_desc_col_idx = raw_headers.index("SKU Desc")
            brand_col_idx = raw_headers.index("Brand")
            ratio_col_idx = raw_headers.index("Safety VS Monthly Sold Qty")
    
            problem_data = []
            for raw_values, safety_val in data_rows:
                if safety_val == "Aitem<25%":
                    buyer = raw_values[buyer_col_idx]
                    sku = raw_values[sku_col_idx]
                    sku_desc = raw_values[sku_desc_col_idx]
                    brand = raw_values[brand_col_idx]
                    supply_source = str(raw_values[supply_source_col_idx]) if raw_values[supply_source_col_idx] else ""
                    safety_stock = raw_values[safety_stock_col_idx]
                    shop_monthly_sold = raw_values[shop_monthly_sold_col_idx]
                    ratio = raw_values[ratio_col_idx]
                    d001_net_stock = raw_values[d001_net_stock_col_idx]
                    shop_stock = raw_values[shop_stock_col_idx]
                    problem_data.append({
                        "Buyer": buyer,
                        "SKU": sku,
                        "SKU Desc": sku_desc,
                        "Brand": brand,
                        "Supply Source": supply_source,
                        "Safety Stock Qty": safety_stock,
                        "Shop Monthly Sold Qty (P90D)": int(shop_monthly_sold) if shop_monthly_sold is not None else None,
                        "Safety/Monthly %": f"{ratio*100:.1f}%" if ratio else "N/A",
                        "D001 Net Stock Qty (CM)": d001_net_stock,
                        "Shop Stock Qty (CM)": shop_stock
                    })
    
            if problem_data:
                problem_df = pd.DataFrame(problem_data)
    
                tab1, tab2 = st.tabs(["All Problem SKUs", "By Buyer"])
                with tab1:
                    st.dataframe(problem_df, use_container_width=True, hide_index=True)
                    st.info(f"**{len(problem_df)}** Aitem<25% problem SKUs found")
    
                with tab2:
                    buyer_summary = problem_df.groupby("Buyer").agg(
                        Problem_SKU_Count=("SKU", "count"),
                        Brands=("Brand", lambda x: ", ".join(sorted(set(x))))
                    ).reset_index()
                    buyer_summary = buyer_summary.sort_values("Problem_SKU_Count", ascending=False)
                    st.dataframe(buyer_summary, use_container_width=True, hide_index=True)
            else:
                st.success("No Aitem<25% problem SKUs found!")
    
            # ── New Section: Aitem>200% & D001<20pcs ──
            st.subheader("Aitem>200% & D001<20pcs Problem SKUs & Buyer Info")
            st.caption("Supply Source = Warehouse only (Vendor items do not go through D001)")
    
            ratio_col_idx_200 = raw_headers.index("Safety VS Monthly Sold Qty")
            supply_source_col_idx = raw_headers.index("Supply Source")
            d001_net_stock_col_idx = raw_headers.index("D001 Net Stock Qty (CM)")
            shop_stock_col_idx = raw_headers.index("Shop Stock Qty (CM)")
            safety_stock_col_idx = raw_headers.index("Adj. Safety Stock")
            shop_monthly_sold_col_idx = raw_headers.index("Shop Monthly Sold Qty (P90D)")
            sku_col_idx_200 = raw_headers.index("SKU")
            sku_desc_col_idx_200 = raw_headers.index("SKU Desc")
            brand_col_idx_200 = raw_headers.index("Brand")
    
            problem_data_200 = []
            for raw_values, safety_val in data_rows:
                ratio = raw_values[ratio_col_idx_200]
                item_class_val = raw_values[raw_headers.index("Item Class")]
                sasa_abc = str(item_class_val)[0] if item_class_val else ""
                supply_source = str(raw_values[supply_source_col_idx]) if raw_values[supply_source_col_idx] else ""
                d001_net_stock = raw_values[d001_net_stock_col_idx]
    
                # A item + ratio > 200% + Warehouse only + D001 < 20
                if (sasa_abc == "A"
                        and ratio is not None and ratio > 2.0
                        and supply_source == "Warehouse"
                        and d001_net_stock is not None and d001_net_stock < 20):
                    buyer = raw_values[raw_headers.index("Buyer")]
                    sku = raw_values[sku_col_idx_200]
                    sku_desc = raw_values[sku_desc_col_idx_200]
                    brand = raw_values[brand_col_idx_200]
                    shop_stock = raw_values[shop_stock_col_idx]
                    safety_stock = raw_values[safety_stock_col_idx]
                    shop_monthly_sold = raw_values[shop_monthly_sold_col_idx]
                    problem_data_200.append({
                        "Buyer": buyer,
                        "SKU": sku,
                        "SKU Desc": sku_desc,
                        "Brand": brand,
                        "Supply Source": supply_source,
                        "Safety Stock Qty": safety_stock,
                        "Shop Monthly Sold Qty (P90D)": int(shop_monthly_sold) if shop_monthly_sold is not None else None,
                        "Safety/Monthly %": f"{ratio*100:.1f}%" if ratio else "N/A",
                        "D001 Net Stock Qty (CM)": d001_net_stock,
                        "Shop Stock Qty (CM)": shop_stock
                    })
    
            if problem_data_200:
                problem_df_200 = pd.DataFrame(problem_data_200)
    
                tab_200_1, tab_200_2 = st.tabs(["All Problem SKUs", "By Buyer"])
                with tab_200_1:
                    st.dataframe(problem_df_200, use_container_width=True, hide_index=True)
                    st.info(f"**{len(problem_df_200)}** Aitem>200% & D001<20pcs problem SKUs found")

                with tab_200_2:
                    buyer_summary_200 = problem_df_200.groupby("Buyer").agg(
                        Problem_SKU_Count=("SKU", "count"),
                        Brands=("Brand", lambda x: ", ".join(sorted(set(x))))
                    ).reset_index()
                    buyer_summary_200 = buyer_summary_200.sort_values("Problem_SKU_Count", ascending=False)
                    st.dataframe(buyer_summary_200, use_container_width=True, hide_index=True)
            else:
                st.success("No Aitem>200% & D001<20pcs problem SKUs found!")
    
            st.markdown("---")
    
            if st.button("Generate Output Files", type="primary"):
                gen_progress = st.progress(0)
                gen_status = st.empty()
                gen_elapsed = st.empty()
                gen_start = time.time()
    
                gen_status.markdown("📝 Building intermediate workbook...")
                gen_progress.progress(0.2)
                wb_intermediate = build_intermediate_wb(new_headers, data_rows, raw_headers)
    
                gen_status.markdown("📊 Building final workbook with summary...")
                gen_progress.progress(0.5)
                wb_final = build_final_wb(new_headers, data_rows, raw_headers)
    
                gen_status.markdown("💾 Converting to bytes...")
                gen_progress.progress(0.8)
    
                intermediate_bytes = wb_to_bytes(wb_intermediate)
    
                today = datetime.date.today()
                date_str = today.strftime("%d%b%Y")
                final_filename = f"Check Safety A Item by SKU_To Buyer_{date_str}.xlsx"
    
                final_bytes = wb_to_bytes(wb_final)
    
                gen_progress.progress(1.0)
                gen_m, gen_s = divmod(int(time.time() - gen_start), 60)
                gen_status.markdown(f"✅ **Files generated!** (`{gen_m:02d}:{gen_s:02d}`)")
                gen_elapsed.empty()
    
                dl1, dl2 = st.columns(2)
                with dl1:
                    st.download_button(
                        label="Download By SKU.xlsx",
                        data=intermediate_bytes,
                        file_name="By SKU.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                with dl2:
                    st.download_button(
                        label=f"Download {final_filename}",
                        data=final_bytes,
                        file_name=final_filename,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )

    except Exception as e:
        st.error(f"Error: {e}")
