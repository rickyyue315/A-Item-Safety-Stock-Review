import io
import datetime
import re

import streamlit as st
import openpyxl
from openpyxl.styles import PatternFill
from openpyxl.utils import get_column_letter
import pandas as pd

SAFETY_ORDER = {
    "Aitem<25%": 0, "New item<15%": 1, "B/C item<15%": 2,
    "Aitem>75%": 3, "New item>100%": 4, "B/C item>100%": 5,
    "OK": 6, "Check": 7
}
YELLOW_FILL = PatternFill(start_color="FFFFFF00", end_color="FFFFFF00", fill_type="solid")
AMBER_FILL = PatternFill(start_color="FFFFC000", end_color="FFFFC000", fill_type="solid")
PCT_FORMAT = "0.00%"

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

[data-testid="stMetricValue"] {
    font-weight: 900 !important;
    font-size: 2.8rem !important;
}
[data-testid="stMetricLabel"] {
    font-weight: 700 !important;
    font-size: 1rem !important;
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
                return "B/C item<15%"
            if ratio > 1.0:
                return "B/C item>100%"
            return "OK"
    else:
        return "Check"


def extract_buyer_number(buyer_str):
    m = re.match(r"(\d+)", str(buyer_str))
    return int(m.group(1)) if m else 0


def process_uploaded_file(uploaded_bytes):
    wb_in = openpyxl.load_workbook(io.BytesIO(uploaded_bytes), data_only=True)
    ws = wb_in["Export"]
    raw_headers = [cell.value for cell in ws[1]]

    item_col_idx = raw_headers.index("Item")
    buyer_col_idx = raw_headers.index("Buyer")
    ratio_col_idx = raw_headers.index("Safety VS Monthly Sold Qty")
    item_class_idx = raw_headers.index("Item Class")

    data_rows = []
    for r in range(2, ws.max_row + 1):
        col_a_val = ws.cell(row=r, column=1).value
        if col_a_val == "Total" or (isinstance(col_a_val, str) and col_a_val.startswith("Applied filters:")):
            continue
        item_val = ws.cell(row=r, column=item_col_idx + 1).value
        buyer_val = ws.cell(row=r, column=buyer_col_idx + 1).value
        if buyer_val is None:
            continue
        row_values = [ws.cell(row=r, column=c).value for c in range(1, ws.max_column + 1)]
        ratio_val = row_values[ratio_col_idx]
        item_class_val = row_values[item_class_idx]
        sasa_abc = str(item_class_val)[0] if item_class_val else ""
        safety_val = get_safety_checking(ratio_val, sasa_abc)
        data_rows.append((row_values, safety_val))

    wb_in.close()

    new_headers = list(raw_headers)
    new_headers.insert(raw_headers.index("Item Class") + 1, "SASA ABC")
    new_headers.insert(new_headers.index("Safety VS Monthly Sold Qty") + 1, "Safety Checking")

    return raw_headers, new_headers, data_rows


def build_row(raw_values, safety_val, raw_headers):
    row = list(raw_values)
    item_class_val = raw_values[raw_headers.index("Item Class")]
    sasa_abc = str(item_class_val)[0] if item_class_val else ""
    row.insert(raw_headers.index("Item Class") + 1, sasa_abc)
    row.insert(raw_headers.index("Safety VS Monthly Sold Qty") + 1 + 1, safety_val)
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

    safety_col = new_headers.index("Safety Checking") + 1
    for r_idx, (raw_values, safety_val) in enumerate(sorted_data, 2):
        row = build_row(raw_values, safety_val, raw_headers)
        for c, val in enumerate(row, 1):
            ws_sku.cell(row=r_idx, column=c, value=val)
        ws_sku.cell(row=r_idx, column=safety_col).fill = AMBER_FILL

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


st.set_page_config(page_title="Safety Stock Review", page_icon="\U0001f4ca", layout="wide")
st.markdown(DARK_THEME_CSS, unsafe_allow_html=True)

st.title("A Item Safety Stock Review vs Monthly Sold Qty")
st.markdown("---")

uploaded_file = st.file_uploader(
    "\U0001f4c1  Upload By SKU_RAW.xlsx",
    type=["xlsx"],
    help="Select the RAW export file with 'Export' sheet"
)

if uploaded_file is not None:
    with st.spinner("\U0001f504 Processing..."):
        try:
            file_bytes = uploaded_file.read()
            raw_headers, new_headers, data_rows = process_uploaded_file(file_bytes)

            categories_display = sorted(set(sv for _, sv in data_rows), key=lambda x: SAFETY_ORDER.get(x, 99))
            safety_counts = {cat: 0 for cat in categories_display}
            for _, sv in data_rows:
                safety_counts[sv] += 1

            st.subheader("Key Metrics")
            cols = st.columns(len(categories_display) + 1)
            cols[0].metric("Total SKU", len(data_rows))
            for i, cat in enumerate(categories_display):
                cols[i + 1].metric(cat, safety_counts[cat])

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

            problem_data = []
            for raw_values, safety_val in data_rows:
                if safety_val == "Aitem<25%":
                    buyer = raw_values[buyer_col_idx]
                    item = raw_values[raw_headers.index("Item")]
                    sku = raw_values[raw_headers.index("SKU")]
                    sku_desc = raw_values[raw_headers.index("SKU Desc")]
                    brand = raw_values[raw_headers.index("Brand")]
                    ratio = raw_values[raw_headers.index("Safety VS Monthly Sold Qty")]
                    problem_data.append({
                        "Buyer": buyer,
                        "Item": item,
                        "SKU": sku,
                        "SKU Desc": sku_desc,
                        "Brand": brand,
                        "Safety/Monthly %": f"{ratio*100:.1f}%" if ratio else "N/A"
                    })

            if problem_data:
                problem_df = pd.DataFrame(problem_data)

                tab1, tab2 = st.tabs(["All Problem SKUs", "By Buyer"])
                with tab1:
                    st.dataframe(problem_df, use_container_width=True, hide_index=True)
                    st.info(f"**{len(problem_df)}** Aitem<25% problem SKUs found")

                with tab2:
                    buyer_summary = problem_df.groupby("Buyer").agg(
                        Problem_SKU_Count=("Item", "count"),
                        Brands=("Brand", lambda x: ", ".join(sorted(set(x))))
                    ).reset_index()
                    buyer_summary = buyer_summary.sort_values("Problem_SKU_Count", ascending=False)
                    st.dataframe(buyer_summary, use_container_width=True, hide_index=True)
            else:
                st.success("No Aitem<25% problem SKUs found!")

            st.markdown("---")

            if st.button("Generate Output Files", type="primary"):
                with st.spinner("Generating..."):
                    wb_intermediate = build_intermediate_wb(new_headers, data_rows, raw_headers)
                    wb_final = build_final_wb(new_headers, data_rows, raw_headers)

                    intermediate_bytes = wb_to_bytes(wb_intermediate)

                    today = datetime.date.today()
                    date_str = today.strftime("%d%b%Y")
                    final_filename = f"Check Safety A Item by SKU_To Buyer_{date_str}.xlsx"

                    final_bytes = wb_to_bytes(wb_final)

                st.success("Files generated successfully!")

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
