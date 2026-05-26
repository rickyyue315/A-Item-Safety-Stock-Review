import io
import datetime
import re

import streamlit as st
import openpyxl
from openpyxl.styles import PatternFill
from openpyxl.utils import get_column_letter
import pandas as pd

SAFETY_ORDER = {"Aitem<25%": 0, "Aitem>75%": 1, "OK": 2}
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


def process_uploaded_file(uploaded_bytes):
    wb_in = openpyxl.load_workbook(io.BytesIO(uploaded_bytes), data_only=True)
    ws = wb_in["Export"]
    raw_headers = [cell.value for cell in ws[1]]

    item_col_idx = raw_headers.index("Item")
    buyer_col_idx = raw_headers.index("Buyer")
    ratio_col_idx = raw_headers.index("Safety VS Monthly Sold Qty")

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
        safety_val = get_safety_checking(ratio_val)
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

    buyer_data = {}
    for raw_values, safety_val in data_rows:
        buyer = raw_values[buyer_col_idx]
        if buyer not in buyer_data:
            buyer_data[buyer] = {"Aitem<25%": 0, "Aitem>75%": 0, "OK": 0}
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

            safety_counts = {"Aitem<25%": 0, "Aitem>75%": 0, "OK": 0}
            for _, sv in data_rows:
                safety_counts[sv] += 1

            st.subheader("Key Metrics")
            m1, m2, m3, m4 = st.columns(4)
            m1.metric("Total SKU", len(data_rows))
            m2.metric("Aitem<25%", safety_counts["Aitem<25%"])
            m3.metric("Aitem>75%", safety_counts["Aitem>75%"])
            m4.metric("OK", safety_counts["OK"])

            st.markdown("---")

            buyer_col_idx = raw_headers.index("Buyer")
            buyer_data = {}
            for raw_values, safety_val in data_rows:
                buyer = raw_values[buyer_col_idx]
                if buyer not in buyer_data:
                    buyer_data[buyer] = {"Aitem<25%": 0, "Aitem>75%": 0, "OK": 0}
                buyer_data[buyer][safety_val] += 1

            sorted_buyers = sorted(buyer_data.keys(), key=extract_buyer_number)
            categories = ["Aitem<25%", "Aitem>75%", "OK"]

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
