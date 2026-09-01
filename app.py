import io
import datetime
import time

import streamlit as st
import openpyxl
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter
import pandas as pd

from core import (
    SAFETY_ORDER,
    YELLOW_FILL,
    AMBER_FILL,
    DARK_RED_FONT,
    PCT_FORMAT,
    YELLOW_COLUMNS,
    get_safety_checking,
    extract_buyer_number,
    get_col_idx,
    compute_nd_checking,
    filter_ndrf_data,
    get_new_launch_flag,
    get_launch_days,
    get_default_reference_date,
    DEFAULT_TOTAL_SHOPS,
    is_aitem_overstock_short_d001,
    is_aitem_overstock_short_d001_tier2,
    build_row,
    build_intermediate_wb,
    build_final_wb,
    build_nd_wb,
    wb_to_bytes,
)

CACHE_TTL_SECONDS = 10 * 60
IDLE_TIMEOUT_SECONDS = 15 * 60

# set_page_config must be the first Streamlit UI command for every run.
st.set_page_config(page_title="Safety Stock Review", page_icon="\U0001f4ca", layout="wide")


def _clear_sensitive_data(reset_upload=False):
    """Clear server-side cached data and optionally reset the upload widget."""
    st.cache_data.clear()
    if reset_upload:
        current_version = int(st.session_state.get("_upload_widget_version", 0))
        st.session_state["_upload_widget_version"] = current_version + 1


def _cleanup_after_download():
    """Clean cached upload/processing data after a download button is clicked."""
    _clear_sensitive_data(reset_upload=True)
    st.session_state["_last_activity"] = time.time()


def _enforce_idle_cleanup():
    """Clear temporary data after 15 minutes without a Streamlit interaction."""
    now = time.time()
    last_activity = st.session_state.get("_last_activity")
    expired = last_activity is not None and now - last_activity > IDLE_TIMEOUT_SECONDS
    if expired:
        _clear_sensitive_data(reset_upload=True)
    st.session_state["_last_activity"] = now
    return expired

LIGHT_THEME_CSS = """
<style>
.stApp, .main {
    background: #ffffff !important;
    color: #1f2328 !important;
}

h1 {
    color: #0a3d62 !important;
    text-shadow: none;
    font-weight: 900 !important;
}
h2, h3 {
    color: #b3541e !important;
    text-shadow: none;
}

hr {
    border: none;
    height: 2px;
    background: linear-gradient(90deg, transparent, #0a3d62, transparent) !important;
    margin: 1.5rem 0 !important;
}

[data-testid="stDataFrame"] {
    background: #ffffff !important;
    border-radius: 12px !important;
    border: 1px solid #d0d7de !important;
}
[data-testid="stDataFrame"] th {
    background: linear-gradient(135deg, #f6f8fa 0%, #eaeef2 100%) !important;
    color: #0a3d62 !important;
    font-weight: 700 !important;
    border-bottom: 2px solid #b3541e !important;
}
[data-testid="stDataFrame"] td {
    color: #1f2328 !important;
    border-color: #e1e4e8 !important;
}
[data-testid="stDataFrame"] tr:hover {
    background: #eef4fb !important;
}

.stButton > button {
    background: linear-gradient(135deg, #0969da 0%, #084298 100%) !important;
    color: #ffffff !important;
    border: none !important;
    border-radius: 12px !important;
    font-weight: 700 !important;
    padding: 0.8rem 2.5rem !important;
    box-shadow: 0 4px 12px rgba(9,105,218,0.3) !important;
    transition: all 0.3s ease !important;
}
.stButton > button:hover {
    transform: scale(1.05);
    box-shadow: 0 6px 20px rgba(9,105,218,0.45) !important;
}

.stDownloadButton > button {
    background: linear-gradient(135deg, #1a7f37 0%, #0f5e25 100%) !important;
    color: #ffffff !important;
    font-weight: 700 !important;
}

.streamlit-expanderHeader {
    background: linear-gradient(90deg, rgba(9,105,218,0.08) 0%, rgba(179,84,30,0.06) 100%) !important;
    border-radius: 12px !important;
    border: 1px solid #d0d7de !important;
    color: #0a3d62 !important;
    font-weight: 700 !important;
}

[data-testid="stFileUploadDropzone"] {
    background: #f6f8fa !important;
    border: 2px dashed #0969da !important;
    border-radius: 16px !important;
}

.stSuccess {
    background: rgba(26,127,55,0.1) !important;
    border: 1px solid #1a7f37 !important;
    color: #0f5e25 !important;
}
.stError {
    background: rgba(207,34,46,0.1) !important;
    border: 1px solid #cf242e !important;
    color: #a40e26 !important;
}

.stTabs [data-baseweb="tab"] {
    background: #f6f8fa !important;
    border-radius: 8px 8px 0 0 !important;
    color: #1f2328 !important;
}
.stTabs [aria-selected="true"] {
    background: rgba(9,105,218,0.12) !important;
    color: #0969da !important;
    font-weight: 700 !important;
}

.metric-grid {
    display: grid;
    grid-template-columns: repeat(4, 1fr);
    gap: 16px;
    margin: 1rem 0 1.5rem 0;
}

.metric-card {
    background: linear-gradient(135deg, #ffffff 0%, #f6f8fa 100%);
    border-radius: 14px;
    border: 1px solid #d0d7de;
    padding: 1.25rem 1.5rem;
    position: relative;
    overflow: hidden;
    transition: all 0.3s ease;
    min-height: 130px;
    display: flex;
    align-items: stretch;
    box-shadow: 0 2px 8px rgba(0,0,0,0.05);
}

.metric-card:hover {
    transform: translateY(-3px);
    box-shadow: 0 12px 40px rgba(0,0,0,0.12);
    border-color: #0969da;
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
    color: #0a3d62;
    letter-spacing: 0.5px;
}

.metric-card__value {
    font-size: 2.1rem;
    font-weight: 900;
    color: #1f2328;
    line-height: 1.1;
    letter-spacing: -0.5px;
}

.metric-card__description {
    font-size: 1.05rem;
    color: #57606a;
    margin-top: 1px;
    line-height: 1.3;
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


@st.cache_data(show_spinner=False, ttl=CACHE_TTL_SECONDS, max_entries=1)
def process_uploaded_file_cached(file_bytes: bytes):
    """Cached version of file processing without progress callback."""
    start_time = time.time()
    wb_in = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True, read_only=True)
    ws = wb_in["Export"]
    raw_headers = [cell.value for cell in ws[1]]
    col_idx = get_col_idx(raw_headers)

    item_col_idx = col_idx["Item"]
    buyer_col_idx = col_idx["Buyer"]
    ratio_col_idx = col_idx["Safety VS Monthly Sold Qty"]
    item_class_idx = col_idx["Item Class"]
    sku_desc_col_idx = col_idx["SKU Desc"]

    data_rows = []
    for r, row in enumerate(ws.iter_rows(min_row=2, values_only=True), 2):
        if not row:
            continue
        col_a_val = row[0]
        if col_a_val == "Total" or (isinstance(col_a_val, str) and col_a_val.startswith("Applied filters:")):
            continue
        item_val = row[item_col_idx]
        buyer_val = row[buyer_col_idx]
        if buyer_val is None or item_val is None:
            continue
        sku_desc_val = str(row[sku_desc_col_idx]) if row[sku_desc_col_idx] else ""
        if "(STAFF)" in sku_desc_val or "(MAC)" in sku_desc_val:
            continue
        ratio_val = row[ratio_col_idx]
        item_class_val = row[item_class_idx]
        sasa_abc = str(item_class_val)[0] if item_class_val else ""
        safety_val = get_safety_checking(ratio_val, sasa_abc)
        data_rows.append((list(row), safety_val))

    wb_in.close()

    new_headers = list(raw_headers)
    new_headers.insert(col_idx["Item Class"] + 1, "SASA ABC")
    new_headers.insert(new_headers.index("Safety VS Monthly Sold Qty") + 1, "Safety Checking")

    elapsed = time.time() - start_time
    return raw_headers, new_headers, data_rows, elapsed


@st.cache_data(show_spinner=False, ttl=CACHE_TTL_SECONDS, max_entries=1)
def _compute_nd_cached(file_bytes, total_shops, reference_date, sub_mode):
    """Pure ND computation (no Streamlit calls)."""
    wb_in = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True, read_only=True)
    ws = wb_in["Export"]
    raw_headers, data_rows = filter_ndrf_data(ws, progress_callback=None)
    wb_in.close()

    col_idx = get_col_idx(raw_headers)
    buyer_col_idx = col_idx["Buyer"]
    launch_date_col_idx = col_idx["Launch Date"]
    rf_shop_col_idx = col_idx["No. of RF Shop"]
    today = reference_date if reference_date is not None else datetime.date.today()

    enriched_data = []
    for raw_values in data_rows:
        nd_ratios, nd_checking = compute_nd_checking(raw_values[rf_shop_col_idx], total_shops)
        new_launch_flag = get_new_launch_flag(raw_values[launch_date_col_idx], today)
        launch_days = get_launch_days(raw_values[launch_date_col_idx], today)
        enriched_data.append((raw_values, nd_ratios, nd_checking, new_launch_flag, launch_days))

    if sub_mode == "只顯示 New SKU":
        enriched_data = [x for x in enriched_data if x[3] == "New SKU"]

    buyer_counts = {}
    for rv, _, nc, *_ in enriched_data:
        buyer = rv[buyer_col_idx]
        if buyer not in buyer_counts:
            buyer_counts[buyer] = {"ND Ratio > 80%": 0, "ND Ratio 70%-80%": 0, "Total": 0}
        if nc == "ND Ratio > 80%":
            buyer_counts[buyer]["ND Ratio > 80%"] += 1
        elif nc == "ND Ratio 70%-80%":
            buyer_counts[buyer]["ND Ratio 70%-80%"] += 1
        buyer_counts[buyer]["Total"] += 1

    return raw_headers, enriched_data, buyer_counts


def process_ndrf_mode(file_bytes, total_shops, sub_mode="顯示所有 SKU", reference_date=None):
    with st.spinner("Processing ND data..."):
        raw_headers, enriched_data, buyer_counts = _compute_nd_cached(
            file_bytes, total_shops, reference_date, sub_mode,
        )

    col_idx = get_col_idx(raw_headers)
    buyer_col_idx = col_idx["Buyer"]

    total_sku = len(enriched_data)
    gt_80 = btw_70_80 = check_nd = ok_nd = 0
    for item in enriched_data:
        nc = item[2] if len(item) >= 3 else None
        if nc == "ND Ratio > 80%":
            gt_80 += 1
        elif nc == "ND Ratio 70%-80%":
            btw_70_80 += 1
        elif nc == "Check":
            check_nd += 1
        else:
            ok_nd += 1

    METRIC_ACCENT_ND = {
        "Total Filtered SKU": "#0969da",
        "ND Ratio > 80%": "#d6336c",
        "ND Ratio 70%-80%": "#e8590c",
        "Check": "#f08c00",
        "ND Ratio < 70%": "#2f9e44",
    }
    METRIC_DESC_ND = {
        "Total Filtered SKU": "SKUs after first-layer filtering",
        "ND Ratio > 80%": "ND ratio exceeds 80% threshold",
        "ND Ratio 70%-80%": "ND ratio between 70%-80%",
        "Check": "Data anomaly: RF shops exceed total shops",
        "ND Ratio < 70%": "ND ratio within acceptable range",
    }

    st.subheader("Key Metrics")
    metric_items = [
        ("Total Filtered SKU", total_sku),
        ("ND Ratio > 80%", gt_80),
        ("ND Ratio 70%-80%", btw_70_80),
        ("Check", check_nd),
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
    st.caption("Filtered: Warehouse only ≥ 100 units; Vendor only ≥ 200 units. 'Total' = all filtered SKUs (incl. ND Ratio < 70%).")
    st.caption(f"Current Total Shops = {total_shops}")
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
    # Show every row (incl. Grand Total) without scrolling, for easy screenshotting.
    # height = header (~40px) + rows (each ~35px) + small buffer.
    df_height = len(summary_df) * 35 + 55
    st.dataframe(summary_df, use_container_width=True, hide_index=True, height=df_height)

    st.markdown("---")

    sku_col_idx = col_idx["SKU"]
    sku_desc_col_idx = col_idx["SKU Desc"]
    brand_col_idx = col_idx["Brand"]
    rf_shop_idx = col_idx["No. of RF Shop"]
    nd_shop_stock_idx = col_idx["No. of ND Shop with Stock"]

    gt80_data = []
    btw70_80_data = []
    for item in enriched_data:
        if len(item) == 5:
            raw_values, nd_ratios, nd_checking, new_launch_flag, launch_days = item
        else:
            raw_values, nd_ratios, nd_checking, new_launch_flag = item
            launch_days = get_launch_days(raw_values[col_idx["Launch Date"]], reference_date if reference_date is not None else datetime.date.today())
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
                "Days Since Launch": launch_days,
            })
        elif nd_checking == "ND Ratio 70%-80%":
            btw70_80_data.append({
                "Buyer": raw_values[buyer_col_idx],
                "SKU": raw_values[sku_col_idx],
                "SKU Desc": raw_values[sku_desc_col_idx],
                "Brand": raw_values[brand_col_idx],
                "No. of RF Shop": raw_values[rf_shop_idx],
                "No. of ND Shop with Stock": raw_values[nd_shop_stock_idx],
                "ND ratios": f"{nd_ratios*100:.1f}%",
                "ND Checking": nd_checking,
                "New Launch Flag (L3M)": new_launch_flag,
                "Days Since Launch": launch_days,
            })

    if gt80_data:
        with st.expander(f"ND Ratio > 80% Problem SKUs ({len(gt80_data)} SKUs)"):
            gt80_df = pd.DataFrame(gt80_data)
            st.dataframe(gt80_df, use_container_width=True, hide_index=True, height=min(len(gt80_df), 50) * 35 + 38)

    if btw70_80_data:
        with st.expander(f"ND Ratio 70%-80% Problem SKUs ({len(btw70_80_data)} SKUs)"):
            btw70_80_df = pd.DataFrame(btw70_80_data)
            st.dataframe(btw70_80_df, use_container_width=True, hide_index=True, height=min(len(btw70_80_df), 50) * 35 + 38)

    st.markdown("---")

    if st.button("Generate Output Files", type="primary"):
        gen_progress = st.progress(0)
        gen_status = st.empty()
        gen_start = time.time()

        gen_status.markdown("📝 Building workbook (RAW + Summary + problem sheets)...")
        gen_progress.progress(0.3)

        wb = build_nd_wb(enriched_data, buyer_counts, raw_headers, total_shops, reference_date)

        gen_progress.progress(0.8)
        gen_status.markdown("💾 Converting to bytes...")
        wb_bytes = wb_to_bytes(wb)

        gen_progress.progress(1.0)
        today = datetime.date.today()
        date_str = today.strftime("%d%b%Y")
        filename = f"Check NDRF_To Buyer_{date_str}.xlsx"
        gen_m, gen_s = divmod(int(time.time() - gen_start), 60)
        gen_status.markdown(f"✅ **File generated!** (`{gen_m:02d}:{gen_s:02d}`)")

        st.caption("下載後系統會清除暫存資料，需重新上傳檔案。")
        st.download_button(
            label=f"Download {filename}",
            data=wb_bytes,
            file_name=filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            on_click=_cleanup_after_download,
        )



if _enforce_idle_cleanup():
    st.info("因閒置超過 15 分鐘，已清除暫存資料。")

st.markdown(LIGHT_THEME_CSS, unsafe_allow_html=True)
if st.sidebar.button("清除暫存", key="cleanup_button"):
    _clear_sensitive_data(reset_upload=True)
    st.rerun()
st.sidebar.caption("閒置 15 分鐘後自動清除暫存")

st.title("A Item Safety Stock Review vs Monthly Sold Qty")
st.markdown("---")

mode = st.radio("Mode", ["Review Safety", "ND Ratio Review"], index=0, horizontal=True)

uploaded_file = st.file_uploader(
    "\U0001f4c1  Upload By SKU_RAW.xlsx",
    type=["xlsx"],
    help="Select the RAW export file with 'Export' sheet",
    key=f"uploaded_file_{st.session_state.get('_upload_widget_version', 0)}"
)

if mode == "ND Ratio Review":
    total_shops = st.number_input(
        "Total Shops",
        min_value=1, max_value=500, value=DEFAULT_TOTAL_SHOPS, step=1,
        help="Total number of shops used for ND ratio calculation"
    )
    st.caption(f"備註: Current Total Shops = {total_shops}")
    reference_date = st.date_input(
        "Reference Date (New Launch cutoff)",
        value=get_default_reference_date(),
        help="Used to decide New SKU (launch within last 3 months). "
             "Defaults to yesterday (T-1) in Hong Kong time, since uploaded data is usually T-1."
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
                process_ndrf_mode(file_bytes, total_shops, nd_sub_mode, reference_date)
            except Exception as e:
                st.error(f"ND Mode Error: {e}")
                st.exception(e)

        else:
            # Safety Review path. File parsing is cached by Streamlit
            # (@st.cache_data on process_uploaded_file_cached), so the second run
            # onward is instant and there is no per-row progress to report.
            # Show an honest "processing / done" status instead of a fake bar.
            status_text = st.empty()
            elapsed_text = st.empty()
            processing_start = time.time()

            status_text.markdown("🔄 **Processing file...**")
            raw_headers, new_headers, data_rows, file_elapsed = process_uploaded_file_cached(file_bytes)

            total_elapsed = time.time() - processing_start
            mins, secs = divmod(int(total_elapsed), 60)
            status_text.markdown(
                f"✅ **Read complete!** `{len(data_rows):,}` valid rows processed"
            )
            elapsed_text.markdown(
                f"⏱️ **Total read time:** `{mins:02d}:{secs:02d}`"
            )

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
                "Aitem>200% & D001<20pcs": "A Item > 200% & D001 < 20pcs",
                "Aitem>150% & D001<50pcs": "A Item > 150% & D001 < 50pcs",
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
                "Aitem>200% & D001<20pcs": "Safety stock >200% with D001 below 20pcs",
                "Aitem>150% & D001<50pcs": "Safety stock >150% with D001 below 50pcs",
            }
            METRIC_ACCENT = {
                "Total SKU": "#0969da",
                "Aitem<25%": "#d6336c",
                "New item<15%": "#e8590c",
                "B/C/D item<15%": "#f08c00",
                "Aitem>75%": "#6741d9",
                "New item>100%": "#c2255c",
                "B/C/D item>100%": "#e64980",
                "OK": "#2f9e44",
                "Check": "#f08c00",
                "Aitem>200% & D001<20pcs": "#9a3412",
                "Aitem>150% & D001<50pcs": "#e8590c",
            }

            metric_items = [("Total SKU", len(data_rows))] + [
                (cat, safety_counts[cat]) for cat in categories_display
            ]

            kpi_col_idx = get_col_idx(raw_headers)
            kpi_ratio_idx = kpi_col_idx["Safety VS Monthly Sold Qty"]
            kpi_item_class_idx = kpi_col_idx["Item Class"]
            kpi_supply_idx = kpi_col_idx["Supply Source"]
            kpi_d001_idx = kpi_col_idx["D001 Net Stock Qty (CM)"]

            kpi_count_200 = sum(
                1 for raw_values, _ in data_rows
                if is_aitem_overstock_short_d001(
                    raw_values[kpi_ratio_idx],
                    raw_values[kpi_item_class_idx],
                    raw_values[kpi_supply_idx],
                    raw_values[kpi_d001_idx],
                )
            )
            kpi_count_150 = sum(
                1 for raw_values, _ in data_rows
                if is_aitem_overstock_short_d001_tier2(
                    raw_values[kpi_ratio_idx],
                    raw_values[kpi_item_class_idx],
                    raw_values[kpi_supply_idx],
                    raw_values[kpi_d001_idx],
                )
            )

            metric_items += [
                ("Aitem>200% & D001<20pcs", kpi_count_200),
                ("Aitem>150% & D001<50pcs", kpi_count_150),
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

            col_idx = get_col_idx(raw_headers)
            buyer_col_idx = col_idx["Buyer"]
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
            st.dataframe(summary_df, use_container_width=True, hide_index=True, height=min(len(summary_df), 50) * 35 + 38)

            st.markdown("---")

            st.subheader("Aitem<25% Problem SKUs & Buyer Info")

            supply_source_col_idx = col_idx["Supply Source"]
            shop_stock_col_idx = col_idx["Shop Stock Qty (CM)"]
            safety_stock_col_idx = col_idx["Adj. Safety Stock"]
            shop_monthly_sold_col_idx = col_idx["Shop Monthly Sold Qty (P90D)"]
            d001_net_stock_col_idx = col_idx["D001 Net Stock Qty (CM)"]
            sku_col_idx = col_idx["SKU"]
            sku_desc_col_idx = col_idx["SKU Desc"]
            brand_col_idx = col_idx["Brand"]
            ratio_col_idx = col_idx["Safety VS Monthly Sold Qty"]

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
                        "Safety/Monthly %": f"{ratio*100:.1f}%" if ratio is not None else "N/A",
                        "D001 Net Stock Qty (CM)": d001_net_stock,
                        "Shop Stock Qty (CM)": shop_stock
                    })

            if problem_data:
                problem_df = pd.DataFrame(problem_data)

                tab1, tab2 = st.tabs(["All Problem SKUs", "By Buyer"])
                with tab1:
                    st.dataframe(problem_df, use_container_width=True, hide_index=True, height=min(len(problem_df), 50) * 35 + 38)
                    st.info(f"**{len(problem_df)}** Aitem<25% problem SKUs found")

                with tab2:
                    buyer_summary = problem_df.groupby("Buyer").agg(
                        Problem_SKU_Count=("SKU", "count"),
                        Brands=("Brand", lambda x: ", ".join(sorted(set(x))))
                    ).reset_index()
                    buyer_summary = buyer_summary.sort_values("Problem_SKU_Count", ascending=False)
                    st.dataframe(buyer_summary, use_container_width=True, hide_index=True, height=min(len(buyer_summary), 50) * 35 + 38)
            else:
                st.success("No Aitem<25% problem SKUs found!")

            # ── New Section: Aitem>200% & D001<20pcs ──
            st.subheader("Aitem>200% & D001<20pcs Problem SKUs & Buyer Info")
            st.caption("Supply Source = Warehouse only (Vendor items do not go through D001)")

            ratio_col_idx_200 = col_idx["Safety VS Monthly Sold Qty"]
            supply_source_col_idx = col_idx["Supply Source"]
            d001_net_stock_col_idx = col_idx["D001 Net Stock Qty (CM)"]
            shop_stock_col_idx = col_idx["Shop Stock Qty (CM)"]
            safety_stock_col_idx = col_idx["Adj. Safety Stock"]
            shop_monthly_sold_col_idx = col_idx["Shop Monthly Sold Qty (P90D)"]
            sku_col_idx_200 = col_idx["SKU"]
            sku_desc_col_idx_200 = col_idx["SKU Desc"]
            brand_col_idx_200 = col_idx["Brand"]
            item_class_idx = col_idx["Item Class"]

            problem_data_200 = []
            for raw_values, safety_val in data_rows:
                ratio = raw_values[ratio_col_idx_200]
                item_class_val = raw_values[item_class_idx]
                supply_source = raw_values[supply_source_col_idx]
                d001_net_stock = raw_values[d001_net_stock_col_idx]

                if is_aitem_overstock_short_d001(ratio, item_class_val, supply_source, d001_net_stock):
                    buyer = raw_values[buyer_col_idx]
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
                        "Supply Source": str(supply_source) if supply_source else "",
                        "Safety Stock Qty": safety_stock,
                        "Shop Monthly Sold Qty (P90D)": int(shop_monthly_sold) if shop_monthly_sold is not None else None,
                        "Safety/Monthly %": f"{ratio*100:.1f}%" if ratio is not None else "N/A",
                        "D001 Net Stock Qty (CM)": d001_net_stock,
                        "Shop Stock Qty (CM)": shop_stock
                    })

            if problem_data_200:
                problem_df_200 = pd.DataFrame(problem_data_200)

                tab_200_1, tab_200_2 = st.tabs(["All Problem SKUs", "By Buyer"])
                with tab_200_1:
                    st.dataframe(problem_df_200, use_container_width=True, hide_index=True, height=min(len(problem_df_200), 50) * 35 + 38)
                    st.info(f"**{len(problem_df_200)}** Aitem>200% & D001<20pcs problem SKUs found")

                with tab_200_2:
                    buyer_summary_200 = problem_df_200.groupby("Buyer").agg(
                        Problem_SKU_Count=("SKU", "count"),
                        Brands=("Brand", lambda x: ", ".join(sorted(set(x))))
                    ).reset_index()
                    buyer_summary_200 = buyer_summary_200.sort_values("Problem_SKU_Count", ascending=False)
                    st.dataframe(buyer_summary_200, use_container_width=True, hide_index=True, height=min(len(buyer_summary_200), 50) * 35 + 38)
            else:
                st.success("No Aitem>200% & D001<20pcs problem SKUs found!")

            # ── New Section: Aitem>150% & D001<50pcs ──
            st.subheader("Aitem>150% & D001<50pcs Problem SKUs & Buyer Info")
            st.caption("Supply Source = Warehouse only (Vendor items do not go through D001). Includes the Aitem>200% & D001<20pcs SKUs above.")

            ratio_col_idx_150 = col_idx["Safety VS Monthly Sold Qty"]
            supply_source_col_idx_t2 = col_idx["Supply Source"]
            d001_net_stock_col_idx_t2 = col_idx["D001 Net Stock Qty (CM)"]
            shop_stock_col_idx_t2 = col_idx["Shop Stock Qty (CM)"]
            safety_stock_col_idx_t2 = col_idx["Adj. Safety Stock"]
            shop_monthly_sold_col_idx_t2 = col_idx["Shop Monthly Sold Qty (P90D)"]
            sku_col_idx_150 = col_idx["SKU"]
            sku_desc_col_idx_150 = col_idx["SKU Desc"]
            brand_col_idx_150 = col_idx["Brand"]
            item_class_idx_t2 = col_idx["Item Class"]

            problem_data_150 = []
            for raw_values, safety_val in data_rows:
                ratio = raw_values[ratio_col_idx_150]
                item_class_val = raw_values[item_class_idx_t2]
                supply_source = raw_values[supply_source_col_idx_t2]
                d001_net_stock = raw_values[d001_net_stock_col_idx_t2]

                if is_aitem_overstock_short_d001_tier2(ratio, item_class_val, supply_source, d001_net_stock):
                    buyer = raw_values[buyer_col_idx]
                    sku = raw_values[sku_col_idx_150]
                    sku_desc = raw_values[sku_desc_col_idx_150]
                    brand = raw_values[brand_col_idx_150]
                    shop_stock = raw_values[shop_stock_col_idx_t2]
                    safety_stock = raw_values[safety_stock_col_idx_t2]
                    shop_monthly_sold = raw_values[shop_monthly_sold_col_idx_t2]
                    problem_data_150.append({
                        "Buyer": buyer,
                        "SKU": sku,
                        "SKU Desc": sku_desc,
                        "Brand": brand,
                        "Supply Source": str(supply_source) if supply_source else "",
                        "Safety Stock Qty": safety_stock,
                        "Shop Monthly Sold Qty (P90D)": int(shop_monthly_sold) if shop_monthly_sold is not None else None,
                        "Safety/Monthly %": f"{ratio*100:.1f}%" if ratio is not None else "N/A",
                        "D001 Net Stock Qty (CM)": d001_net_stock,
                        "Shop Stock Qty (CM)": shop_stock
                    })

            if problem_data_150:
                problem_df_150 = pd.DataFrame(problem_data_150)

                tab_150_1, tab_150_2 = st.tabs(["All Problem SKUs", "By Buyer"])
                with tab_150_1:
                    st.dataframe(problem_df_150, use_container_width=True, hide_index=True, height=min(len(problem_df_150), 50) * 35 + 38)
                    st.info(f"**{len(problem_df_150)}** Aitem>150% & D001<50pcs problem SKUs found")

                with tab_150_2:
                    buyer_summary_150 = problem_df_150.groupby("Buyer").agg(
                        Problem_SKU_Count=("SKU", "count"),
                        Brands=("Brand", lambda x: ", ".join(sorted(set(x))))
                    ).reset_index()
                    buyer_summary_150 = buyer_summary_150.sort_values("Problem_SKU_Count", ascending=False)
                    st.dataframe(buyer_summary_150, use_container_width=True, hide_index=True, height=min(len(buyer_summary_150), 50) * 35 + 38)
            else:
                st.success("No Aitem>150% & D001<50pcs problem SKUs found!")

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

                st.caption("下載任一檔案後系統會清除暫存資料，需重新上傳檔案。")
                dl1, dl2 = st.columns(2)
                with dl1:
                    st.download_button(
                        label="Download By SKU.xlsx",
                        data=intermediate_bytes,
                        file_name="By SKU.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        on_click=_cleanup_after_download,
                    )
                with dl2:
                    st.download_button(
                        label=f"Download {final_filename}",
                        data=final_bytes,
                        file_name=final_filename,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        on_click=_cleanup_after_download,
                    )

    except Exception as e:
        st.error(f"Error: {e}")
        st.exception(e)
