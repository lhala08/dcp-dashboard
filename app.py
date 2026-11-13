import io
import requests
import numpy as np
import pandas as pd
import plotly.express as px
import streamlit as st

# ==========================================
# PAGE CONFIG & EXECUTIVE STYLING
# ==========================================
st.set_page_config(
    page_title="DCP Delivery Monitoring Dashboard",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown(
    """
    <style>
    .block-container {
        padding-top: 1.5rem;
        padding-bottom: 2rem;
    }
    .header-banner {
        background: linear-gradient(135deg, #0f294a 0%, #1d4ed8 100%);
        padding: 1.4rem 1.8rem;
        border-radius: 12px;
        color: white;
        margin-bottom: 1.2rem;
        box-shadow: 0 4px 14px rgba(15, 41, 74, 0.15);
    }
    .header-banner h1 {
        margin: 0;
        font-size: 1.65rem;
        font-weight: 700;
        letter-spacing: -0.01em;
    }
    .header-banner p {
        margin: 0.3rem 0 0 0;
        font-size: 0.92rem;
        color: #dbeafe;
    }
    .kpi-card {
        background-color: #ffffff;
        border: 1px solid #e2e8f0;
        border-left: 5px solid #1d4ed8;
        border-radius: 10px;
        padding: 1rem 1.1rem;
        box-shadow: 0 2px 6px rgba(15, 23, 42, 0.04);
    }
    .kpi-title {
        font-size: 0.78rem;
        font-weight: 600;
        text-transform: uppercase;
        color: #64748b;
        margin-bottom: 0.3rem;
        letter-spacing: 0.03em;
    }
    .kpi-value {
        font-size: 1.6rem;
        font-weight: 700;
        color: #0f172a;
        line-height: 1.2;
    }
    .kpi-sub {
        font-size: 0.75rem;
        color: #94a3b8;
        margin-top: 0.25rem;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

SHEET_ID = "17deZpiZgrTQJ9nEBRjFKh6VqHk_hYVs4B7cEaPtcqrw"

REGION_ORDER = {
    "1": 1, "2": 2, "3": 3, "4A": 4, "4B": 5, "MIMAROPA": 6,
    "5": 7, "6": 8, "7": 9, "8": 10, "9": 11, "10": 12,
    "11": 13, "12": 14, "CAR": 15, "CARAGA": 16, "NCR": 17,
    "NIR": 18, "BARMM": 19,
}

PACKAGE_DISPLAY = {
    "laptop for teaching": "Laptop for Teaching",
    "laptop for non-teaching": "Laptop for Non-Teaching",
    "smart tv package": "Smart TV Package",
    "elearning cart package 1": "eLearning Cart Package 1",
    "elearning cart package 2": "eLearning Cart Package 2",
    "elearning cart package 3": "eLearning Cart Package 3",
}

COMPONENT_DISPLAY = {
    "storage/charging cart": "Storage/Charging Cart",
    "storage cart": "Storage Cart",
    "charging cart": "Charging Cart",
    "smart tv": "Smart TV",
    "tablet": "Tablet",
    "laptop": "Laptop",
}


# ==========================================
# DATA LOADING & CLEANING ENGINE
# ==========================================
def clean_series(s: pd.Series) -> pd.Series:
    return (
        s.fillna("")
        .astype(str)
        .str.replace("\xa0", " ", regex=False)
        .str.strip()
    )


def standardize_tab(df_raw: pd.DataFrame, default_fund: str) -> pd.DataFrame:
    # Ensure at least 16 columns (A to P)
    df = df_raw.copy()
    while df.shape[1] < 16:
        df[f"extra_{df.shape[1]}"] = ""

    df = df.iloc[:, :16].copy()
    df.columns = [
        "LOT", "Region", "DCP_Package", "No_of_Package", "School_ID",
        "Recipient_Office", "Component", "Units_Per_Package", "Unit_Cost",
        "Supplier", "DR_Date", "IAR_Date", "IAR_Docs", "Col_N_Status",
        "Status", "Fund_Source"
    ]

    for col in df.columns:
        df[col] = clean_series(df[col])

    df["Fund_Clean"] = default_fund
    df = df[df["DCP_Package"] != ""].copy()
    return df


@st.cache_data(ttl=300, show_spinner=False)
def load_workbook_from_gsheet(sheet_id: str) -> bytes:
    url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=xlsx"
    resp = requests.get(url, timeout=30)
    resp.raise_for_status()
    return resp.content


def process_data(excel_bytes: bytes) -> pd.DataFrame:
    xl = pd.ExcelFile(io.BytesIO(excel_bytes), engine="openpyxl")
    sheet_names = xl.sheet_names

    frames = []
    if "2025c" in sheet_names:
        frames.append(standardize_tab(xl.parse("2025c", dtype=str), "2025 Cont"))
    if "2026" in sheet_names:
        frames.append(standardize_tab(xl.parse("2026", dtype=str), "GAA 2026"))

    if not frames:
        return pd.DataFrame()

    df = pd.concat(frames, ignore_index=True)

    # Read Raw_Delivery if present so status updates automatically even without running Apps Script
    raw_lookup = {}
    if "Raw_Delivery" in sheet_names:
        raw_df = xl.parse("Raw_Delivery", dtype=str)
        if raw_df.shape[1] >= 7:
            site_codes = clean_series(raw_df.iloc[:, 5])
            cur_statuses = clean_series(raw_df.iloc[:, 6])
            for sid, st_val in zip(site_codes, cur_statuses):
                if sid and st_val:
                    raw_lookup[sid] = st_val

    # Map delivery status from Raw_Delivery or Column N
    raw_mapped = df["School_ID"].map(raw_lookup).fillna("")
    effective_raw_status = np.where(raw_mapped != "", raw_mapped, df["Col_N_Status"])
    eff_lower = pd.Series(effective_raw_status).str.lower().str.strip()
    orig_status_lower = df["Status"].str.lower().str.strip()

    conditions = [
        eff_lower.isin(["complete", "completed"]) | (orig_status_lower == "completed"),
        eff_lower.str.contains("pending mov", na=False) | (orig_status_lower == "pending mov") | (df["DR_Date"] != ""),
        orig_status_lower == "successful rebid",
    ]
    choices = ["Completed", "Delivered, Docs Pending", "Successful Rebid"]
    df["Delivery_Stage"] = np.select(conditions, choices, default="Ongoing Delivery")

    # Normalize Region & Package
    df["Region_Clean"] = df["Region"].str.upper()
    df["Package_Lower"] = df["DCP_Package"].str.lower()
    df["Package_Display"] = df["Package_Lower"].map(PACKAGE_DISPLAY).fillna(df["DCP_Package"])
    df["Component_Lower"] = df["Component"].str.lower()
    df["Component_Display"] = df["Component_Lower"].map(COMPONENT_DISPLAY).fillna(df["Component"])

    # Recipient Key (Region|Office for Laptop for Non-Teaching, School ID for others)
    is_nonteaching = df["Package_Lower"] == "laptop for non-teaching"
    office_key = df["Region_Clean"] + "|" + df["Recipient_Office"]
    df["Recipient_Key"] = np.where(
        is_nonteaching,
        np.where(df["Recipient_Office"] != "", office_key, np.where(df["School_ID"] != "", df["School_ID"], df["LOT"])),
        np.where(df["School_ID"] != "", df["School_ID"], np.where(df["Recipient_Office"] != "", office_key, df["LOT"]))
    )

    df = df[df["Recipient_Key"] != ""].copy()

    # Numeric columns
    df["Pkg_Count"] = pd.to_numeric(df["No_of_Package"].str.replace("[^0-9.-]", "", regex=True), errors="coerce").fillna(0)
    df["Units_Per_Pkg"] = pd.to_numeric(df["Units_Per_Package"].str.replace("[^0-9.-]", "", regex=True), errors="coerce").fillna(0)
    df["Unit_Cost"] = pd.to_numeric(df["Unit_Cost"].str.replace("[^0-9.-]", "", regex=True), errors="coerce").fillna(0)

    return df


def compute_remaining_packages(sub_df: pd.DataFrame, group_cols: list) -> pd.Series:
    undeliv = sub_df[sub_df["Delivery_Stage"].isin(["Ongoing Delivery", "Successful Rebid"])].copy()
    if undeliv.empty:
        return pd.Series(dtype=float)

    is_cart = undeliv["Package_Lower"].str.startswith("elearning cart package")
    non_cart = undeliv[~is_cart].groupby(group_cols)["Pkg_Count"].sum()

    cart_dedup_cols = list(dict.fromkeys(group_cols + ["Fund_Clean", "LOT", "Recipient_Key", "Package_Lower"]))
    cart = (
        undeliv[is_cart]
        .groupby(cart_dedup_cols)["Pkg_Count"]
        .max()
        .groupby(group_cols)
        .sum()
    )
    return non_cart.add(cart, fill_value=0)


# ==========================================
# SIDEBAR & DATA CONNECTION
# ==========================================
with st.sidebar:
    st.markdown("### ⚙️ Data Source & Filters")
    source_mode = st.radio(
        "Source Mode",
        ["Live Google Sheet", "Upload Excel (.xlsx)"],
        index=0,
    )

    excel_data = None
    if source_mode == "Live Google Sheet":
        if st.button("🔄 Refresh Live Data", use_container_width=True):
            st.cache_data.clear()
        try:
            with st.spinner("Loading Google Sheets..."):
                excel_data = load_workbook_from_gsheet(SHEET_ID)
        except Exception as e:
            st.error(
                "Hindi ma-access ang Google Sheet. Siguraduhing naka-set sa "
                "'Anyone with the link > Viewer' ang Share settings, o gamitin ang 'Upload Excel (.xlsx)'."
            )
            st.caption(f"Details: {e}")
    else:
        uploaded = st.file_uploader("Upload DCP Monitoring Excel", type=["xlsx"])
        if uploaded is not None:
            excel_data = uploaded.read()

if excel_data is None:
    st.info("👈 Pumili ng data source sa kaliwa para mag-load ang dashboard.")
    st.stop()

df = process_data(excel_data)
if df.empty:
    st.warning("Walang nakitang valid records sa '2025c' o '2026' tabs.")
    st.stop()

# Sidebar Filters
with st.sidebar:
    st.markdown("---")
    fund_options = ["All"] + sorted(df["Fund_Clean"].unique().tolist())
    sel_fund = st.selectbox("Fund Source", fund_options)

    region_options = ["All"] + sorted(
        df["Region_Clean"].unique().tolist(),
        key=lambda r: REGION_ORDER.get(r, 999),
    )
    sel_region = st.selectbox("Region", region_options)

    pkg_options = ["All"] + sorted(df["Package_Display"].unique().tolist())
    sel_package = st.selectbox("DCP Package", pkg_options)

# Apply Filters
filtered = df.copy()
if sel_fund != "All":
    filtered = filtered[filtered["Fund_Clean"] == sel_fund]
if sel_region != "All":
    filtered = filtered[filtered["Region_Clean"] == sel_region]
if sel_package != "All":
    filtered = filtered[filtered["Package_Display"] == sel_package]

# ==========================================
# HEADER BANNER
# ==========================================
st.markdown(
    f"""
    <div class="header-banner">
        <h1>📦 DCP Delivery Monitoring Dashboard</h1>
        <p>Fund Source: <b>{sel_fund}</b> &nbsp;|&nbsp; Region: <b>{sel_region}</b> &nbsp;|&nbsp; Package: <b>{sel_package}</b></p>
    </div>
    """,
    unsafe_allow_html=True,
)

if filtered.empty:
    st.warning("Walang records na tumugma sa napiling filters.")
    st.stop()

# ==========================================
# EXECUTIVE KPI CALCULATIONS
# ==========================================
# 1 School = 1 Assigned Region (matches Regional sum to Overall total)
unique_rec = filtered.drop_duplicates(subset=["Fund_Clean", "Recipient_Key"], keep="last").copy()

# Determine highest priority stage per unique recipient
stage_priority = {"Completed": 4, "Delivered, Docs Pending": 3, "Ongoing Delivery": 2, "Successful Rebid": 1}
filtered["Stage_Rank"] = filtered["Delivery_Stage"].map(stage_priority).fillna(0)
rec_stage = (
    filtered.sort_values("Stage_Rank")
    .drop_duplicates(subset=["Fund_Clean", "Recipient_Key"], keep="last")
)

total_recipients = len(unique_rec)
total_rem_pkgs = int(compute_remaining_packages(filtered, ["Fund_Clean"]).sum())
cnt_ongoing = int((rec_stage["Delivery_Stage"] == "Ongoing Delivery").sum())
cnt_rebid = int((rec_stage["Delivery_Stage"] == "Successful Rebid").sum())
cnt_pending_docs = int((rec_stage["Delivery_Stage"] == "Delivered, Docs Pending").sum())
cnt_completed = int((rec_stage["Delivery_Stage"] == "Completed").sum())

k1, k2, k3, k4, k5, k6 = st.columns(6)
kpis = [
    (k1, "🏢 Recipients", f"{total_recipients:,}", "Unique schools/offices", "#1d4ed8"),
    (k2, "📦 Remaining Pkgs", f"{total_rem_pkgs:,}", "For delivery", "#0284c7"),
    (k3, "🚚 Ongoing Delivery", f"{cnt_ongoing:,}", "Awarded / in transit", "#d97706"),
    (k4, "🔁 Successful Rebid", f"{cnt_rebid:,}", "Rebid allocation", "#7c3aed"),
    (k5, "📄 Docs Pending", f"{cnt_pending_docs:,}", "Delivered / Pending MOV", "#ea580c"),
    (k6, "✅ Completed", f"{cnt_completed:,}", "100% Completed", "#16a34a"),
]

for col, title, val, sub, color in kpis:
    with col:
        st.markdown(
            f"""
            <div class="kpi-card" style="border-left-color: {color};">
                <div class="kpi-title">{title}</div>
                <div class="kpi-value">{val}</div>
                <div class="kpi-sub">{sub}</div>
            </div>
            """,
            unsafe_allow_html=True,
        )

st.markdown("<br>", unsafe_allow_html=True)

# ==========================================
# 1. REGIONAL DELIVERY MONITORING TABLE
# ==========================================
regions_present = sorted(
    unique_rec["Region_Clean"].unique().tolist(),
    key=lambda r: REGION_ORDER.get(r, 999),
)

rem_by_reg = compute_remaining_packages(filtered, ["Region_Clean"])

reg_rows = []
for reg in regions_present:
    r_sub = rec_stage[rec_stage["Region_Clean"] == reg]
    reg_rows.append({
        "Region": reg,
        "Recipient Schools": int((unique_rec["Region_Clean"] == reg).sum()),
        "Remaining Packages": int(rem_by_reg.get(reg, 0)),
        "Ongoing Delivery": int((r_sub["Delivery_Stage"] == "Ongoing Delivery").sum()),
        "Successful Rebid": int((r_sub["Delivery_Stage"] == "Successful Rebid").sum()),
        "Delivered, Docs Pending": int((r_sub["Delivery_Stage"] == "Delivered, Docs Pending").sum()),
        "Completed": int((r_sub["Delivery_Stage"] == "Completed").sum()),
    })

reg_df = pd.DataFrame(reg_rows)

# Interactive Chart + Table Layout
chart_col1, chart_col2 = st.columns([1.3, 1])

with chart_col1:
    st.subheader("📍 Delivery Progress by Region")
    chart_melt = reg_df.melt(
        id_vars="Region",
        value_vars=["Completed", "Delivered, Docs Pending", "Ongoing Delivery", "Successful Rebid"],
        var_name="Stage",
        value_name="Recipients",
    )
    fig_reg = px.bar(
        chart_melt,
        x="Region",
        y="Recipients",
        color="Stage",
        barmode="stack",
        color_discrete_map={
            "Completed": "#16a34a",
            "Delivered, Docs Pending": "#ea580c",
            "Ongoing Delivery": "#f59e0b",
            "Successful Rebid": "#8b5cf6",
        },
        height=370,
    )
    fig_reg.update_layout(margin=dict(l=10, r=10, t=20, b=10), legend_title_text="")
    st.plotly_chart(fig_reg, use_container_width=True)

with chart_col2:
    st.subheader("📊 Overall Stage Share")
    pie_df = pd.DataFrame({
        "Stage": ["Completed", "Delivered, Docs Pending", "Ongoing Delivery", "Successful Rebid"],
        "Count": [cnt_completed, cnt_pending_docs, cnt_ongoing, cnt_rebid],
    })
    pie_df = pie_df[pie_df["Count"] > 0]
    fig_pie = px.pie(
        pie_df,
        names="Stage",
        values="Count",
        hole=0.52,
        color="Stage",
        color_discrete_map={
            "Completed": "#16a34a",
            "Delivered, Docs Pending": "#ea580c",
            "Ongoing Delivery": "#f59e0b",
            "Successful Rebid": "#8b5cf6",
        },
        height=370,
    )
    fig_pie.update_layout(margin=dict(l=10, r=10, t=20, b=10))
    st.plotly_chart(fig_pie, use_container_width=True)

st.subheader("🗺️ Regional Delivery Monitoring")
reg_total = pd.DataFrame([{
    "Region": "TOTAL",
    "Recipient Schools": reg_df["Recipient Schools"].sum(),
    "Remaining Packages": reg_df["Remaining Packages"].sum(),
    "Ongoing Delivery": reg_df["Ongoing Delivery"].sum(),
    "Successful Rebid": reg_df["Successful Rebid"].sum(),
    "Delivered, Docs Pending": reg_df["Delivered, Docs Pending"].sum(),
    "Completed": reg_df["Completed"].sum(),
}])
reg_display = pd.concat([reg_df, reg_total], ignore_index=True)
st.dataframe(
    reg_display.style.format({
        "Recipient Schools": "{:,.0f}",
        "Remaining Packages": "{:,.0f}",
        "Ongoing Delivery": "{:,.0f}",
        "Successful Rebid": "{:,.0f}",
        "Delivered, Docs Pending": "{:,.0f}",
        "Completed": "{:,.0f}",
    }),
    use_container_width=True,
    hide_index=True,
)

# ==========================================
# 2. DCP PACKAGE DISTRIBUTION TABLE
# ==========================================
st.subheader("📦 DCP Package Distribution")
rem_by_pkg = compute_remaining_packages(filtered, ["Package_Display"])
pkg_rows = []
for pkg in sorted(filtered["Package_Display"].unique()):
    p_sub = filtered[filtered["Package_Display"] == pkg]
    p_rec = p_sub.drop_duplicates(subset=["Fund_Clean", "Recipient_Key"])
    p_undeliv_devices = p_sub[
        p_sub["Delivery_Stage"].isin(["Ongoing Delivery", "Successful Rebid"])
        & (~p_sub["Component_Lower"].str.contains("storage|charging cart", na=False))
    ]
    rem_units = int((p_undeliv_devices["Pkg_Count"] * p_undeliv_devices["Units_Per_Pkg"]).sum())

    pkg_rows.append({
        "DCP Package": pkg,
        "Recipients": len(p_rec),
        "Remaining Packages": int(rem_by_pkg.get(pkg, 0)),
        "Remaining Device Units": rem_units,
    })

pkg_df = pd.DataFrame(pkg_rows)
pkg_total = pd.DataFrame([{
    "DCP Package": "TOTAL",
    "Recipients": pkg_df["Recipients"].sum(),
    "Remaining Packages": pkg_df["Remaining Packages"].sum(),
    "Remaining Device Units": pkg_df["Remaining Device Units"].sum(),
}])
st.dataframe(
    pd.concat([pkg_df, pkg_total], ignore_index=True).style.format({
        "Recipients": "{:,.0f}",
        "Remaining Packages": "{:,.0f}",
        "Remaining Device Units": "{:,.0f}",
    }),
    use_container_width=True,
    hide_index=True,
)

# ==========================================
# 3. COMPONENT BREAKDOWN TABLE
# ==========================================
st.subheader("💻 Component Breakdown")
comp_valid = filtered[filtered["Component_Display"] != ""]
comp_rows = []
for comp in sorted(comp_valid["Component_Display"].unique()):
    c_sub = comp_valid[comp_valid["Component_Display"] == comp]
    c_rec = c_sub.drop_duplicates(subset=["Fund_Clean", "Recipient_Key"])
    c_undeliv = c_sub[c_sub["Delivery_Stage"].isin(["Ongoing Delivery", "Successful Rebid"])]

    units = (c_undeliv["Pkg_Count"] * c_undeliv["Units_Per_Pkg"]).sum()
    cost = (c_undeliv["Pkg_Count"] * c_undeliv["Units_Per_Pkg"] * c_undeliv["Unit_Cost"]).sum()
    avg_cost = (cost / units) if units > 0 else 0.0

    comp_rows.append({
        "Component": comp,
        "Recipients": len(c_rec),
        "Remaining Units": int(units),
        "Avg Unit Cost": avg_cost,
        "Remaining Cost": cost,
    })

comp_df = pd.DataFrame(comp_rows)
if not comp_df.empty:
    comp_total = pd.DataFrame([{
        "Component": "TOTAL",
        "Recipients": comp_df["Recipients"].sum(),
        "Remaining Units": comp_df["Remaining Units"].sum(),
        "Avg Unit Cost": (comp_df["Remaining Cost"].sum() / comp_df["Remaining Units"].sum()) if comp_df["Remaining Units"].sum() > 0 else 0.0,
        "Remaining Cost": comp_df["Remaining Cost"].sum(),
    }])
    st.dataframe(
        pd.concat([comp_df, comp_total], ignore_index=True).style.format({
            "Recipients": "{:,.0f}",
            "Remaining Units": "{:,.0f}",
            "Avg Unit Cost": "₱{:,.2f}",
            "Remaining Cost": "₱{:,.2f}",
        }),
        use_container_width=True,
        hide_index=True,
    )

# ==========================================
# EXPORT FORMATTED EXCEL REPORT
# ==========================================
st.markdown("---")
output = io.BytesIO()
with pd.ExcelWriter(output, engine="openpyxl") as writer:
    reg_display.to_excel(writer, sheet_name="Regional_Summary", index=False)
    pkg_df.to_excel(writer, sheet_name="Package_Distribution", index=False)
    comp_df.to_excel(writer, sheet_name="Component_Breakdown", index=False)

st.download_button(
    label="📥 Download Executive Summary Report (.xlsx)",
    data=output.getvalue(),
    file_name=f"DCP_Monitoring_Report_{sel_fund}_{sel_region}.xlsx".replace(" ", "_"),
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)