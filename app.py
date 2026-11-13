import io
import requests
import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st
from openpyxl.styles import Font, PatternFill, Alignment

# ==========================================
# PAGE CONFIG & MODERN EXECUTIVE STYLING
# ==========================================
st.set_page_config(
    page_title="DCP Delivery Monitoring Portal",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown(
    """
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800&display=swap');

    html, body, [class*="css"] {
        font-family: 'Plus Jakarta Sans', sans-serif;
    }

    .stApp {
        background-color: #f8fafc;
    }

    .block-container {
        padding-top: 1.3rem;
        padding-bottom: 2.5rem;
        max-width: 1440px;
    }

    /* Executive Hero Banner */
    .hero-banner {
        background: linear-gradient(120deg, #0f172a 0%, #1e3a8a 55%, #2563eb 100%);
        padding: 1.6rem 2rem;
        border-radius: 16px;
        color: #ffffff;
        margin-bottom: 1.4rem;
        box-shadow: 0 10px 25px -5px rgba(30, 58, 138, 0.25);
        display: flex;
        justify-content: space-between;
        align-items: center;
        flex-wrap: wrap;
        gap: 1rem;
    }
    .hero-title {
        margin: 0;
        font-size: 1.75rem;
        font-weight: 800;
        letter-spacing: -0.02em;
        color: #ffffff;
    }
    .hero-subtitle {
        margin: 0.35rem 0 0 0;
        font-size: 0.9rem;
        color: #bfdbfe;
        font-weight: 500;
    }
    .hero-badge {
        background: rgba(255, 255, 255, 0.14);
        border: 1px solid rgba(255, 255, 255, 0.25);
        backdrop-filter: blur(6px);
        padding: 0.55rem 1rem;
        border-radius: 999px;
        font-size: 0.82rem;
        font-weight: 600;
        color: #f8fafc;
    }

    /* Modern KPI Cards */
    .kpi-card {
        background: #ffffff;
        border: 1px solid #e2e8f0;
        border-radius: 14px;
        padding: 1.1rem 1.2rem;
        box-shadow: 0 2px 10px rgba(15, 23, 42, 0.03);
        transition: transform 0.15s ease, box-shadow 0.15s ease;
        height: 100%;
    }
    .kpi-card:hover {
        transform: translateY(-2px);
        box-shadow: 0 8px 20px rgba(15, 23, 42, 0.07);
    }
    .kpi-top {
        display: flex;
        justify-content: space-between;
        align-items: center;
        margin-bottom: 0.45rem;
    }
    .kpi-label {
        font-size: 0.76rem;
        font-weight: 700;
        text-transform: uppercase;
        letter-spacing: 0.04em;
        color: #64748b;
    }
    .kpi-pill {
        font-size: 0.72rem;
        font-weight: 700;
        padding: 0.18rem 0.55rem;
        border-radius: 999px;
    }
    .kpi-number {
        font-size: 1.75rem;
        font-weight: 800;
        color: #0f172a;
        line-height: 1.15;
        letter-spacing: -0.02em;
    }
    .kpi-foot {
        font-size: 0.76rem;
        color: #94a3b8;
        margin-top: 0.35rem;
    }

    /* Section Card Wrapper */
    .section-header {
        font-size: 1.08rem;
        font-weight: 700;
        color: #0f172a;
        margin: 0.4rem 0 0.75rem 0;
        display: flex;
        align-items: center;
        gap: 0.45rem;
    }

    /* Tabs Styling */
    .stTabs [data-baseweb="tab-list"] {
        gap: 8px;
        background-color: #e2e8f0;
        padding: 5px;
        border-radius: 12px;
    }
    .stTabs [data-baseweb="tab"] {
        border-radius: 8px;
        padding: 8px 18px;
        font-weight: 600;
        font-size: 0.88rem;
        color: #475569;
    }
    .stTabs [aria-selected="true"] {
        background-color: #ffffff !important;
        color: #1e3a8a !important;
        box-shadow: 0 2px 6px rgba(15, 23, 42, 0.08);
    }
    </style>
    """,
    unsafe_allow_html=True,
)
SHEET_ID = "1uGsxNYJHhL0DOfODQbn60L2wXnC2qd-lKO56O8OzF4g"

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

STAGE_COLORS = {
    "Completed": "#10b981",
    "Delivered, Docs Pending": "#f97316",
    "Ongoing Delivery": "#3b82f6",
    "Successful Rebid": "#8b5cf6",
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

    # Auto-read Raw_Delivery tab if present
    raw_lookup = {}
    if "Raw_Delivery" in sheet_names:
        raw_df = xl.parse("Raw_Delivery", dtype=str)
        if raw_df.shape[1] >= 7:
            site_codes = clean_series(raw_df.iloc[:, 5])
            cur_statuses = clean_series(raw_df.iloc[:, 6])
            for sid, st_val in zip(site_codes, cur_statuses):
                if sid and st_val:
                    raw_lookup[sid] = st_val

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


def build_styled_excel(reg_display: pd.DataFrame, pkg_display: pd.DataFrame, comp_display: pd.DataFrame) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        reg_display.to_excel(writer, sheet_name="Regional_Summary", index=False)
        pkg_display.to_excel(writer, sheet_name="Package_Distribution", index=False)
        if not comp_display.empty:
            comp_display.to_excel(writer, sheet_name="Component_Breakdown", index=False)

        header_fill = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")
        header_font = Font(color="FFFFFF", bold=True, size=11)
        total_fill = PatternFill(start_color="EFF6FF", end_color="EFF6FF", fill_type="solid")
        total_font = Font(bold=True, size=11)

        for sheet_name in writer.sheets:
            ws = writer.sheets[sheet_name]
            for cell in ws[1]:
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = Alignment(horizontal="center", vertical="center")

            last_row = ws.max_row
            if last_row > 1:
                for cell in ws[last_row]:
                    cell.fill = total_fill
                    cell.font = total_font

            for col in ws.columns:
                max_len = max((len(str(cell.value or "")) for cell in col), default=10)
                ws.column_dimensions[col[0].column_letter].width = min(max_len + 5, 38)

    return output.getvalue()


# ==========================================
# SIDEBAR & FILTERS
# ==========================================
with st.sidebar:
    st.markdown("### 🎛️ Control Panel")
    source_mode = st.radio(
        "Data Connection",
        ["Live Google Sheet", "Upload Excel (.xlsx)"],
        index=0,
    )

    excel_data = None
    if source_mode == "Live Google Sheet":
        if st.button("🔄 Sync Latest Data", use_container_width=True, type="primary"):
            st.cache_data.clear()
        try:
            with st.spinner("Syncing live database..."):
                excel_data = load_workbook_from_gsheet(SHEET_ID)
        except Exception as e:
            st.error("Unable to fetch Google Sheet. Check 'Anyone with the link > Viewer' permission.")
            st.caption(f"Details: {e}")
    else:
        uploaded = st.file_uploader("Upload DCP Monitoring File", type=["xlsx"])
        if uploaded is not None:
            excel_data = uploaded.read()

if excel_data is None:
    st.info("👈 Select a data source on the left sidebar to load the dashboard.")
    st.stop()

df = process_data(excel_data)
if df.empty:
    st.warning("No valid records found in '2025c' or '2026'.")
    st.stop()

with st.sidebar:
    st.markdown("---")
    st.markdown("### 🔍 Report Filters")
    fund_options = ["All"] + sorted(df["Fund_Clean"].unique().tolist())
    sel_fund = st.selectbox("Fund Source", fund_options)

    region_options = ["All"] + sorted(
        df["Region_Clean"].unique().tolist(),
        key=lambda r: REGION_ORDER.get(r, 999),
    )
    sel_region = st.selectbox("Region", region_options)

    pkg_options = ["All"] + sorted(df["Package_Display"].unique().tolist())
    sel_package = st.selectbox("DCP Package", pkg_options)

    stage_options = ["All", "Completed", "Delivered, Docs Pending", "Ongoing Delivery", "Successful Rebid"]
    sel_stage = st.selectbox("Delivery Status", stage_options)

# Apply Filters
filtered = df.copy()
if sel_fund != "All":
    filtered = filtered[filtered["Fund_Clean"] == sel_fund]
if sel_region != "All":
    filtered = filtered[filtered["Region_Clean"] == sel_region]
if sel_package != "All":
    filtered = filtered[filtered["Package_Display"] == sel_package]
if sel_stage != "All":
    filtered = filtered[filtered["Delivery_Stage"] == sel_stage]

# ==========================================
# EXECUTIVE HERO BANNER
# ==========================================
st.markdown(
    f"""
    <div class="hero-banner">
        <div>
            <div class="hero-title">DepEd DCP Delivery Monitoring Portal</div>
            <div class="hero-subtitle">
                Fund: <b>{sel_fund}</b> &nbsp;•&nbsp; Region: <b>{sel_region}</b> &nbsp;•&nbsp; Package: <b>{sel_package}</b>
            </div>
        </div>
        <div class="hero-badge">● Live Executive Report</div>
    </div>
    """,
    unsafe_allow_html=True,
)

if filtered.empty:
    st.warning("No records match the selected filters.")
    st.stop()

# ==========================================
# KPI COMPUTATIONS (1 SCHOOL = 1 REGION)
# ==========================================
unique_rec = filtered.drop_duplicates(subset=["Fund_Clean", "Recipient_Key"], keep="last").copy()
assigned_region_map = unique_rec.set_index(["Fund_Clean", "Recipient_Key"])["Region_Clean"]

stage_priority = {"Completed": 4, "Delivered, Docs Pending": 3, "Ongoing Delivery": 2, "Successful Rebid": 1}
filtered["Stage_Rank"] = filtered["Delivery_Stage"].map(stage_priority).fillna(0)
rec_stage = (
    filtered.sort_values("Stage_Rank")
    .drop_duplicates(subset=["Fund_Clean", "Recipient_Key"], keep="last")
    .copy()
)
rec_stage["Region_Clean"] = rec_stage.set_index(["Fund_Clean", "Recipient_Key"]).index.map(assigned_region_map)

total_recipients = len(unique_rec)
total_rem_pkgs = int(compute_remaining_packages(filtered, ["Fund_Clean"]).sum())
cnt_ongoing = int((rec_stage["Delivery_Stage"] == "Ongoing Delivery").sum())
cnt_rebid = int((rec_stage["Delivery_Stage"] == "Successful Rebid").sum())
cnt_pending_docs = int((rec_stage["Delivery_Stage"] == "Delivered, Docs Pending").sum())
cnt_completed = int((rec_stage["Delivery_Stage"] == "Completed").sum())
cnt_delivered_total = cnt_completed + cnt_pending_docs

pct_delivered = (cnt_delivered_total / total_recipients * 100) if total_recipients > 0 else 0.0
pct_completed = (cnt_completed / total_recipients * 100) if total_recipients > 0 else 0.0
pct_ongoing = (cnt_ongoing / total_recipients * 100) if total_recipients > 0 else 0.0

# Render 6 Modern KPI Cards
cols = st.columns(6)
kpi_specs = [
    ("Total Recipients", f"{total_recipients:,}", "100%", "#eff6ff", "#1d4ed8", "Unique schools & offices"),
    ("Remaining Pkgs", f"{total_rem_pkgs:,}", "Pending", "#f0f9ff", "#0284c7", "Undelivered packages"),
    ("Ongoing Delivery", f"{cnt_ongoing:,}", f"{pct_ongoing:.1f}%", "#eff6ff", "#2563eb", "Awarded / in transit"),
    ("Successful Rebid", f"{cnt_rebid:,}", "Rebid", "#f5f3ff", "#7c3aed", "Rebid allocation"),
    ("Docs Pending", f"{cnt_pending_docs:,}", f"{pct_delivered:.1f}% Deliv.", "#fff7ed", "#ea580c", "Delivered / Pending MOV"),
    ("Fully Completed", f"{cnt_completed:,}", f"{pct_completed:.1f}%", "#ecfdf5", "#059669", "100% MOV & IAR done"),
]

for col, (label, val, badge, bg_badge, fg_badge, foot) in zip(cols, kpi_specs):
    with col:
        st.markdown(
            f"""
            <div class="kpi-card">
                <div class="kpi-top">
                    <span class="kpi-label">{label}</span>
                    <span class="kpi-pill" style="background:{bg_badge}; color:{fg_badge};">{badge}</span>
                </div>
                <div class="kpi-number">{val}</div>
                <div class="kpi-foot">{foot}</div>
            </div>
            """,
            unsafe_allow_html=True,
        )

st.markdown("<div style='height: 1rem;'></div>", unsafe_allow_html=True)

# ==========================================
# PREPARE SUMMARY DATAFRAMES
# ==========================================
regions_present = sorted(
    unique_rec["Region_Clean"].unique().tolist(),
    key=lambda r: REGION_ORDER.get(r, 999),
)
rem_by_reg = compute_remaining_packages(filtered, ["Region_Clean"])

reg_rows = []
for reg in regions_present:
    r_sub = rec_stage[rec_stage["Region_Clean"] == reg]
    tot_r = int((unique_rec["Region_Clean"] == reg).sum())
    comp_r = int((r_sub["Delivery_Stage"] == "Completed").sum())
    pend_r = int((r_sub["Delivery_Stage"] == "Delivered, Docs Pending").sum())
    comp_rate = ((comp_r + pend_r) / tot_r * 100) if tot_r > 0 else 0.0

    reg_rows.append({
        "Region": reg,
        "Recipient Schools": tot_r,
        "Remaining Packages": int(rem_by_reg.get(reg, 0)),
        "Ongoing Delivery": int((r_sub["Delivery_Stage"] == "Ongoing Delivery").sum()),
        "Successful Rebid": int((r_sub["Delivery_Stage"] == "Successful Rebid").sum()),
        "Delivered, Docs Pending": pend_r,
        "Completed": comp_r,
        "Delivery Rate (%)": round(comp_rate, 1),
    })

reg_df = pd.DataFrame(reg_rows)
reg_total = pd.DataFrame([{
    "Region": "TOTAL",
    "Recipient Schools": reg_df["Recipient Schools"].sum(),
    "Remaining Packages": reg_df["Remaining Packages"].sum(),
    "Ongoing Delivery": reg_df["Ongoing Delivery"].sum(),
    "Successful Rebid": reg_df["Successful Rebid"].sum(),
    "Delivered, Docs Pending": reg_df["Delivered, Docs Pending"].sum(),
    "Completed": reg_df["Completed"].sum(),
    "Delivery Rate (%)": round(pct_delivered, 1),
}])
reg_display = pd.concat([reg_df, reg_total], ignore_index=True)

# Package Distribution DF
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
pkg_display = pd.concat([pkg_df, pkg_total], ignore_index=True)

# Component Breakdown DF
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
    comp_display = pd.concat([comp_df, comp_total], ignore_index=True)
else:
    comp_display = pd.DataFrame()

# ==========================================
# MODERN INTERACTIVE TABS
# ==========================================
tab1, tab2, tab3, tab4 = st.tabs([
    "📊 Executive Visuals",
    "🗺️ Regional Monitoring Table",
    "📦 Packages & Components",
    "🔎 School / Office Lookup",
])

with tab1:
    c_left, c_right = st.columns([1.45, 1])

    with c_left:
        st.markdown('<div class="section-header">📍 Regional Delivery Volume by Stage</div>', unsafe_allow_html=True)
        chart_melt = reg_df.melt(
            id_vars="Region",
            value_vars=["Completed", "Delivered, Docs Pending", "Ongoing Delivery", "Successful Rebid"],
            var_name="Stage",
            value_name="Recipients",
        )
        fig_bar = px.bar(
            chart_melt,
            x="Region",
            y="Recipients",
            color="Stage",
            barmode="stack",
            color_discrete_map=STAGE_COLORS,
            height=390,
        )
        fig_bar.update_layout(
            plot_bgcolor="rgba(0,0,0,0)",
            paper_bgcolor="rgba(0,0,0,0)",
            margin=dict(l=10, r=10, t=15, b=10),
            legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="left", x=0, title_text=""),
            xaxis=dict(title="", showgrid=False),
            yaxis=dict(title="Recipients", gridcolor="#e2e8f0"),
            hovermode="x unified",
        )
        st.plotly_chart(fig_bar, use_container_width=True)

    with c_right:
        st.markdown('<div class="section-header">🍩 Overall Delivery Status Share</div>', unsafe_allow_html=True)
        pie_df = pd.DataFrame({
            "Stage": ["Completed", "Delivered, Docs Pending", "Ongoing Delivery", "Successful Rebid"],
            "Count": [cnt_completed, cnt_pending_docs, cnt_ongoing, cnt_rebid],
        })
        pie_df = pie_df[pie_df["Count"] > 0]
        fig_donut = go.Figure(
            data=[
                go.Pie(
                    labels=pie_df["Stage"],
                    values=pie_df["Count"],
                    hole=0.62,
                    marker=dict(colors=[STAGE_COLORS[s] for s in pie_df["Stage"]], line=dict(color="#ffffff", width=2)),
                    textinfo="percent",
                    textposition="outside",
                    hovertemplate="<b>%{label}</b><br>Recipients: %{value:,}<br>Share: %{percent}<extra></extra>",
                )
            ]
        )
        fig_donut.update_layout(
            height=390,
            margin=dict(l=10, r=10, t=25, b=10),
            showlegend=True,
            legend=dict(orientation="h", yanchor="bottom", y=-0.15, xanchor="center", x=0.5),
            annotations=[
                dict(
                    text=f"<b>{total_recipients:,}</b><br><span style='font-size:12px;color:#64748b'>Recipients</span>",
                    x=0.5, y=0.5, font_size=20, showarrow=False, font_color="#0f172a"
                )
            ],
        )
        st.plotly_chart(fig_donut, use_container_width=True)

    st.markdown("<hr style='border:0; border-top:1px solid #e2e8f0; margin: 1rem 0;'>", unsafe_allow_html=True)

    c_bot1, c_bot2 = st.columns([1.2, 1])

    with c_bot1:
        st.markdown('<div class="section-header">📈 Delivery Completion Rate (%) per Region</div>', unsafe_allow_html=True)
        rate_sorted = reg_df.sort_values("Delivery Rate (%)", ascending=True)
        fig_rate = px.bar(
            rate_sorted,
            x="Delivery Rate (%)",
            y="Region",
            orientation="h",
            text="Delivery Rate (%)",
            color="Delivery Rate (%)",
            color_continuous_scale=["#93c5fd", "#2563eb", "#10b981"],
            range_x=[0, 112],
            height=max(360, len(rate_sorted) * 26),
        )
        fig_rate.update_traces(texttemplate="%{text:.1f}%", textposition="outside")
        fig_rate.update_layout(
            plot_bgcolor="rgba(0,0,0,0)",
            paper_bgcolor="rgba(0,0,0,0)",
            coloraxis_showscale=False,
            margin=dict(l=10, r=20, t=10, b=10),
            xaxis=dict(title="Delivered + Completed (%)", gridcolor="#e2e8f0"),
            yaxis=dict(title=""),
        )
        st.plotly_chart(fig_rate, use_container_width=True)

    with c_bot2:
        st.markdown('<div class="section-header">📦 Remaining Packages by DCP Package</div>', unsafe_allow_html=True)
        pkg_sorted = pkg_df.sort_values("Remaining Packages", ascending=True)
        fig_pkg = px.bar(
            pkg_sorted,
            x="Remaining Packages",
            y="DCP Package",
            orientation="h",
            text="Remaining Packages",
            color_discrete_sequence=["#1d4ed8"],
            height=max(360, len(rate_sorted) * 26),
        )
        fig_pkg.update_traces(texttemplate="%{text:,}", textposition="outside")
        fig_pkg.update_layout(
            plot_bgcolor="rgba(0,0,0,0)",
            paper_bgcolor="rgba(0,0,0,0)",
            margin=dict(l=10, r=40, t=10, b=10),
            xaxis=dict(title="Remaining Packages", gridcolor="#e2e8f0"),
            yaxis=dict(title=""),
        )
        st.plotly_chart(fig_pkg, use_container_width=True)

with tab2:
    st.markdown('<div class="section-header">🗺️ Regional Delivery Monitoring Breakdown</div>', unsafe_allow_html=True)
    st.dataframe(
        reg_display.style.format({
            "Recipient Schools": "{:,.0f}",
            "Remaining Packages": "{:,.0f}",
            "Ongoing Delivery": "{:,.0f}",
            "Successful Rebid": "{:,.0f}",
            "Delivered, Docs Pending": "{:,.0f}",
            "Completed": "{:,.0f}",
            "Delivery Rate (%)": "{:.1f}%",
        }),
        use_container_width=True,
        hide_index=True,
        height=680,
    )

with tab3:
    st.markdown('<div class="section-header">📦 DCP Package Distribution</div>', unsafe_allow_html=True)
    st.dataframe(
        pkg_display.style.format({
            "Recipients": "{:,.0f}",
            "Remaining Packages": "{:,.0f}",
            "Remaining Device Units": "{:,.0f}",
        }),
        use_container_width=True,
        hide_index=True,
    )

    if not comp_display.empty:
        st.markdown("<br>", unsafe_allow_html=True)
        st.markdown('<div class="section-header">💻 Component Breakdown & Valuation</div>', unsafe_allow_html=True)
        st.dataframe(
            comp_display.style.format({
                "Recipients": "{:,.0f}",
                "Remaining Units": "{:,.0f}",
                "Avg Unit Cost": "₱{:,.2f}",
                "Remaining Cost": "₱{:,.2f}",
            }),
            use_container_width=True,
            hide_index=True,
        )

with tab4:
    st.markdown('<div class="section-header">🔎 Search Recipient School ID, Office, or Supplier</div>', unsafe_allow_html=True)
    search_q = st.text_input("Type School ID, Recipient Office, LOT, or Supplier...", placeholder="e.g. 502655 or Columbia Technologies")
    lookup_cols = [
        "Fund_Clean", "LOT", "Region_Clean", "School_ID", "Recipient_Office",
        "Package_Display", "Component_Display", "Pkg_Count", "Supplier", "Delivery_Stage"
    ]
    lookup_df = filtered[lookup_cols].copy()
    lookup_df.columns = [
        "Fund", "LOT", "Region", "School ID", "Recipient Office",
        "DCP Package", "Component", "Packages", "Supplier", "Status"
    ]
    if search_q.strip():
        q = search_q.strip().lower()
        mask = lookup_df.apply(lambda col: col.astype(str).str.lower().str.contains(q, na=False)).any(axis=1)
        lookup_df = lookup_df[mask]

    st.caption(f"Showing {len(lookup_df):,} records")
    st.dataframe(lookup_df.head(1000), use_container_width=True, hide_index=True, height=520)

# ==========================================
# SIDEBAR DOWNLOAD REPORT BUTTON
# ==========================================
with st.sidebar:
    st.markdown("---")
    excel_bytes = build_styled_excel(reg_display, pkg_display, comp_display)
    st.download_button(
        label="📥 Export Executive Excel (.xlsx)",
        data=excel_bytes,
        file_name=f"DCP_Executive_Report_{sel_fund}_{sel_region}.xlsx".replace(" ", "_"),
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )
