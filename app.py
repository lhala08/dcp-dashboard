import io
import re
from concurrent.futures import ThreadPoolExecutor
from urllib.parse import quote

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import requests
import streamlit as st
from openpyxl.styles import Alignment, Font, PatternFill

# =========================================================
# PAGE CONFIG
# Status-based version: no actual DR Date or IAR Date needed.
# "Completed" means delivery and document validation are cleared
# in the source system and the recipient may proceed to billing.
# =========================================================
st.set_page_config(
    page_title="DCP Delivery Monitoring",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="collapsed",
)

st.markdown(
    """
    <style>
    :root { color-scheme: light !important; }
    html, body, [class*="css"] { font-family: Arial, sans-serif; color:#0f172a; }
    .stApp { background:#f8fafc !important; }
    .block-container { padding-top:1.1rem; padding-bottom:2.5rem; max-width:1480px; }
    .hero { background:linear-gradient(120deg,#0f172a,#1e3a8a 55%,#2563eb);
            color:white; padding:1.35rem 1.7rem; border-radius:16px; margin-bottom:1rem;
            box-shadow:0 10px 25px rgba(30,58,138,.22); }
    .hero h1 { font-size:1.65rem; margin:0; color:white; }
    .hero p { margin:.35rem 0 0; color:#bfdbfe; font-size:.88rem; }
    .kpi { background:white; border:1px solid #e2e8f0; border-radius:14px;
           padding:1rem; min-height:118px; box-shadow:0 2px 10px rgba(15,23,42,.04); }
    .kpi-label { color:#64748b; font-size:.72rem; font-weight:700; text-transform:uppercase; }
    .kpi-value { color:#0f172a; font-size:1.55rem; font-weight:800; margin-top:.35rem; }
    .kpi-foot { color:#94a3b8; font-size:.73rem; margin-top:.3rem; }
    .section { font-size:1.05rem; font-weight:700; margin:.5rem 0 .65rem; }
    </style>
    """,
    unsafe_allow_html=True,
)

# =========================================================
# SETTINGS / SECURITY
# =========================================================
DEFAULT_SHEET_ID = "1uGsxNYJHhL0DOfODQbn60L2wXnC2qd-lKO56O8OzF4g"
DEFAULT_APP_PASSWORD = "!dcp2026*"


def get_secret(key: str, default: str = "") -> str:
    try:
        value = st.secrets.get(key, default)
        return str(value) if value else default
    except Exception:
        return default


SHEET_ID = get_secret("SHEET_ID", DEFAULT_SHEET_ID)
APP_PASSWORD = get_secret("APP_PASSWORD", DEFAULT_APP_PASSWORD)

if "authenticated" not in st.session_state:
    st.session_state.authenticated = False

try:
    url_key = st.query_params.get("key", "")
except Exception:
    url_key = ""

if APP_PASSWORD and url_key == APP_PASSWORD:
    st.session_state.authenticated = True

if APP_PASSWORD and not st.session_state.authenticated:
    _, middle, _ = st.columns([1.2, 1.3, 1.2])
    with middle:
        st.markdown("## 🔐 DCP Monitoring Portal")
        st.caption("Restricted internal dashboard")
        with st.form("login"):
            pwd = st.text_input("Access Password", type="password")
            submit = st.form_submit_button("Unlock Dashboard", use_container_width=True, type="primary")
            if submit:
                if pwd == APP_PASSWORD:
                    st.session_state.authenticated = True
                    st.rerun()
                else:
                    st.error("Incorrect password.")
    st.stop()

REGION_ORDER = {
    "1":1,"2":2,"3":3,"4A":4,"4B":5,"MIMAROPA":6,"5":7,"6":8,"7":9,
    "8":10,"9":11,"10":12,"11":13,"12":14,"CAR":15,"CARAGA":16,
    "NCR":17,"NIR":18,"BARMM":19,
}

PACKAGE_DISPLAY = {
    "laptop for teaching":"Laptop for Teaching",
    "laptop for non-teaching":"Laptop for Non-Teaching",
    "smart tv package":"Smart TV Package",
    "elearning cart package 1":"eLearning Cart Package 1",
    "elearning cart package 2":"eLearning Cart Package 2",
    "elearning cart package 3":"eLearning Cart Package 3",
}

COMPONENT_DISPLAY = {
    "storage/charging cart":"Storage/Charging Cart",
    "storage cart":"Storage Cart",
    "charging cart":"Charging Cart",
    "smart tv":"Smart TV",
    "tablet":"Tablet",
    "laptop":"Laptop",
}

STAGES = ["Completed", "Delivered (Pending Validation)", "Ongoing Delivery", "Successful Rebid"]
STAGE_COLORS = {
    "Completed":"#10b981",
    "Delivered (Pending Validation)":"#f59e0b",
    "Ongoing Delivery":"#3b82f6",
    "Successful Rebid":"#8b5cf6",
}

# =========================================================
# DATA HELPERS
# =========================================================
def clean_series(series: pd.Series) -> pd.Series:
    return series.fillna("").astype(str).str.replace("\xa0", " ", regex=False).str.strip()


def find_col_by_keywords(columns, keywords) -> str:
    for col in columns:
        name = str(col).lower().strip()
        if any(k in name for k in keywords):
            return col
    return ""


def normalize_division_name(raw_div: str, region: str) -> tuple[str, bool]:
    value = re.sub(r"\s+", " ", str(raw_div or "")).strip()
    reg = str(region or "").upper().strip()
    if not value:
        return "UNSPECIFIED DIVISION", False

    upper = re.sub(r"\s+", " ", value.upper().replace(".", " ")).strip()
    ro_patterns = [
        r"^(RO|R\s*O)[\s\-]*(.+)$",
        r"^REGION[\s\-]*(.+)$",
        r"REGIONAL\s+OFFICE",
    ]
    if any(re.search(pattern, upper) for pattern in ro_patterns):
        return upper, False

    upper = re.sub(
        r"^(CITY\s+SCHOOLS\s+DIVISION\s+OF|SCHOOLS\s+DIVISION\s+OF|DIVISION\s+OF|SDO[\s\-]+)\s*",
        "", upper,
    ).strip()
    upper = re.sub(r"\bSTO\b", "SANTO", upper)
    upper = re.sub(r"\bSTA\b", "SANTA", upper)
    if upper.startswith("CITY OF "):
        upper = upper[8:].strip() + " CITY"
    if upper in {"SAN CARLOS CITY", "TALISAY CITY"} and reg:
        upper = f"{upper} - R{reg}"
    return upper or "UNSPECIFIED DIVISION", True


def raw_status_priority(status: str) -> int:
    low = str(status or "").lower().strip()
    if not low:
        return 0
    if any(x in low for x in ["complete", "completed", "validated", "accepted"]):
        return 4
    if any(x in low for x in ["delivered", "pending validation", "pending docs", "pending mov", "pending dr", "pending iar"]):
        return 3
    if re.match(r"^(rebid|for rebid|successful rebid)$", low):
        return 2
    return 1


def classify_status(status: str) -> str:
    """Classify exclusively from system status. Dates are intentionally ignored."""
    low = str(status or "").lower().strip()
    if any(x in low for x in ["complete", "completed", "validated", "accepted"]):
        return "Completed"
    if any(x in low for x in ["delivered", "pending validation", "pending docs", "pending mov", "pending dr", "pending iar"]):
        return "Delivered (Pending Validation)"
    if re.match(r"^(rebid|for rebid|successful rebid)$", low):
        return "Successful Rebid"
    return "Ongoing Delivery"


def standardize_tab(raw: pd.DataFrame, default_fund: str) -> pd.DataFrame:
    df = raw.copy()
    original_cols = list(df.columns)
    div_col = find_col_by_keywords(original_cols, ["division", "sdo"])
    explicit_div = clean_series(df[div_col]) if div_col else pd.Series([""] * len(df), index=df.index)

    while df.shape[1] < 16:
        df[f"extra_{df.shape[1]}"] = ""

    std = df.iloc[:, :16].copy()
    std.columns = [
        "LOT", "Region", "DCP_Package", "No_of_Package", "School_ID",
        "Recipient_Office", "Component", "Units_Per_Package", "Unit_Cost",
        "Supplier", "Legacy_DR_Date", "Legacy_IAR_Date", "IAR_Docs",
        "Col_N_Status", "Status", "Fund_Source",
    ]
    for col in std.columns:
        std[col] = clean_series(std[col])
    std["Explicit_Division"] = explicit_div.values
    std["Fund_Clean"] = default_fund
    return std[std["DCP_Package"] != ""].copy()


def fetch_tab(sheet_id: str, tab: str) -> pd.DataFrame:
    url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/gviz/tq?tqx=out:csv&sheet={quote(tab)}"
    response = requests.get(url, timeout=20)
    response.raise_for_status()
    return pd.read_csv(io.BytesIO(response.content), dtype=str)


@st.cache_data(ttl=1800, show_spinner=False)
def load_from_gsheet(sheet_id: str) -> dict[str, pd.DataFrame]:
    tabs = ["2025c", "2026", "Raw_Delivery"]
    try:
        with ThreadPoolExecutor(max_workers=3) as pool:
            futures = {tab: pool.submit(fetch_tab, sheet_id, tab) for tab in tabs}
            result = {tab: future.result() for tab, future in futures.items()}
        if "2025c" in result and "2026" in result and result["2025c"].equals(result["2026"]):
            del result["2026"]
        return result
    except Exception:
        url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=xlsx"
        response = requests.get(url, timeout=30)
        response.raise_for_status()
        xl = pd.ExcelFile(io.BytesIO(response.content), engine="openpyxl")
        return {tab: xl.parse(tab, dtype=str) for tab in tabs if tab in xl.sheet_names}


@st.cache_data(ttl=1800, show_spinner=False)
def load_from_excel(content: bytes) -> dict[str, pd.DataFrame]:
    xl = pd.ExcelFile(io.BytesIO(content), engine="openpyxl")
    return {tab: xl.parse(tab, dtype=str) for tab in ["2025c", "2026", "Raw_Delivery"] if tab in xl.sheet_names}


@st.cache_data(ttl=1800, show_spinner=False)
def build_processed_df(tables: dict[str, pd.DataFrame]) -> pd.DataFrame:
    frames = []
    if "2025c" in tables:
        frames.append(standardize_tab(tables["2025c"], "2025 Cont"))
    if "2026" in tables:
        frames.append(standardize_tab(tables["2026"], "GAA 2026"))
    if not frames:
        return pd.DataFrame()

    df = pd.concat(frames, ignore_index=True)
    raw_status_by_sid_lot, raw_status_by_sid, raw_div_lookup = {}, {}, {}

    if "Raw_Delivery" in tables:
        raw = tables["Raw_Delivery"].copy()
        cols = list(raw.columns)
        sid_col = find_col_by_keywords(cols, ["site code", "school id", "school_id", "beis"])
        status_col = find_col_by_keywords(cols, ["current status", "delivery status", "status"])
        lot_col = find_col_by_keywords(cols, ["lot"])
        div_col = find_col_by_keywords(cols, ["division", "sdo"])
        if not sid_col and raw.shape[1] >= 6:
            sid_col = cols[5]
        if not status_col and raw.shape[1] >= 7:
            status_col = cols[6]

        if sid_col:
            sids = clean_series(raw[sid_col])
            lots = clean_series(raw[lot_col]) if lot_col else pd.Series([""] * len(raw))
            statuses = clean_series(raw[status_col]) if status_col else pd.Series([""] * len(raw))
            divisions = clean_series(raw[div_col]) if div_col else pd.Series([""] * len(raw))
            for sid, lot, status, division in zip(sids, lots, statuses, divisions):
                if sid and status:
                    if raw_status_priority(status) >= raw_status_priority(raw_status_by_sid.get(sid, "")):
                        raw_status_by_sid[sid] = status
                    if lot and raw_status_priority(status) >= raw_status_priority(raw_status_by_sid_lot.get((sid, lot), "")):
                        raw_status_by_sid_lot[(sid, lot)] = status
                if sid and division:
                    raw_div_lookup[sid] = division

    effective = []
    for sid, lot, col_n, status in zip(df["School_ID"], df["LOT"], df["Col_N_Status"], df["Status"]):
        candidates = [raw_status_by_sid_lot.get((sid, lot), ""), raw_status_by_sid.get(sid, ""), col_n, status]
        effective.append(max(candidates, key=raw_status_priority))
    df["Effective_Status"] = effective
    df["Delivery_Stage"] = df["Effective_Status"].map(classify_status)

    df["Region_Clean"] = df["Region"].str.upper().str.strip()
    df["Package_Lower"] = df["DCP_Package"].str.lower().str.strip()
    df["Package_Display"] = df["Package_Lower"].map(PACKAGE_DISPLAY).fillna(df["DCP_Package"])
    df["Component_Lower"] = df["Component"].str.lower().str.strip()
    df["Component_Display"] = df["Component_Lower"].map(COMPONENT_DISPLAY).fillna(df["Component"])

    raw_div = df["School_ID"].map(raw_div_lookup).fillna("")
    resolved = np.where(
        df["Explicit_Division"] != "", df["Explicit_Division"],
        np.where(raw_div != "", raw_div, np.where(df["Recipient_Office"] != "", df["Recipient_Office"], "UNSPECIFIED DIVISION")),
    )
    normalized = [normalize_division_name(value, region) for value, region in zip(resolved, df["Region_Clean"])]
    df["Division_Clean"] = [value[0] for value in normalized]
    df["Is_SDO"] = [value[1] for value in normalized]

    nonteaching = df["Package_Lower"] == "laptop for non-teaching"
    office_key = df["Region_Clean"] + "|" + df["Division_Clean"]
    df["Recipient_Key"] = np.where(
        nonteaching,
        np.where(df["Division_Clean"] != "UNSPECIFIED DIVISION", office_key, np.where(df["School_ID"] != "", df["School_ID"], df["LOT"])),
        np.where(df["School_ID"] != "", df["School_ID"], np.where(df["Recipient_Office"] != "", office_key, df["LOT"])),
    )
    df = df[df["Recipient_Key"] != ""].copy()
    df["Row_ID"] = np.arange(len(df))
    df["Pkg_Count"] = pd.to_numeric(df["No_of_Package"].str.replace(r"[^0-9.-]", "", regex=True), errors="coerce").fillna(0)
    df["Units_Per_Pkg"] = pd.to_numeric(df["Units_Per_Package"].str.replace(r"[^0-9.-]", "", regex=True), errors="coerce").fillna(0)
    df["Unit_Cost_Num"] = pd.to_numeric(df["Unit_Cost"].str.replace(r"[^0-9.-]", "", regex=True), errors="coerce").fillna(0)
    return df


def compute_remaining_packages(df: pd.DataFrame, group_cols: list[str]) -> pd.Series:
    pending = df[df["Delivery_Stage"].isin(["Ongoing Delivery", "Successful Rebid"])].copy()
    if pending.empty:
        return pd.Series(dtype=float)
    is_cart = pending["Package_Lower"].str.startswith("elearning cart package")
    regular = pending[~is_cart].groupby(group_cols)["Pkg_Count"].sum()
    dedup_cols = list(dict.fromkeys(group_cols + ["Fund_Clean", "LOT", "Recipient_Key", "Package_Lower"]))
    carts = pending[is_cart].groupby(dedup_cols)["Pkg_Count"].max().groupby(group_cols).sum()
    return regular.add(carts, fill_value=0)


def build_excel(regional, divisions, packages, components) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        regional.to_excel(writer, sheet_name="Regional_Summary", index=False)
        divisions.to_excel(writer, sheet_name="Division_Billing_Status", index=False)
        packages.to_excel(writer, sheet_name="Package_Distribution", index=False)
        if not components.empty:
            components.to_excel(writer, sheet_name="Component_Breakdown", index=False)

        blue = PatternFill("solid", fgColor="1E3A8A")
        green = PatternFill("solid", fgColor="DCFCE7")
        yellow = PatternFill("solid", fgColor="FEF3C7")
        for ws in writer.sheets.values():
            for cell in ws[1]:
                cell.fill = blue
                cell.font = Font(color="FFFFFF", bold=True)
                cell.alignment = Alignment(horizontal="center")
            if ws.title == "Division_Billing_Status":
                for row in ws.iter_rows(min_row=2):
                    text = " ".join(str(c.value or "") for c in row)
                    fill = green if "READY FOR BILLING" in text else yellow if "PENDING VALIDATION" in text else None
                    if fill:
                        for cell in row:
                            cell.fill = fill
            for column in ws.columns:
                width = min(max(len(str(c.value or "")) for c in column) + 4, 44)
                ws.column_dimensions[column[0].column_letter].width = width
    return output.getvalue()

# =========================================================
# SESSION STATE
# =========================================================
for key, default in {
    "sel_fund":"All", "sel_region":"All", "sel_division":"All",
    "sel_package":"All", "sel_stage":"All", "include_ro":False,
    "stage_overrides":{},
}.items():
    if key not in st.session_state:
        st.session_state[key] = default


def reset_filters():
    for key in ["sel_fund", "sel_region", "sel_division", "sel_package", "sel_stage"]:
        st.session_state[key] = "All"

# =========================================================
# DATA CONNECTION
# =========================================================
with st.sidebar:
    st.markdown("### 🎛️ Data Connection")
    source_mode = st.radio("Source Mode", ["Live Google Sheet", "Upload Excel (.xlsx)"])
    tables = None
    if source_mode == "Live Google Sheet":
        if st.button("🔄 Sync Latest Google Sheet", use_container_width=True, type="primary"):
            st.cache_data.clear()
            st.session_state.stage_overrides = {}
        try:
            with st.spinner("Syncing..."):
                tables = load_from_gsheet(SHEET_ID)
        except Exception as exc:
            st.error("Unable to fetch Google Sheet.")
            st.caption(str(exc))
    else:
        uploaded = st.file_uploader("Upload monitoring workbook", type=["xlsx"])
        if uploaded:
            tables = load_from_excel(uploaded.read())

    if APP_PASSWORD and st.button("🔒 Lock / Logout", use_container_width=True):
        st.session_state.authenticated = False
        st.rerun()

if not tables:
    st.info("Open the sidebar to connect the Google Sheet or upload an Excel file.")
    st.stop()

base = build_processed_df(tables)
if base.empty:
    st.warning("No valid records found in 2025c or 2026.")
    st.stop()

if st.session_state.stage_overrides:
    override = base["Row_ID"].map(st.session_state.stage_overrides)
    base["Delivery_Stage"] = np.where(override.notna(), override, base["Delivery_Stage"])

# =========================================================
# HEADER / FILTERS
# =========================================================
st.markdown(
    """
    <div class="hero">
      <h1>DCP Delivery & Division Billing Portal</h1>
      <p>Status-based monitoring. A recipient marked Completed is treated as cleared for billing; actual DR and IAR dates are not required in this dashboard.</p>
    </div>
    """,
    unsafe_allow_html=True,
)

fund_options = ["All"] + sorted(base["Fund_Clean"].unique())
region_options = ["All"] + sorted(base["Region_Clean"].unique(), key=lambda x: REGION_ORDER.get(x, 999))
div_pool_df = base if st.session_state.include_ro else base[base["Is_SDO"]]
if st.session_state.sel_region != "All":
    div_pool_df = div_pool_df[div_pool_df["Region_Clean"] == st.session_state.sel_region]
division_options = ["All"] + sorted(div_pool_df["Division_Clean"].unique())
if st.session_state.sel_division not in division_options:
    st.session_state.sel_division = "All"
package_options = ["All"] + sorted(base["Package_Display"].unique())
stage_options = ["All"] + STAGES

c1, c2, c3, c4, c5, c6 = st.columns([1.1, 1, 1.5, 1.4, 1.4, .8])
with c1: st.selectbox("Fund Source", fund_options, key="sel_fund")
with c2: st.selectbox("Region", region_options, key="sel_region")
with c3: st.selectbox("Division (SDO)", division_options, key="sel_division")
with c4: st.selectbox("DCP Package", package_options, key="sel_package")
with c5: st.selectbox("Delivery Status", stage_options, key="sel_stage")
with c6:
    st.write("")
    st.button("🔄 Reset", on_click=reset_filters, use_container_width=True)

filtered = base.copy()
for state_key, column in [
    ("sel_fund", "Fund_Clean"), ("sel_region", "Region_Clean"),
    ("sel_division", "Division_Clean"), ("sel_package", "Package_Display"),
    ("sel_stage", "Delivery_Stage"),
]:
    value = st.session_state[state_key]
    if value != "All":
        filtered = filtered[filtered[column] == value]

if filtered.empty:
    st.warning("No records match the selected filters.")
    st.stop()

# =========================================================
# RECIPIENT-LEVEL STATUS
# Use highest-priority status for each recipient.
# =========================================================
priority = {"Completed":4, "Delivered (Pending Validation)":3, "Successful Rebid":2, "Ongoing Delivery":1}
filtered["Stage_Rank"] = filtered["Delivery_Stage"].map(priority).fillna(0)
recipient_status = filtered.sort_values("Stage_Rank").drop_duplicates(["Fund_Clean", "Recipient_Key"], keep="last")
recipient_total = len(recipient_status)
counts = recipient_status["Delivery_Stage"].value_counts()
completed = int(counts.get("Completed", 0))
pending_validation = int(counts.get("Delivered (Pending Validation)", 0))
ongoing = int(counts.get("Ongoing Delivery", 0))
rebid = int(counts.get("Successful Rebid", 0))
remaining_packages = int(compute_remaining_packages(filtered, ["Fund_Clean"]).sum())

# Division billing is evaluated from the full relevant slice, not the stage filter.
div_eval = base.copy()
if not st.session_state.include_ro:
    div_eval = div_eval[div_eval["Is_SDO"]]
for state_key, column in [
    ("sel_fund", "Fund_Clean"), ("sel_region", "Region_Clean"),
    ("sel_division", "Division_Clean"), ("sel_package", "Package_Display"),
]:
    value = st.session_state[state_key]
    if value != "All":
        div_eval = div_eval[div_eval[column] == value]

div_eval["Stage_Rank"] = div_eval["Delivery_Stage"].map(priority).fillna(0)
div_rec = div_eval.sort_values("Stage_Rank").drop_duplicates(["Fund_Clean", "Recipient_Key"], keep="last")
rem_by_div = compute_remaining_packages(div_eval, ["Division_Clean"])

div_rows = []
for division, group in div_rec.groupby("Division_Clean"):
    total = len(group)
    stage_counts = group["Delivery_Stage"].value_counts()
    comp = int(stage_counts.get("Completed", 0))
    pend = int(stage_counts.get("Delivered (Pending Validation)", 0))
    ong = int(stage_counts.get("Ongoing Delivery", 0))
    reb = int(stage_counts.get("Successful Rebid", 0))
    region = group["Region_Clean"].mode().iloc[0] if not group["Region_Clean"].mode().empty else ""
    if total > 0 and comp == total:
        billing = "✅ READY FOR BILLING"
    elif total > 0 and comp + pend == total:
        billing = "🟡 PENDING VALIDATION"
    else:
        billing = "❌ INCOMPLETE DELIVERY"
    div_rows.append({
        "Region":region, "Division (SDO)":division, "Total Recipients":total,
        "Remaining Packages":int(rem_by_div.get(division, 0)),
        "Ongoing Delivery":ong, "Successful Rebid":reb,
        "Delivered (Pending Validation)":pend, "Completed":comp,
        "Completion (%)":round(comp / total * 100, 1) if total else 0,
        "Billing Status":billing,
    })

div_df = pd.DataFrame(div_rows)
if not div_df.empty:
    div_df["_order"] = div_df["Region"].map(lambda x: REGION_ORDER.get(x, 999))
    div_df = div_df.sort_values(["Completion (%)", "_order", "Division (SDO)"], ascending=[False, True, True]).drop(columns="_order")
ready_divisions = int((div_df["Billing Status"] == "✅ READY FOR BILLING").sum()) if not div_df.empty else 0

# KPIs
kpis = [
    ("Total Recipients", f"{recipient_total:,}", "Unique schools/offices"),
    ("Remaining Packages", f"{remaining_packages:,}", "Ongoing delivery or rebid"),
    ("Ongoing Delivery", f"{ongoing:,}", f"{ongoing/recipient_total*100:.1f}%" if recipient_total else "0%"),
    ("Successful Rebid", f"{rebid:,}", "Rebid allocation"),
    ("Pending Validation", f"{pending_validation:,}", "Delivered, not yet completed"),
    ("Completed", f"{completed:,}", "Cleared for billing"),
    ("Divisions Billable", f"{ready_divisions:,} / {len(div_df):,}", "All recipients completed"),
]
cols = st.columns(7)
for col, (label, value, foot) in zip(cols, kpis):
    with col:
        st.markdown(f'<div class="kpi"><div class="kpi-label">{label}</div><div class="kpi-value">{value}</div><div class="kpi-foot">{foot}</div></div>', unsafe_allow_html=True)

# =========================================================
# SUMMARY TABLES
# =========================================================
regions = sorted(recipient_status["Region_Clean"].unique(), key=lambda x: REGION_ORDER.get(x, 999))
rem_by_region = compute_remaining_packages(filtered, ["Region_Clean"])
regional_rows = []
for region in regions:
    group = recipient_status[recipient_status["Region_Clean"] == region]
    sc = group["Delivery_Stage"].value_counts()
    total = len(group)
    comp = int(sc.get("Completed", 0))
    regional_rows.append({
        "Region":region, "Total Recipients":total,
        "Remaining Packages":int(rem_by_region.get(region, 0)),
        "Ongoing Delivery":int(sc.get("Ongoing Delivery", 0)),
        "Successful Rebid":int(sc.get("Successful Rebid", 0)),
        "Delivered (Pending Validation)":int(sc.get("Delivered (Pending Validation)", 0)),
        "Completed":comp, "Completion (%)":round(comp / total * 100, 1) if total else 0,
    })
regional_df = pd.DataFrame(regional_rows)

rem_by_package = compute_remaining_packages(filtered, ["Package_Display"])
package_rows = []
for package in sorted(filtered["Package_Display"].unique()):
    group = filtered[filtered["Package_Display"] == package]
    recipients = group.drop_duplicates(["Fund_Clean", "Recipient_Key"])
    pending_units = group[
        group["Delivery_Stage"].isin(["Ongoing Delivery", "Successful Rebid"])
        & ~group["Component_Lower"].str.contains("storage|charging cart", na=False)
    ]
    package_rows.append({
        "DCP Package":package, "Recipients":len(recipients),
        "Remaining Packages":int(rem_by_package.get(package, 0)),
        "Remaining Device Units":int((pending_units["Pkg_Count"] * pending_units["Units_Per_Pkg"]).sum()),
    })
package_df = pd.DataFrame(package_rows)

component_rows = []
for component in sorted(filtered.loc[filtered["Component_Display"] != "", "Component_Display"].unique()):
    group = filtered[filtered["Component_Display"] == component]
    pending = group[group["Delivery_Stage"].isin(["Ongoing Delivery", "Successful Rebid"])]
    units = (pending["Pkg_Count"] * pending["Units_Per_Pkg"]).sum()
    cost = (pending["Pkg_Count"] * pending["Units_Per_Pkg"] * pending["Unit_Cost_Num"]).sum()
    component_rows.append({
        "Component":component,
        "Recipients":len(group.drop_duplicates(["Fund_Clean", "Recipient_Key"])),
        "Remaining Units":int(units),
        "Avg Unit Cost":cost / units if units else 0,
        "Remaining Cost":cost,
    })
component_df = pd.DataFrame(component_rows)

# =========================================================
# TABS
# =========================================================
tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "📊 Executive Visuals", "💰 Division Billing Status", "🗺️ Regional Table",
    "📦 Packages & Components", "🔎 Lookup & Status Simulator",
])

with tab1:
    left, right = st.columns([1.45, 1])
    with left:
        st.markdown('<div class="section">📍 Regional Delivery Volume by Stage</div>', unsafe_allow_html=True)
        melt = regional_df.melt("Region", ["Completed", "Delivered (Pending Validation)", "Ongoing Delivery", "Successful Rebid"], var_name="Stage", value_name="Recipients")
        fig = px.bar(melt, x="Region", y="Recipients", color="Stage", barmode="stack", color_discrete_map=STAGE_COLORS, height=390)
        fig.update_layout(plot_bgcolor="white", paper_bgcolor="white", legend_title_text="")
        st.plotly_chart(fig, use_container_width=True)
    with right:
        st.markdown('<div class="section">🍩 Overall Delivery Status</div>', unsafe_allow_html=True)
        pie = pd.DataFrame({"Stage":STAGES, "Count":[completed, pending_validation, ongoing, rebid]})
        pie = pie[pie["Count"] > 0]
        fig = go.Figure(go.Pie(labels=pie["Stage"], values=pie["Count"], hole=.62, marker_colors=[STAGE_COLORS[x] for x in pie["Stage"]]))
        fig.update_layout(height=390, annotations=[dict(text=f"<b>{recipient_total:,}</b><br>Recipients", x=.5, y=.5, showarrow=False)])
        st.plotly_chart(fig, use_container_width=True)

with tab2:
    st.markdown('<div class="section">💰 Schools Division Billing Status</div>', unsafe_allow_html=True)
    st.caption("READY FOR BILLING means every recipient under the SDO is marked Completed in the source system. No DR/IAR date columns are used.")
    a, b = st.columns([2.5, 1])
    with a:
        billing_filter = st.radio("Billing Status", ["All", "✅ READY FOR BILLING", "🟡 PENDING VALIDATION", "❌ INCOMPLETE DELIVERY"], horizontal=True)
    with b:
        st.checkbox("Include Regional Offices", key="include_ro")
    view = div_df if billing_filter == "All" else div_df[div_df["Billing Status"] == billing_filter]
    st.dataframe(view, use_container_width=True, hide_index=True, height=min(650, max(150, len(view) * 35 + 50)))

with tab3:
    st.markdown('<div class="section">🗺️ Regional Delivery Monitoring</div>', unsafe_allow_html=True)
    st.dataframe(regional_df, use_container_width=True, hide_index=True, height=680)

with tab4:
    st.markdown('<div class="section">📦 DCP Package Distribution</div>', unsafe_allow_html=True)
    st.dataframe(package_df, use_container_width=True, hide_index=True)
    if not component_df.empty:
        st.markdown('<div class="section">💻 Component Breakdown & Valuation</div>', unsafe_allow_html=True)
        st.dataframe(component_df.style.format({"Avg Unit Cost":"₱{:,.2f}", "Remaining Cost":"₱{:,.2f}"}), use_container_width=True, hide_index=True)

with tab5:
    st.markdown('<div class="section">🔎 School / Division Lookup & Status Simulator</div>', unsafe_allow_html=True)
    st.caption("Status only. Simulation does not change the live Google Sheet.")
    q = st.text_input("Search School ID, Division, LOT, Supplier, or Recipient")
    lookup = filtered[[
        "Row_ID", "Fund_Clean", "LOT", "Region_Clean", "Division_Clean", "School_ID",
        "Recipient_Office", "Supplier", "Package_Display", "Pkg_Count", "Effective_Status", "Delivery_Stage",
    ]].copy()
    lookup.columns = ["Row_ID", "Fund", "LOT", "Region", "Division (SDO)", "School ID", "Recipient Office", "Supplier", "DCP Package", "Packages", "Source System Status", "Dashboard Status"]
    if q.strip():
        mask = lookup.drop(columns=["Row_ID"]).apply(lambda col: col.astype(str).str.contains(q.strip(), case=False, regex=False, na=False)).any(axis=1)
        lookup = lookup[mask]
    edited = st.data_editor(
        lookup.head(500), use_container_width=True, hide_index=True, height=520,
        disabled=["Row_ID", "Fund", "LOT", "Region", "Division (SDO)", "School ID", "Recipient Office", "Supplier", "DCP Package", "Packages", "Source System Status"],
        column_config={
            "Row_ID":None,
            "Dashboard Status":st.column_config.SelectboxColumn("Dashboard Status (Simulation)", options=STAGES, required=True),
        },
        key="status_editor",
    )
    original = dict(zip(lookup.head(500)["Row_ID"], lookup.head(500)["Dashboard Status"]))
    changed = False
    for _, row in edited.iterrows():
        rid, status = row["Row_ID"], row["Dashboard Status"]
        if status != original.get(rid):
            st.session_state.stage_overrides[rid] = status
            changed = True
    if changed:
        st.rerun()
    if st.button("↩️ Reset Simulated Status Changes"):
        st.session_state.stage_overrides = {}
        st.rerun()

# =========================================================
# EXPORT
# =========================================================
excel_bytes = build_excel(regional_df, div_df, package_df, component_df)
st.divider()
st.download_button(
    "📥 Export Full Excel Report (.xlsx)", data=excel_bytes,
    file_name=f"DCP_Status_Based_Billing_Report_{st.session_state.sel_fund}_{st.session_state.sel_region}.xlsx".replace(" ", "_"),
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    use_container_width=True, type="primary",
)
