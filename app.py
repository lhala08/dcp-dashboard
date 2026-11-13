import io
import re
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
    page_title="DCP Delivery & Division Billing Portal",
    page_icon="🔒",
    layout="wide",
    initial_sidebar_state="collapsed",
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
        padding-top: 1.2rem;
        padding-bottom: 2.5rem;
        max-width: 1480px;
    }
    .login-box {
        background: #ffffff;
        border: 1px solid #e2e8f0;
        border-top: 5px solid #1e3a8a;
        border-radius: 16px;
        padding: 2rem 2rem 1.5rem 2rem;
        box-shadow: 0 10px 25px -5px rgba(15, 23, 42, 0.08);
        margin-top: 3rem;
        text-align: center;
    }
    .login-title {
        font-size: 1.35rem;
        font-weight: 800;
        color: #0f172a;
        margin-bottom: 0.3rem;
    }
    .login-sub {
        font-size: 0.86rem;
        color: #64748b;
        margin-bottom: 1.2rem;
    }
    .hero-banner {
        background: linear-gradient(120deg, #0f172a 0%, #1e3a8a 55%, #2563eb 100%);
        padding: 1.4rem 1.8rem;
        border-radius: 16px;
        color: #ffffff;
        margin-bottom: 1rem;
        box-shadow: 0 10px 25px -5px rgba(30, 58, 138, 0.25);
        display: flex;
        justify-content: space-between;
        align-items: center;
        flex-wrap: wrap;
        gap: 1rem;
    }
    .hero-title {
        margin: 0;
        font-size: 1.65rem;
        font-weight: 800;
        letter-spacing: -0.02em;
        color: #ffffff;
    }
    .hero-subtitle {
        margin: 0.3rem 0 0 0;
        font-size: 0.88rem;
        color: #bfdbfe;
        font-weight: 500;
    }
    .hero-badge {
        background: rgba(255, 255, 255, 0.14);
        border: 1px solid rgba(255, 255, 255, 0.25);
        backdrop-filter: blur(6px);
        padding: 0.5rem 0.95rem;
        border-radius: 999px;
        font-size: 0.8rem;
        font-weight: 600;
        color: #f8fafc;
    }
    .kpi-card {
        background: #ffffff;
        border: 1px solid #e2e8f0;
        border-radius: 14px;
        padding: 0.95rem 1.1rem;
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
        margin-bottom: 0.35rem;
    }
    .kpi-label {
        font-size: 0.74rem;
        font-weight: 700;
        text-transform: uppercase;
        letter-spacing: 0.04em;
        color: #64748b;
    }
    .kpi-pill {
        font-size: 0.7rem;
        font-weight: 700;
        padding: 0.16rem 0.52rem;
        border-radius: 999px;
    }
    .kpi-number {
        font-size: 1.65rem;
        font-weight: 800;
        color: #0f172a;
        line-height: 1.15;
        letter-spacing: -0.02em;
    }
    .kpi-foot {
        font-size: 0.74rem;
        color: #94a3b8;
        margin-top: 0.28rem;
    }
    .section-header {
        font-size: 1.05rem;
        font-weight: 700;
        color: #0f172a;
        margin: 0.35rem 0 0.65rem 0;
        display: flex;
        align-items: center;
        gap: 0.45rem;
    }
    .click-hint {
        font-size: 0.78rem;
        color: #2563eb;
        background: #eff6ff;
        border: 1px solid #bfdbfe;
        padding: 0.35rem 0.75rem;
        border-radius: 8px;
        margin-bottom: 0.75rem;
        display: inline-block;
    }
    .stTabs [data-baseweb="tab-list"] {
        gap: 8px;
        background-color: #e2e8f0;
        padding: 5px;
        border-radius: 12px;
    }
    .stTabs [data-baseweb="tab"] {
        border-radius: 8px;
        padding: 8px 16px;
        font-weight: 600;
        font-size: 0.86rem;
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


# ==========================================
# SECRETS & PASSWORD GATE (WITH 1-CLICK VIP LINK)
# ==========================================
def get_secret(key: str, default: str = "") -> str:
    try:
        return str(st.secrets.get(key, default))
    except Exception:
        return default


SHEET_ID = "1uGsxNYJHhL0DOfODQbn60L2wXnC2qd-lKO56O8OzF4g"
APP_PASSWORD = "!dcp2026*"

if "authenticated" not in st.session_state:
    st.session_state["authenticated"] = False

try:
    url_key = st.query_params.get("key", "")
except Exception:
    url_key = ""

if APP_PASSWORD and url_key == APP_PASSWORD:
    st.session_state["authenticated"] = True

if APP_PASSWORD and not st.session_state["authenticated"]:
    _, center_col, _ = st.columns([1.2, 1.3, 1.2])
    with center_col:
        st.markdown(
            """
            <div class="login-box">
                <div style="font-size: 2.2rem; margin-bottom: 0.4rem;">🔐</div>
                <div class="login-title">DepEd DCP Monitoring Portal</div>
                <div class="login-sub">Restricted Internal Dashboard • Enter authorized access password</div>
            </div>
            """,
            unsafe_allow_html=True,
        )
        with st.form("login_form"):
            pwd_input = st.text_input("Access Password", type="password", placeholder="Enter office password...")
            submitted = st.form_submit_button("Unlock Dashboard", use_container_width=True, type="primary")
            if submitted:
                if pwd_input == APP_PASSWORD:
                    st.session_state["authenticated"] = True
                    st.rerun()
                else:
                    st.error("Incorrect password. Please try again.")
    st.stop()


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
# DATA LOADING, SDO vs RO SEPARATION & DOCS CHECK
# ==========================================
def clean_series(s: pd.Series) -> pd.Series:
    return (
        s.fillna("")
        .astype(str)
        .str.replace("\xa0", " ", regex=False)
        .str.strip()
    )


def has_valid_entry(s: pd.Series) -> pd.Series:
    invalid_tokens = {
        "", "-", "--", "n/a", "na", "none", "null", "tba",
        "pending", "incomplete", "for iar", "for dr", "for submission",
        "no", "0", "nan", "nat"
    }
    cleaned = s.fillna("").astype(str).str.strip().str.lower()
    return ~cleaned.isin(invalid_tokens)


def is_explicitly_incomplete(s: pd.Series) -> pd.Series:
    cleaned = s.fillna("").astype(str).str.strip().str.lower()
    return cleaned.str.contains("incomplete|pending|lacking|missing|for submission|no mov|no iar", regex=True, na=False)


def normalize_office_and_sdo(raw_div: str, region_str: str) -> tuple[str, bool]:
    val = re.sub(r"\s+", " ", str(raw_div or "")).strip()
    reg = str(region_str or "").upper().strip()
    if not val:
        return ("UNSPECIFIED DIVISION", False)

    up = val.upper()

    ro_patterns = [
        r"^(DEPED\s+)?(RO|R\.O\.)[\s\-]*([IVX0-9]+|CAR|CARAGA|NCR|NIR|BARMM|MIMAROPA)\b",
        r"^REGION\s+([IVX0-9]+|CAR|CARAGA|NCR|NIR|BARMM|MIMAROPA)$",
        r"REGIONAL\s+OFFICE",
        r"CENTRAL\s+OFFICE",
        r"^DEPED\s+CO\b",
        r"^CO\s*\-",
        r"BUREAU\s+OF",
    ]
    if any(re.search(pat, up) for pat in ro_patterns):
        return (up, False)

    up = re.sub(r"^(CITY\s+SCHOOLS\s+DIVISION\s+OF|SCHOOLS\s+DIVISION\s+OF|DIVISION\s+OF|SDO[\s\-]+)\s*", "", up).strip()

    if up == "SAN CARLOS CITY" and reg:
        up = f"SAN CARLOS CITY ({reg})"

    return (up if up else "UNSPECIFIED DIVISION", True)


def find_col_by_keywords(columns, keywords) -> str:
    for col in columns:
        c_low = str(col).lower().strip()
        if any(k in c_low for k in keywords):
            return col
    return ""


def standardize_tab(df_raw: pd.DataFrame, default_fund: str) -> pd.DataFrame:
    df = df_raw.copy()
    orig_cols = list(df.columns)

    div_col_name = find_col_by_keywords(orig_cols, ["division", "sdo"])
    explicit_div = clean_series(df[div_col_name]) if div_col_name else pd.Series([""] * len(df))

    while df.shape[1] < 16:
        df[f"extra_{df.shape[1]}"] = ""

    df_std = df.iloc[:, :16].copy()
    df_std.columns = [
        "LOT", "Region", "DCP_Package", "No_of_Package", "School_ID",
        "Recipient_Office", "Component", "Units_Per_Package", "Unit_Cost",
        "Supplier", "DR_Date", "IAR_Date", "IAR_Docs", "Col_N_Status",
        "Status", "Fund_Source"
    ]

    for col in df_std.columns:
        df_std[col] = clean_series(df_std[col])

    df_std["Explicit_Division"] = explicit_div.values
    df_std["Fund_Clean"] = default_fund
    df_std = df_std[df_std["DCP_Package"] != ""].copy()
    return df_std


@st.cache_data(ttl=300, show_spinner=False)
def load_workbook_from_gsheet(sheet_id: str) -> bytes:
    url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=xlsx"
    resp = requests.get(url, timeout=30)
    resp.raise_for_status()
    return resp.content


@st.cache_data(ttl=300, show_spinner=False)
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

    raw_status_lookup = {}
    raw_div_lookup = {}
    if "Raw_Delivery" in sheet_names:
        raw_df = xl.parse("Raw_Delivery", dtype=str)
        raw_cols = list(raw_df.columns)

        site_col = find_col_by_keywords(raw_cols, ["site code", "school id", "school_id", "beis"])
        if not site_col and raw_df.shape[1] >= 6:
            site_col = raw_cols[5]

        stat_col = find_col_by_keywords(raw_cols, ["current status", "delivery status", "status"])
        if not stat_col and raw_df.shape[1] >= 7:
            stat_col = raw_cols[6]

        raw_div_col = find_col_by_keywords(raw_cols, ["division", "sdo"])

        if site_col:
            site_codes = clean_series(raw_df[site_col])
            if stat_col:
                cur_statuses = clean_series(raw_df[stat_col])
                for sid, st_val in zip(site_codes, cur_statuses):
                    if sid and st_val:
                        raw_status_lookup[sid] = st_val
            if raw_div_col:
                cur_divs = clean_series(raw_df[raw_div_col])
                for sid, d_val in zip(site_codes, cur_divs):
                    if sid and d_val:
                        raw_div_lookup[sid] = d_val

    raw_mapped = df["School_ID"].map(raw_status_lookup).fillna("")
    effective_raw_status = np.where(raw_mapped != "", raw_mapped, df["Col_N_Status"])
    df["Effective_Raw_Status"] = effective_raw_status

    df["Region_Clean"] = df["Region"].str.upper().str.strip()
    df["Package_Lower"] = df["DCP_Package"].str.lower()
    df["Package_Display"] = df["Package_Lower"].map(PACKAGE_DISPLAY).fillna(df["DCP_Package"])
    df["Component_Lower"] = df["Component"].str.lower()
    df["Component_Display"] = df["Component_Lower"].map(COMPONENT_DISPLAY).fillna(df["Component"])

    raw_div_mapped = df["School_ID"].map(raw_div_lookup).fillna("")
    div_resolved = np.where(
        df["Explicit_Division"] != "",
        df["Explicit_Division"],
        np.where(
            raw_div_mapped != "",
            raw_div_mapped,
            np.where(df["Recipient_Office"] != "", df["Recipient_Office"], "UNSPECIFIED DIVISION"),
        ),
    )

    norm_results = [normalize_office_and_sdo(x, r) for x, r in zip(div_resolved, df["Region_Clean"])]
    df["Division_Clean"] = [r[0] for r in norm_results]
    df["Is_SDO"] = [r[1] for r in norm_results]

    is_nonteaching = df["Package_Lower"] == "laptop for non-teaching"
    df["Delivery_Level"] = np.where(
        is_nonteaching,
        np.where(df["Is_SDO"], "Division Office (L4NT)", "Regional / Central Office (L4NT)"),
        "School-Level Delivery",
    )

    office_key = df["Region_Clean"] + "|" + df["Division_Clean"]
    df["Recipient_Key"] = np.where(
        is_nonteaching,
        np.where(df["Division_Clean"] != "UNSPECIFIED DIVISION", office_key, np.where(df["School_ID"] != "", df["School_ID"], df["LOT"])),
        np.where(df["School_ID"] != "", df["School_ID"], np.where(df["Recipient_Office"] != "", office_key, df["LOT"])),
    )

    df = df[df["Recipient_Key"] != ""].copy()
    df["Row_ID"] = np.arange(len(df))

    df["Pkg_Count"] = pd.to_numeric(df["No_of_Package"].str.replace("[^0-9.-]", "", regex=True), errors="coerce").fillna(0)
    df["Units_Per_Pkg"] = pd.to_numeric(df["Units_Per_Package"].str.replace("[^0-9.-]", "", regex=True), errors="coerce").fillna(0)
    df["Unit_Cost"] = pd.to_numeric(df["Unit_Cost"].str.replace("[^0-9.-]", "", regex=True), errors="coerce").fillna(0)

    return df


def classify_delivery_stages(df_in: pd.DataFrame, strict_docs_mode: bool = False) -> pd.DataFrame:
    df = df_in.copy()
    df["Has_DR"] = has_valid_entry(df["DR_Date"])
    df["Has_IAR"] = has_valid_entry(df["IAR_Date"])
    df["Has_IAR_Docs"] = has_valid_entry(df["IAR_Docs"])
    docs_incomplete = is_explicitly_incomplete(df["IAR_Docs"])

    eff_lower = df["Effective_Raw_Status"].str.lower().str.strip()
    orig_status_lower = df["Status"].str.lower().str.strip()

    if strict_docs_mode:
        is_completed = df["Has_IAR"] & df["Has_IAR_Docs"] & (~docs_incomplete)
        is_delivered_pending = (
            df["Has_DR"]
            | df["Has_IAR"]
            | eff_lower.isin(["complete", "completed"])
            | eff_lower.str.contains("pending mov", na=False)
            | (orig_status_lower.isin(["completed", "pending mov"]))
        ) & (~is_completed)
    else:
        raw_or_status_done = (eff_lower.isin(["complete", "completed"]) | (orig_status_lower == "completed")) & (~docs_incomplete)
        iar_done = df["Has_IAR"] & (~docs_incomplete)
        is_completed = iar_done | raw_or_status_done
        is_delivered_pending = (
            df["Has_DR"]
            | df["Has_IAR"]
            | docs_incomplete
            | eff_lower.str.contains("pending mov", na=False)
            | (orig_status_lower == "pending mov")
        ) & (~is_completed)

    conditions = [
        is_completed,
        is_delivered_pending,
        orig_status_lower == "successful rebid",
    ]
    choices = ["Completed", "Delivered, Docs Pending", "Successful Rebid"]
    df["Delivery_Stage"] = np.select(conditions, choices, default="Ongoing Delivery")
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


def build_styled_excel(
    reg_display: pd.DataFrame,
    sdo_display: pd.DataFrame,
    ro_display: pd.DataFrame,
    pkg_display: pd.DataFrame,
    comp_display: pd.DataFrame,
) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        reg_display.to_excel(writer, sheet_name="Regional_Summary", index=False)
        if not sdo_display.empty:
            sdo_display.to_excel(writer, sheet_name="SDO_Billing_Status", index=False)
        if not ro_display.empty:
            ro_display.to_excel(writer, sheet_name="RO_CO_L4NT_Billing", index=False)
        pkg_display.to_excel(writer, sheet_name="Package_Distribution", index=False)
        if not comp_display.empty:
            comp_display.to_excel(writer, sheet_name="Component_Breakdown", index=False)

        header_fill = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")
        header_font = Font(color="FFFFFF", bold=True, size=11)
        total_fill = PatternFill(start_color="EFF6FF", end_color="EFF6FF", fill_type="solid")
        total_font = Font(bold=True, size=11)
        ready_fill = PatternFill(start_color="DCFCE7", end_color="DCFCE7", fill_type="solid")

        for sheet_name in writer.sheets:
            ws = writer.sheets[sheet_name]
            for cell in ws[1]:
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = Alignment(horizontal="center", vertical="center")

            if sheet_name in ("SDO_Billing_Status", "RO_CO_L4NT_Billing"):
                for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
                    if any("ELIGIBLE FOR BILLING" in str(c.value or "") and "NOT" not in str(c.value or "") for c in row):
                        for c in row:
                            c.fill = ready_fill

            last_row = ws.max_row
            if last_row > 1:
                for cell in ws[last_row]:
                    cell.fill = total_fill
                    cell.font = total_font

            for col in ws.columns:
                max_len = max((len(str(cell.value or "")) for cell in col), default=10)
                ws.column_dimensions[col[0].column_letter].width = min(max_len + 5, 44)

    return output.getvalue()


# ==========================================
# SESSION STATE FOR INTERACTIVE FILTERING & EDITS
# ==========================================
for k, v in {
    "sel_fund": "All",
    "sel_region": "All",
    "sel_division": "All",
    "sel_package": "All",
    "sel_stage": "All",
    "strict_docs_mode": False,
    "stage_overrides": {},
    "dr_overrides": {},
    "iar_overrides": {},
    "docs_overrides": {},
}.items():
    if k not in st.session_state:
        st.session_state[k] = v


def reset_all_filters():
    st.session_state["sel_fund"] = "All"
    st.session_state["sel_region"] = "All"
    st.session_state["sel_division"] = "All"
    st.session_state["sel_package"] = "All"
    st.session_state["sel_stage"] = "All"


# ==========================================
# SIDEBAR DATA CONNECTION
# ==========================================
with st.sidebar:
    st.markdown("### 🎛️ Data Connection")
    default_idx = 0 if SHEET_ID else 1
    source_mode = st.radio(
        "Source Mode",
        ["Live Google Sheet (via Secrets)", "Upload Excel (.xlsx - 100% Private)"],
        index=default_idx,
    )

    excel_data = None
    if source_mode.startswith("Live Google Sheet"):
        if not SHEET_ID:
            st.warning("`SHEET_ID` is not set in Streamlit Secrets yet. Add it in App Settings > Secrets, or use Upload Excel below.")
        else:
            if st.button("🔄 Sync Latest Google Sheet", use_container_width=True, type="primary"):
                st.cache_data.clear()
                st.session_state["stage_overrides"] = {}
                st.session_state["dr_overrides"] = {}
                st.session_state["iar_overrides"] = {}
                st.session_state["docs_overrides"] = {}
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

    if APP_PASSWORD:
        st.markdown("---")
        if st.button("🔒 Lock / Logout", use_container_width=True):
            st.session_state["authenticated"] = False
            try:
                st.query_params.clear()
            except Exception:
                pass
            st.rerun()

if excel_data is None:
    st.info("👈 Open the top-left sidebar (`>`) to connect your Google Sheet or upload an Excel file.")
    st.stop()

df_base = process_data(excel_data)
if df_base.empty:
    st.warning("No valid records found in '2025c' or '2026'.")
    st.stop()

df = df_base.copy()
if st.session_state["dr_overrides"]:
    dr_ov = df["Row_ID"].map(st.session_state["dr_overrides"])
    df["DR_Date"] = np.where(dr_ov.notna(), dr_ov, df["DR_Date"])
if st.session_state["iar_overrides"]:
    iar_ov = df["Row_ID"].map(st.session_state["iar_overrides"])
    df["IAR_Date"] = np.where(iar_ov.notna(), iar_ov, df["IAR_Date"])
if st.session_state["docs_overrides"]:
    docs_ov = df["Row_ID"].map(st.session_state["docs_overrides"])
    df["IAR_Docs"] = np.where(docs_ov.notna(), docs_ov, df["IAR_Docs"])

df = classify_delivery_stages(df, strict_docs_mode=st.session_state["strict_docs_mode"])

if st.session_state["stage_overrides"]:
    override_series = df["Row_ID"].map(st.session_state["stage_overrides"])
    df["Delivery_Stage"] = np.where(override_series.notna(), override_series, df["Delivery_Stage"])

# ==========================================
# EXECUTIVE HERO BANNER
# ==========================================
active_badges = []
for label, key in [
    ("Fund", "sel_fund"),
    ("Region", "sel_region"),
    ("Division/Office", "sel_division"),
    ("Package", "sel_package"),
    ("Status", "sel_stage"),
]:
    if st.session_state[key] != "All":
        active_badges.append(f"{label}: <b>{st.session_state[key]}</b>")

subtitle_html = (
    " &nbsp;•&nbsp; ".join(active_badges)
    if active_badges
    else "SDO Billing (Schools + Division L4NT) separated from Regional/Central Office L4NT &nbsp;•&nbsp; Strict Docs Verified"
)

total_sim_edits = (
    len(st.session_state["stage_overrides"])
    + len(st.session_state["dr_overrides"])
    + len(st.session_state["iar_overrides"])
    + len(st.session_state["docs_overrides"])
)
sim_note = f" &nbsp;|&nbsp; 🧪 {total_sim_edits} simulated edit(s) active" if total_sim_edits else ""
sec_badge = "🔒 Secured & Protected" if APP_PASSWORD else "● Executive Monitoring"

st.markdown(
    f"""
    <div class="hero-banner">
        <div>
            <div class="hero-title">DepEd DCP Delivery & Division Billing Portal</div>
            <div class="hero-subtitle">{subtitle_html}{sim_note}</div>
        </div>
        <div class="hero-badge">{sec_badge}</div>
    </div>
    """,
    unsafe_allow_html=True,
)

# ==========================================
# INLINE INTERACTIVE SLICER BAR
# ==========================================
fund_options = ["All"] + sorted(df["Fund_Clean"].unique().tolist())
region_options = ["All"] + sorted(
    df["Region_Clean"].unique().tolist(),
    key=lambda r: REGION_ORDER.get(r, 999),
)

if st.session_state["sel_region"] != "All":
    div_pool = df[df["Region_Clean"] == st.session_state["sel_region"]]["Division_Clean"].unique().tolist()
else:
    div_pool = df["Division_Clean"].unique().tolist()
division_options = ["All"] + sorted(div_pool)

if st.session_state["sel_division"] not in division_options:
    st.session_state["sel_division"] = "All"

pkg_options = ["All"] + sorted(df["Package_Display"].unique().tolist())
stage_options = ["All", "Completed", "Delivered, Docs Pending", "Ongoing Delivery", "Successful Rebid"]

s1, s2, s3, s4, s5, s6 = st.columns([1.1, 1.0, 1.45, 1.35, 1.25, 0.85])
with s1:
    st.selectbox("Fund Source", fund_options, key="sel_fund")
with s2:
    st.selectbox("Region", region_options, key="sel_region")
with s3:
    st.selectbox("Division (SDO) / RO", division_options, key="sel_division")
with s4:
    st.selectbox("DCP Package", pkg_options, key="sel_package")
with s5:
    st.selectbox("Delivery Status", stage_options, key="sel_stage")
with s6:
    st.markdown("<div style='height: 1.75rem;'></div>", unsafe_allow_html=True)
    st.button("🔄 Reset", on_click=reset_all_filters, use_container_width=True)

filtered = df.copy()
if st.session_state["sel_fund"] != "All":
    filtered = filtered[filtered["Fund_Clean"] == st.session_state["sel_fund"]]
if st.session_state["sel_region"] != "All":
    filtered = filtered[filtered["Region_Clean"] == st.session_state["sel_region"]]
if st.session_state["sel_division"] != "All":
    filtered = filtered[filtered["Division_Clean"] == st.session_state["sel_division"]]
if st.session_state["sel_package"] != "All":
    filtered = filtered[filtered["Package_Display"] == st.session_state["sel_package"]]
if st.session_state["sel_stage"] != "All":
    filtered = filtered[filtered["Delivery_Stage"] == st.session_state["sel_stage"]]

if filtered.empty:
    st.warning("No records match the selected filters. Click **🔄 Reset** above to clear filters.")
    st.stop()

# ==========================================
# KPI & SEPARATE SDO vs RO/CO BILLING COMPUTATIONS
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


def build_billing_table(eval_subset: pd.DataFrame, entity_col_label: str) -> pd.DataFrame:
    if eval_subset.empty:
        return pd.DataFrame()

    sub = eval_subset.copy()
    sub["Stage_Rank"] = sub["Delivery_Stage"].map(stage_priority).fillna(0)

    dr_by_rec = sub.groupby(["Fund_Clean", "Recipient_Key"])["Has_DR"].any()
    iar_by_rec = sub.groupby(["Fund_Clean", "Recipient_Key"])["Has_IAR"].any()
    docs_by_rec = sub.groupby(["Fund_Clean", "Recipient_Key"])["Has_IAR_Docs"].any()

    rec_u = (
        sub.sort_values("Stage_Rank")
        .drop_duplicates(subset=["Fund_Clean", "Recipient_Key"], keep="last")
        .copy()
    )
    rec_u["Rec_Has_DR"] = rec_u.set_index(["Fund_Clean", "Recipient_Key"]).index.map(dr_by_rec).fillna(False)
    rec_u["Rec_Has_IAR"] = rec_u.set_index(["Fund_Clean", "Recipient_Key"]).index.map(iar_by_rec).fillna(False)
    rec_u["Rec_Has_Docs"] = rec_u.set_index(["Fund_Clean", "Recipient_Key"]).index.map(docs_by_rec).fillna(False)

    rem_by_ent = compute_remaining_packages(sub, ["Division_Clean"])

    rows = []
    for ent_name, grp in rec_u.groupby("Division_Clean", sort=False):
        reg_d = grp["Region_Clean"].mode().iloc[0] if not grp["Region_Clean"].mode().empty else grp["Region_Clean"].iloc[0]
        tot_d = len(grp)
        school_cnt = int((grp["Delivery_Level"] == "School-Level Delivery").sum())
        l4nt_cnt = int((grp["Delivery_Level"] != "School-Level Delivery").sum())

        comp_d = int((grp["Delivery_Stage"] == "Completed").sum())
        pend_d = int((grp["Delivery_Stage"] == "Delivered, Docs Pending").sum())

        dr_cnt = int(grp["Rec_Has_DR"].sum())
        iar_cnt = int(grp["Rec_Has_IAR"].sum())
        docs_cnt = int(grp["Rec_Has_Docs"].sum())

        deliv_rate = ((comp_d + pend_d) / tot_d * 100) if tot_d > 0 else 0.0
        comp_rate = (comp_d / tot_d * 100) if tot_d > 0 else 0.0

        if tot_d > 0 and comp_d == tot_d and pend_d == 0:
            pay_status = "✅ ELIGIBLE FOR BILLING"
        else:
            pay_status = "❌ NOT ELIGIBLE (INCOMPLETE DOCS / DELIVERY)"

        rows.append({
            "Region": reg_d,
            entity_col_label: ent_name,
            "School Recipients": school_cnt,
            "L4NT Office Allocations": l4nt_cnt,
            "Total Recipients": tot_d,
            "Remaining Packages": int(rem_by_ent.get(ent_name, 0)),
            "Delivered, Docs Pending": pend_d,
            "Completed (Docs Verified)": comp_d,
            "With DR Date": dr_cnt,
            "With IAR Date": iar_cnt,
            "With IAR Docs": docs_cnt,
            "Delivery Rate (%)": round(deliv_rate, 1),
            "Completion Rate (%)": round(comp_rate, 1),
            "Billing Eligibility": pay_status,
        })

    res_df = pd.DataFrame(rows)
    if not res_df.empty:
        res_df["_reg_ord"] = res_df["Region"].map(lambda r: REGION_ORDER.get(r, 999))
        res_df = res_df.sort_values(
            by=["Completion Rate (%)", "Delivery Rate (%)", "_reg_ord", entity_col_label],
            ascending=[False, False, True, True],
        ).drop(columns=["_reg_ord"])
    return res_df


eval_base = df.copy()
if st.session_state["sel_fund"] != "All":
    eval_base = eval_base[eval_base["Fund_Clean"] == st.session_state["sel_fund"]]
if st.session_state["sel_region"] != "All":
    eval_base = eval_base[eval_base["Region_Clean"] == st.session_state["sel_region"]]
if st.session_state["sel_division"] != "All":
    eval_base = eval_base[eval_base["Division_Clean"] == st.session_state["sel_division"]]
if st.session_state["sel_package"] != "All":
    eval_base = eval_base[eval_base["Package_Display"] == st.session_state["sel_package"]]

sdo_df = build_billing_table(eval_base[eval_base["Is_SDO"]], "Division (SDO)")
ro_df = build_billing_table(eval_base[~eval_base["Is_SDO"]], "Regional / Central Office")

sdos_ready_cnt = int((sdo_df["Billing Eligibility"] == "✅ ELIGIBLE FOR BILLING").sum()) if not sdo_df.empty else 0
sdos_total_cnt = len(sdo_df)

ros_ready_cnt = int((ro_df["Billing Eligibility"] == "✅ ELIGIBLE FOR BILLING").sum()) if not ro_df.empty else 0
ros_total_cnt = len(ro_df)

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

kpi_row1 = [
    ("Total Recipients", f"{total_recipients:,}", "100%", "#eff6ff", "#1d4ed8", "Schools + Division/RO L4NT"),
    ("Remaining Packages", f"{total_rem_pkgs:,}", "Pending", "#f0f9ff", "#0284c7", "Undelivered packages"),
    ("Ongoing Delivery", f"{cnt_ongoing:,}", f"{pct_ongoing:.1f}%", "#eff6ff", "#2563eb", "Awarded / in transit"),
    ("Successful Rebid", f"{cnt_rebid:,}", "Rebid", "#f5f3ff", "#7c3aed", "Rebid allocation"),
]
kpi_row2 = [
    ("Docs Pending", f"{cnt_pending_docs:,}", f"{pct_delivered:.1f}% Deliv.", "#fff7ed", "#ea580c", "Delivered, incomplete docs"),
    ("Fully Completed", f"{cnt_completed:,}", f"{pct_completed:.1f}%", "#ecfdf5", "#059669", "Delivery + Docs verified"),
    ("SDOs Billable (~221)", f"{sdos_ready_cnt:,} / {sdos_total_cnt:,}", "SDOs", "#dcfce7", "#15803d", "Schools + SDO L4NT 100% done"),
    ("RO / CO L4NT Billable", f"{ros_ready_cnt:,} / {ros_total_cnt:,}", "RO / CO", "#e0f2fe", "#0369a1", "Regional/CO L4NT 100% done"),
]

for row_specs in (kpi_row1, kpi_row2):
    r_cols = st.columns(4)
    for col, (label, val, badge, bg_badge, fg_badge, foot) in zip(r_cols, row_specs):
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
    st.markdown("<div style='height: 0.55rem;'></div>", unsafe_allow_html=True)

st.markdown("<div style='height: 0.45rem;'></div>", unsafe_allow_html=True)

# ==========================================
# PREPARE SUMMARY TABLES
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
    deliv_rate = ((comp_r + pend_r) / tot_r * 100) if tot_r > 0 else 0.0
    comp_rate = (comp_r / tot_r * 100) if tot_r > 0 else 0.0

    reg_rows.append({
        "Region": str(reg),
        "Total Recipients": tot_r,
        "Remaining Packages": int(rem_by_reg.get(reg, 0)),
        "Ongoing Delivery": int((r_sub["Delivery_Stage"] == "Ongoing Delivery").sum()),
        "Successful Rebid": int((r_sub["Delivery_Stage"] == "Successful Rebid").sum()),
        "Delivered, Docs Pending": pend_r,
        "Completed": comp_r,
        "Delivery Rate (%)": round(deliv_rate, 1),
        "Completion Rate (%)": round(comp_rate, 1),
    })

reg_df = pd.DataFrame(reg_rows)
reg_total = pd.DataFrame([{
    "Region": "TOTAL",
    "Total Recipients": reg_df["Total Recipients"].sum(),
    "Remaining Packages": reg_df["Remaining Packages"].sum(),
    "Ongoing Delivery": reg_df["Ongoing Delivery"].sum(),
    "Successful Rebid": reg_df["Successful Rebid"].sum(),
    "Delivered, Docs Pending": reg_df["Delivered, Docs Pending"].sum(),
    "Completed": reg_df["Completed"].sum(),
    "Delivery Rate (%)": round(pct_delivered, 1),
    "Completion Rate (%)": round(pct_completed, 1),
}])
reg_display = pd.concat([reg_df, reg_total], ignore_index=True)


def append_billing_total(df_in: pd.DataFrame, name_col: str, summary_text: str, ready_cnt: int) -> pd.DataFrame:
    if df_in.empty:
        return pd.DataFrame()
    tot_r = df_in["Total Recipients"].sum()
    tot_row = pd.DataFrame([{
        "Region": "TOTAL",
        name_col: summary_text,
        "School Recipients": df_in["School Recipients"].sum(),
        "L4NT Office Allocations": df_in["L4NT Office Allocations"].sum(),
        "Total Recipients": tot_r,
        "Remaining Packages": df_in["Remaining Packages"].sum(),
        "Delivered, Docs Pending": df_in["Delivered, Docs Pending"].sum(),
        "Completed (Docs Verified)": df_in["Completed (Docs Verified)"].sum(),
        "With DR Date": df_in["With DR Date"].sum(),
        "With IAR Date": df_in["With IAR Date"].sum(),
        "With IAR Docs": df_in["With IAR Docs"].sum(),
        "Delivery Rate (%)": round((df_in["Completed (Docs Verified)"].sum() + df_in["Delivered, Docs Pending"].sum()) / tot_r * 100, 1) if tot_r > 0 else 0.0,
        "Completion Rate (%)": round(df_in["Completed (Docs Verified)"].sum() / tot_r * 100, 1) if tot_r > 0 else 0.0,
        "Billing Eligibility": f"{ready_cnt} ELIGIBLE",
    }])
    return pd.concat([df_in, tot_row], ignore_index=True)


sdo_display = append_billing_total(sdo_df, "Division (SDO)", f"{sdos_ready_cnt} of {sdos_total_cnt} SDOs Eligible", sdos_ready_cnt)
ro_display = append_billing_total(ro_df, "Regional / Central Office", f"{ros_ready_cnt} of {ros_total_cnt} RO/CO Eligible", ros_ready_cnt)

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
# INTERACTIVE TABS
# ==========================================
tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "📊 Executive Visuals (Clickable)",
    "💰 SDO & RO L4NT Billing Eligibility",
    "🗺️ Regional Table",
    "📦 Packages & Components",
    "🔎 Lookup & Docs/Date Simulator",
])

with tab1:
    st.markdown(
        '<div class="click-hint">💡 <b>Tip:</b> I-click ang kahit anong bar sa Region chart o slice sa Donut chart para mag-filter ang buong dashboard.</div>',
        unsafe_allow_html=True,
    )
    c_left, c_right = st.columns([1.45, 1])

    with c_left:
        st.markdown('<div class="section-header">📍 Regional Delivery Volume by Stage</div>', unsafe_allow_html=True)
        chart_melt = reg_df.melt(
            id_vars="Region",
            value_vars=["Completed", "Delivered, Docs Pending", "Ongoing Delivery", "Successful Rebid"],
            var_name="Stage",
            value_name="Recipients",
        )
        chart_melt["Region"] = chart_melt["Region"].astype(str)
        fig_bar = px.bar(
            chart_melt,
            x="Region",
            y="Recipients",
            color="Stage",
            barmode="stack",
            color_discrete_map=STAGE_COLORS,
            category_orders={"Region": [str(r) for r in regions_present]},
            height=390,
        )
        fig_bar.update_xaxes(type="category", title="", showgrid=False)
        fig_bar.update_layout(
            plot_bgcolor="rgba(0,0,0,0)",
            paper_bgcolor="rgba(0,0,0,0)",
            margin=dict(l=10, r=10, t=15, b=10),
            legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="left", x=0, title_text=""),
            yaxis=dict(title="Recipients", gridcolor="#e2e8f0"),
            clickmode="event+select",
        )
        reg_event = st.plotly_chart(fig_bar, use_container_width=True, on_select="rerun", key="chart_reg_bar")
        try:
            pts = (reg_event or {}).get("selection", {}).get("points", [])
            if pts:
                clicked_reg = str(pts[0].get("x", ""))
                if clicked_reg and clicked_reg in region_options and st.session_state["sel_region"] != clicked_reg:
                    st.session_state["sel_region"] = clicked_reg
                    st.rerun()
        except Exception:
            pass

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
        donut_event = st.plotly_chart(fig_donut, use_container_width=True, on_select="rerun", key="chart_donut")
        try:
            d_pts = (donut_event or {}).get("selection", {}).get("points", [])
            if d_pts:
                clicked_stage = str(d_pts[0].get("label", ""))
                if clicked_stage and clicked_stage in stage_options and st.session_state["sel_stage"] != clicked_stage:
                    st.session_state["sel_stage"] = clicked_stage
                    st.rerun()
        except Exception:
            pass

    st.markdown("<hr style='border:0; border-top:1px solid #e2e8f0; margin: 1rem 0;'>", unsafe_allow_html=True)

    c_bot1, c_bot2 = st.columns([1.2, 1])

    with c_bot1:
        st.markdown('<div class="section-header">📈 Delivery Rate (%) per Region</div>', unsafe_allow_html=True)
        rate_sorted = reg_df.sort_values("Delivery Rate (%)", ascending=True)
        rate_sorted["Region"] = rate_sorted["Region"].astype(str)
        fig_rate = px.bar(
            rate_sorted,
            x="Delivery Rate (%)",
            y="Region",
            orientation="h",
            text="Delivery Rate (%)",
            color="Delivery Rate (%)",
            color_continuous_scale=["#93c5fd", "#2563eb", "#10b981"],
            range_x=[0, 114],
            height=max(380, len(rate_sorted) * 26),
        )
        fig_rate.update_yaxes(type="category", title="")
        fig_rate.update_traces(texttemplate="%{text:.1f}%", textposition="outside")
        fig_rate.update_layout(
            plot_bgcolor="rgba(0,0,0,0)",
            paper_bgcolor="rgba(0,0,0,0)",
            coloraxis_showscale=False,
            margin=dict(l=10, r=20, t=10, b=10),
            xaxis=dict(title="Delivered + Completed (%)", gridcolor="#e2e8f0"),
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
            height=max(380, len(rate_sorted) * 26),
        )
        fig_pkg.update_traces(texttemplate="%{text:,}", textposition="outside")
        fig_pkg.update_layout(
            plot_bgcolor="rgba(0,0,0,0)",
            paper_bgcolor="rgba(0,0,0,0)",
            margin=dict(l=10, r=45, t=10, b=10),
            xaxis=dict(title="Remaining Packages", gridcolor="#e2e8f0"),
            yaxis=dict(title=""),
        )
        st.plotly_chart(fig_pkg, use_container_width=True)

with tab2:
    st.markdown('<div class="section-header">🏫 1. Schools Division Office (SDO) Billing Eligibility (~221 SDOs)</div>', unsafe_allow_html=True)
    st.caption(
        "Kasama sa bawat SDO ang **lahat ng school deliveries** sa ilalim nito **PLUS ang L4NT allocation ng mismong Division Office**. "
        "Hindi eligible ang SDO kung may kahit 1 recipient pa na kulang ang delivery, IAR Date, o IAR Docs."
    )

    opt_c1, opt_c2 = st.columns([2, 1.5])
    with opt_c1:
        pay_filter = st.radio(
            "Filter Billing Status",
            ["All", "✅ ELIGIBLE FOR BILLING", "❌ NOT ELIGIBLE (INCOMPLETE DOCS / DELIVERY)"],
            horizontal=True,
        )
    with opt_c2:
        st.checkbox(
            "Strict Docs Mode (Require BOTH IAR Date & IAR Docs)",
            key="strict_docs_mode",
            help="Kapag naka-check, kailangang may laman ang parehong IAR Date (Col L) at IAR Docs (Col M) para mabilang na Completed.",
        )

    sdo_view = sdo_display.copy()
    if not sdo_view.empty and pay_filter != "All":
        sdo_view = sdo_view[sdo_view["Billing Eligibility"] == pay_filter]

    if not sdo_view.empty:
        st.dataframe(
            sdo_view.style.format({
                "School Recipients": "{:,.0f}",
                "L4NT Office Allocations": "{:,.0f}",
                "Total Recipients": "{:,.0f}",
                "Remaining Packages": "{:,.0f}",
                "Delivered, Docs Pending": "{:,.0f}",
                "Completed (Docs Verified)": "{:,.0f}",
                "With DR Date": "{:,.0f}",
                "With IAR Date": "{:,.0f}",
                "With IAR Docs": "{:,.0f}",
                "Delivery Rate (%)": "{:.1f}%",
                "Completion Rate (%)": "{:.1f}%",
            }),
            use_container_width=True,
            hide_index=True,
            height=480,
        )
    else:
        st.info("No SDOs match the selected billing status filter.")

    st.markdown("<hr style='border:0; border-top:1px solid #cbd5e1; margin: 1.4rem 0;'>", unsafe_allow_html=True)
    st.markdown('<div class="section-header">🏢 2. Regional & Central Office L4NT Billing Eligibility (Separate from SDOs)</div>', unsafe_allow_html=True)
    st.caption("Para sa **Laptop for Non-Teaching (L4NT)** na direktang dine-deliver sa Regional Offices at Central Office (hiwalay sa SDO billing).")

    ro_view = ro_display.copy()
    if not ro_view.empty and pay_filter != "All":
        ro_view = ro_view[ro_view["Billing Eligibility"] == pay_filter]

    if not ro_view.empty:
        st.dataframe(
            ro_view.style.format({
                "School Recipients": "{:,.0f}",
                "L4NT Office Allocations": "{:,.0f}",
                "Total Recipients": "{:,.0f}",
                "Remaining Packages": "{:,.0f}",
                "Delivered, Docs Pending": "{:,.0f}",
                "Completed (Docs Verified)": "{:,.0f}",
                "With DR Date": "{:,.0f}",
                "With IAR Date": "{:,.0f}",
                "With IAR Docs": "{:,.0f}",
                "Delivery Rate (%)": "{:.1f}%",
                "Completion Rate (%)": "{:.1f}%",
            }),
            use_container_width=True,
            hide_index=True,
            height=360,
        )
    else:
        st.info("No Regional / Central Office L4NT records match the selected filters.")

with tab3:
    st.markdown('<div class="section-header">🗺️ Regional Delivery Monitoring Breakdown</div>', unsafe_allow_html=True)
    st.dataframe(
        reg_display.style.format({
            "Total Recipients": "{:,.0f}",
            "Remaining Packages": "{:,.0f}",
            "Ongoing Delivery": "{:,.0f}",
            "Successful Rebid": "{:,.0f}",
            "Delivered, Docs Pending": "{:,.0f}",
            "Completed": "{:,.0f}",
            "Delivery Rate (%)": "{:.1f}%",
            "Completion Rate (%)": "{:.1f}%",
        }),
        use_container_width=True,
        hide_index=True,
        height=680,
    )

with tab4:
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

with tab5:
    st.markdown('<div class="section-header">🔎 School / SDO / RO Lookup & IAR Date / Docs Simulator</div>', unsafe_allow_html=True)
    st.caption(
        "Mag-type ng petsa sa **DR Date** / **IAR Date** at ilagay ang **IAR Docs** (hal. `Complete`) para i-test kung magiging **✅ ELIGIBLE FOR BILLING** na ang SDO o Regional Office."
    )

    s_left, s_right = st.columns([3, 1])
    with s_left:
        search_q = st.text_input(
            "Search School ID, Division (SDO), Regional Office, LOT, or Supplier",
            placeholder="e.g. 502655, PASIG CITY, RO IV-A, or Columbia Technologies...",
        )
    with s_right:
        st.markdown("<div style='height: 1.75rem;'></div>", unsafe_allow_html=True)
        if st.button("↩️ Reset Simulated Edits", use_container_width=True):
            st.session_state["stage_overrides"] = {}
            st.session_state["dr_overrides"] = {}
            st.session_state["iar_overrides"] = {}
            st.session_state["docs_overrides"] = {}
            st.rerun()

    lookup_cols = [
        "Row_ID", "Fund_Clean", "LOT", "Region_Clean", "Division_Clean", "Delivery_Level",
        "School_ID", "Package_Display", "Pkg_Count", "DR_Date", "IAR_Date", "IAR_Docs", "Delivery_Stage"
    ]
    lookup_df = filtered[lookup_cols].copy()
    lookup_df.columns = [
        "Row_ID", "Fund", "LOT", "Region", "Division / Office", "Delivery Level",
        "School ID", "DCP Package", "Packages", "DR Date", "IAR Date", "IAR Docs", "Status"
    ]
    if search_q.strip():
        q = search_q.strip().lower()
        mask = lookup_df.drop(columns=["Row_ID"]).apply(
            lambda col: col.astype(str).str.lower().str.contains(q, na=False)
        ).any(axis=1)
        lookup_df = lookup_df[mask]

    st.caption(f"Showing {len(lookup_df):,} records (double-click DR Date, IAR Date, IAR Docs, or Status to simulate)")
    edited_df = st.data_editor(
        lookup_df.head(500),
        use_container_width=True,
        hide_index=True,
        height=520,
        disabled=["Row_ID", "Fund", "LOT", "Region", "Division / Office", "Delivery Level", "School ID", "DCP Package", "Packages"],
        column_config={
            "Row_ID": None,
            "DR Date": st.column_config.TextColumn("DR Date (Editable)"),
            "IAR Date": st.column_config.TextColumn("IAR Date (Editable)"),
            "IAR Docs": st.column_config.TextColumn("IAR Docs (Editable)"),
            "Status": st.column_config.SelectboxColumn(
                "Status (Editable)",
                options=["Completed", "Delivered, Docs Pending", "Ongoing Delivery", "Successful Rebid"],
                required=True,
            ),
        },
        key="school_status_editor",
    )

    orig_slice = lookup_df.head(500)
    orig_dr = dict(zip(orig_slice["Row_ID"], orig_slice["DR Date"]))
    orig_iar = dict(zip(orig_slice["Row_ID"], orig_slice["IAR Date"]))
    orig_docs = dict(zip(orig_slice["Row_ID"], orig_slice["IAR Docs"]))
    orig_st = dict(zip(orig_slice["Row_ID"], orig_slice["Status"]))

    changed = False
    for _, row in edited_df.iterrows():
        rid = row["Row_ID"]
        if str(row["DR Date"]) != str(orig_dr.get(rid, "")):
            st.session_state["dr_overrides"][rid] = str(row["DR Date"])
            changed = True
        if str(row["IAR Date"]) != str(orig_iar.get(rid, "")):
            st.session_state["iar_overrides"][rid] = str(row["IAR Date"])
            changed = True
        if str(row["IAR Docs"]) != str(orig_docs.get(rid, "")):
            st.session_state["docs_overrides"][rid] = str(row["IAR Docs"])
            changed = True
        if str(row["Status"]) != str(orig_st.get(rid, "")):
            st.session_state["stage_overrides"][rid] = str(row["Status"])
            changed = True

    if changed:
        st.rerun()

# ==========================================
# EXPORT BAR
# ==========================================
excel_bytes = build_styled_excel(reg_display, sdo_display, ro_display, pkg_display, comp_display)

st.markdown("<hr style='border:0; border-top:1px solid #e2e8f0; margin: 1.5rem 0 0.8rem 0;'>", unsafe_allow_html=True)
dl_c1, dl_c2 = st.columns([3, 1])
with dl_c1:
    st.caption("Kasama sa Excel export ang **Regional Summary**, **SDO Billing Status**, **RO/CO L4NT Billing**, **Package Distribution**, at **Component Breakdown**.")
with dl_c2:
    st.download_button(
        label="📥 Export Full Excel Report (.xlsx)",
        data=excel_bytes,
        file_name=f"DCP_Executive_Billing_Report_{st.session_state['sel_fund']}_{st.session_state['sel_region']}.xlsx".replace(" ", "_"),
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
        type="primary",
    )
