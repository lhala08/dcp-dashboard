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
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="collapsed",
)

st.markdown(
    """
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800&display=swap');

    :root {
        color-scheme: light !important;
    }
    html, body, [class*="css"] {
        font-family: 'Plus Jakarta Sans', sans-serif;
        color: #0f172a;
    }
    .stApp {
        background-color: #f8fafc !important;
        color: #0f172a !important;
    }
    .block-container {
        padding-top: 1.2rem;
        padding-bottom: 2.5rem;
        max-width: 1480px;
    }
    /* Force light styling on selectboxes, labels, and buttons even on dark-mode browsers */
    label[data-testid="stWidgetLabel"] p,
    .stSelectbox label,
    .stRadio label,
    .stCheckbox label {
        color: #334155 !important;
        font-weight: 600 !important;
    }
    div[data-baseweb="select"] > div {
        background-color: #ffffff !important;
        color: #0f172a !important;
        border-color: #cbd5e1 !important;
    }
    button[kind="secondary"] {
        background-color: #ffffff !important;
        color: #0f172a !important;
        border: 1px solid #cbd5e1 !important;
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
        padding: 1rem 1.05rem;
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
        margin-bottom: 0.4rem;
    }
    .kpi-label {
        font-size: 0.72rem;
        font-weight: 700;
        text-transform: uppercase;
        letter-spacing: 0.04em;
        color: #64748b;
    }
    .kpi-pill {
        font-size: 0.69rem;
        font-weight: 700;
        padding: 0.16rem 0.48rem;
        border-radius: 999px;
    }
    .kpi-number {
        font-size: 1.58rem;
        font-weight: 800;
        color: #0f172a;
        line-height: 1.15;
        letter-spacing: -0.02em;
    }
    .kpi-foot {
        font-size: 0.73rem;
        color: #94a3b8;
        margin-top: 0.3rem;
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
        color: #475569 !important;
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
# CREDENTIALS & PASSWORD GATE (PRESERVED)
# ==========================================
DEFAULT_SHEET_ID = "1uGsxNYJHhL0DOfODQbn60L2wXnC2qd-lKO56O8OzF4g"
DEFAULT_APP_PASSWORD = "!dcp2026*"


def get_secret(key: str, default: str = "") -> str:
    try:
        val = st.secrets.get(key, default)
        return str(val) if val else default
    except Exception:
        return default


SHEET_ID = get_secret("SHEET_ID", DEFAULT_SHEET_ID)
APP_PASSWORD = get_secret("APP_PASSWORD", DEFAULT_APP_PASSWORD)

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
    "Delivered (Pending Docs)": "#f59e0b",
    "Ongoing Delivery": "#3b82f6",
    "Successful Rebid": "#8b5cf6",
}


# ==========================================
# DATA LOADING & SDO NORMALIZATION
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


def normalize_division_name(raw_div: str, region_str: str) -> tuple[str, bool]:
    val = re.sub(r"\s+", " ", str(raw_div or "")).strip()
    reg = str(region_str or "").upper().strip()
    if not val:
        return ("UNSPECIFIED DIVISION", False)

    up = val.upper().replace(".", " ")
    up = re.sub(r"\s+", " ", up).strip()

    ro_patterns = [
        r"^(DEPED\s+)?(RO|R\s*O)[\s\-]*([IVX0-9]+[\s\-]*[AB]?|CAR|CARAGA|NCR|NIR|BARMM|MIMAROPA)\b",
        r"^(DEPED\s+)?REGION[\s\-]*([IVX0-9]+[\s\-]*[AB]?|CAR|CARAGA|NCR|NIR|BARMM|MIMAROPA)\b",
        r"REGIONAL\s+OFFICE",
    ]
    if any(re.search(pat, up) for pat in ro_patterns):
        return (up, False)

    up = re.sub(
        r"^(DEPED\s+)?(CITY\s+SCHOOLS\s+DIVISION\s+OF|SCHOOLS\s+DIVISION\s+OF|DIVISION\s+OF|SDO[\s\-]+)\s*",
        "",
        up,
    ).strip()

    up = re.sub(r"\bSTO\b", "SANTO", up)
    up = re.sub(r"\bSTA\b", "SANTA", up)
    if up.startswith("CITY OF "):
        up = up[len("CITY OF "):].strip() + " CITY"
    up = re.sub(r"\s+", " ", up).strip()

    if up in {"SAN CARLOS CITY", "TALISAY CITY"} and reg:
        up = f"{up} - R{reg}"

    return (up if up else "UNSPECIFIED DIVISION", True)


def find_col_by_keywords(columns, keywords) -> str:
    for col in columns:
        c_low = str(col).lower().strip()
        if any(k in c_low for k in keywords):
            return col
    return ""


def raw_status_priority(status_text: str) -> int:
    low = str(status_text or "").lower().strip()
    if not low:
        return 0
    if "complete" in low:
        return 3
    if "delivered" in low or "pending mov" in low or "pending dr" in low or "pending iar" in low:
        return 2
    return 1


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

    raw_status_by_sid_lot = {}
    raw_status_by_sid = {}
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

        lot_col = find_col_by_keywords(raw_cols, ["lot"])
        raw_div_col = find_col_by_keywords(raw_cols, ["division", "sdo"])

        if site_col:
            site_codes = clean_series(raw_df[site_col])
            lot_vals = clean_series(raw_df[lot_col]) if lot_col else pd.Series([""] * len(raw_df))

            if stat_col:
                cur_statuses = clean_series(raw_df[stat_col])
                for sid, lot_v, st_val in zip(site_codes, lot_vals, cur_statuses):
                    if sid and st_val:
                        new_p = raw_status_priority(st_val)
                        if new_p >= raw_status_priority(raw_status_by_sid.get(sid, "")):
                            raw_status_by_sid[sid] = st_val
                        if lot_v:
                            key_sl = (sid, lot_v)
                            if new_p >= raw_status_priority(raw_status_by_sid_lot.get(key_sl, "")):
                                raw_status_by_sid_lot[key_sl] = st_val

            if raw_div_col:
                cur_divs = clean_series(raw_df[raw_div_col])
                for sid, d_val in zip(site_codes, cur_divs):
                    if sid and d_val:
                        raw_div_lookup[sid] = d_val

    mapped_statuses = []
    for sid, lot_v, col_n in zip(df["School_ID"], df["LOT"], df["Col_N_Status"]):
        cand_lot = raw_status_by_sid_lot.get((sid, lot_v), "") if (sid and lot_v) else ""
        cand_sid = raw_status_by_sid.get(sid, "") if sid else ""
        best = max([cand_lot, cand_sid, col_n], key=raw_status_priority)
        mapped_statuses.append(best)

    df["Effective_Raw_Status"] = mapped_statuses

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

    norm_results = [normalize_division_name(x, r) for x, r in zip(div_resolved, df["Region_Clean"])]
    df["Division_Clean"] = [r[0] for r in norm_results]
    df["Is_SDO"] = [r[1] for r in norm_results]

    is_nonteaching = df["Package_Lower"] == "laptop for non-teaching"
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


def classify_delivery_stages(df_in: pd.DataFrame) -> pd.DataFrame:
    df = df_in.copy()
    df["Has_DR"] = has_valid_entry(df["DR_Date"])
    df["Has_IAR"] = has_valid_entry(df["IAR_Date"])

    eff_lower = df["Effective_Raw_Status"].str.lower().str.strip()
    orig_status_lower = df["Status"].str.lower().str.strip()

    is_completed = (
        (df["Has_DR"] & df["Has_IAR"])
        | eff_lower.str.contains("complete", regex=False, na=False)
        | orig_status_lower.str.contains("complete", regex=False, na=False)
    )

    is_delivered_pending_docs = (~is_completed) & (
        df["Has_DR"]
        | df["Has_IAR"]
        | eff_lower.str.contains("delivered|pending mov|pending dr|pending iar", regex=True, na=False)
        | orig_status_lower.str.contains("delivered|pending mov|pending dr|pending iar", regex=True, na=False)
    )

    is_rebid = (~is_completed) & (~is_delivered_pending_docs) & (
        orig_status_lower.str.match(r"^(rebid|for rebid|successful rebid)$", na=False)
    )

    conditions = [
        is_completed,
        is_delivered_pending_docs,
        is_rebid,
    ]
    choices = ["Completed", "Delivered (Pending Docs)", "Successful Rebid"]
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
    div_display: pd.DataFrame,
    pkg_display: pd.DataFrame,
    comp_display: pd.DataFrame,
) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        reg_display.to_excel(writer, sheet_name="Regional_Summary", index=False)
        if not div_display.empty:
            div_display.to_excel(writer, sheet_name="Division_Billing_Status", index=False)
        pkg_display.to_excel(writer, sheet_name="Package_Distribution", index=False)
        if not comp_display.empty:
            comp_display.to_excel(writer, sheet_name="Component_Breakdown", index=False)

        header_fill = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")
        header_font = Font(color="FFFFFF", bold=True, size=11)
        total_fill = PatternFill(start_color="EFF6FF", end_color="EFF6FF", fill_type="solid")
        total_font = Font(bold=True, size=11)
        ready_fill = PatternFill(start_color="DCFCE7", end_color="DCFCE7", fill_type="solid")
        pending_fill = PatternFill(start_color="FEF3C7", end_color="FEF3C7", fill_type="solid")

        for sheet_name in writer.sheets:
            ws = writer.sheets[sheet_name]
            for cell in ws[1]:
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = Alignment(horizontal="center", vertical="center")

            if sheet_name == "Division_Billing_Status":
                for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
                    row_vals = [str(c.value or "") for c in row]
                    if any("READY FOR BILLING" in v for v in row_vals):
                        for c in row:
                            c.fill = ready_fill
                    elif any("PENDING DOCS" in v for v in row_vals):
                        for c in row:
                            c.fill = pending_fill

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
    "include_ro_in_div": False,
    "stage_overrides": {},
    "dr_overrides": {},
    "iar_overrides": {},
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
    source_mode = st.radio(
        "Source Mode",
        ["Live Google Sheet", "Upload Excel (.xlsx - 100% Private)"],
        index=0,
    )

    excel_data = None
    if source_mode == "Live Google Sheet":
        if st.button("🔄 Sync Latest Google Sheet", use_container_width=True, type="primary"):
            st.cache_data.clear()
            st.session_state["stage_overrides"] = {}
            st.session_state["dr_overrides"] = {}
            st.session_state["iar_overrides"] = {}
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

df = classify_delivery_stages(df)

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
    ("Division", "sel_division"),
    ("Package", "sel_package"),
    ("Status", "sel_stage"),
]:
    if st.session_state[key] != "All":
        active_badges.append(f"{label}: <b>{st.session_state[key]}</b>")

subtitle_html = (
    " &nbsp;•&nbsp; ".join(active_badges)
    if active_badges
    else "All Funds &nbsp;•&nbsp; All Regions &nbsp;•&nbsp; All Divisions &nbsp;•&nbsp; Click any chart or slicer to filter"
)

total_sim_edits = (
    len(st.session_state["stage_overrides"])
    + len(st.session_state["dr_overrides"])
    + len(st.session_state["iar_overrides"])
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

div_source_df = df if st.session_state["include_ro_in_div"] else df[df["Is_SDO"]]
if st.session_state["sel_region"] != "All":
    div_pool = div_source_df[div_source_df["Region_Clean"] == st.session_state["sel_region"]]["Division_Clean"].unique().tolist()
else:
    div_pool = div_source_df["Division_Clean"].unique().tolist()
division_options = ["All"] + sorted(div_pool)

if st.session_state["sel_division"] not in division_options:
    st.session_state["sel_division"] = "All"

pkg_options = ["All"] + sorted(df["Package_Display"].unique().tolist())
stage_options = ["All", "Completed", "Delivered (Pending Docs)", "Ongoing Delivery", "Successful Rebid"]

s1, s2, s3, s4, s5, s6 = st.columns([1.1, 1.0, 1.45, 1.35, 1.35, 0.85])
with s1:
    st.selectbox("Fund Source", fund_options, key="sel_fund")
with s2:
    st.selectbox("Region", region_options, key="sel_region")
with s3:
    st.selectbox("Division (SDO)", division_options, key="sel_division")
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
# KPI & DIVISION (SDO) BILLING COMPUTATIONS
# ==========================================
unique_rec = filtered.drop_duplicates(subset=["Fund_Clean", "Recipient_Key"], keep="last").copy()
assigned_region_map = unique_rec.set_index(["Fund_Clean", "Recipient_Key"])["Region_Clean"]

stage_priority = {
    "Completed": 4,
    "Delivered (Pending Docs)": 3,
    "Successful Rebid": 2,
    "Ongoing Delivery": 1,
}
filtered["Stage_Rank"] = filtered["Delivery_Stage"].map(stage_priority).fillna(0)
rec_stage = (
    filtered.sort_values("Stage_Rank")
    .drop_duplicates(subset=["Fund_Clean", "Recipient_Key"], keep="last")
    .copy()
)
rec_stage["Region_Clean"] = rec_stage.set_index(["Fund_Clean", "Recipient_Key"]).index.map(assigned_region_map)

div_eval_df = df.copy()
if not st.session_state["include_ro_in_div"]:
    div_eval_df = div_eval_df[div_eval_df["Is_SDO"]].copy()
if st.session_state["sel_fund"] != "All":
    div_eval_df = div_eval_df[div_eval_df["Fund_Clean"] == st.session_state["sel_fund"]]
if st.session_state["sel_region"] != "All":
    div_eval_df = div_eval_df[div_eval_df["Region_Clean"] == st.session_state["sel_region"]]
if st.session_state["sel_division"] != "All":
    div_eval_df = div_eval_df[div_eval_df["Division_Clean"] == st.session_state["sel_division"]]
if st.session_state["sel_package"] != "All":
    div_eval_df = div_eval_df[div_eval_df["Package_Display"] == st.session_state["sel_package"]]

div_eval_df["Stage_Rank"] = div_eval_df["Delivery_Stage"].map(stage_priority).fillna(0)

dr_by_rec = div_eval_df.groupby(["Fund_Clean", "Recipient_Key"])["Has_DR"].any()
iar_by_rec = div_eval_df.groupby(["Fund_Clean", "Recipient_Key"])["Has_IAR"].any()

div_rec_unique = (
    div_eval_df.sort_values("Stage_Rank")
    .drop_duplicates(subset=["Fund_Clean", "Recipient_Key"], keep="last")
    .copy()
)
div_rec_unique["Rec_Has_DR"] = div_rec_unique.set_index(["Fund_Clean", "Recipient_Key"]).index.map(dr_by_rec).fillna(False)
div_rec_unique["Rec_Has_IAR"] = div_rec_unique.set_index(["Fund_Clean", "Recipient_Key"]).index.map(iar_by_rec).fillna(False)

rem_by_div = compute_remaining_packages(div_eval_df, ["Division_Clean"])

div_rows = []
for div_name, grp in div_rec_unique.groupby("Division_Clean", sort=False):
    reg_d = grp["Region_Clean"].mode().iloc[0] if not grp["Region_Clean"].mode().empty else grp["Region_Clean"].iloc[0]
    tot_d = len(grp)
    comp_d = int((grp["Delivery_Stage"] == "Completed").sum())
    deliv_pend_d = int((grp["Delivery_Stage"] == "Delivered (Pending Docs)").sum())
    ong_d = int((grp["Delivery_Stage"] == "Ongoing Delivery").sum())
    reb_d = int((grp["Delivery_Stage"] == "Successful Rebid").sum())

    dr_cnt = int(grp["Rec_Has_DR"].sum())
    iar_cnt = int(grp["Rec_Has_IAR"].sum())

    comp_rate = (comp_d / tot_d * 100) if tot_d > 0 else 0.0

    if tot_d > 0 and comp_d == tot_d:
        pay_status = "✅ READY FOR BILLING"
    elif tot_d > 0 and (comp_d + deliv_pend_d) == tot_d:
        pay_status = "🟡 PENDING DOCS"
    else:
        pay_status = "❌ INCOMPLETE DELIVERY"

    div_rows.append({
        "Region": reg_d,
        "Division (SDO)": div_name,
        "Total Recipients": tot_d,
        "Remaining Packages": int(rem_by_div.get(div_name, 0)),
        "Ongoing Delivery": ong_d,
        "Successful Rebid": reb_d,
        "Delivered (Pending Docs)": deliv_pend_d,
        "Completed": comp_d,
        "With DR Date": dr_cnt,
        "With IAR Date": iar_cnt,
        "Completion (%)": round(comp_rate, 1),
        "Billing Status": pay_status,
    })

div_df = pd.DataFrame(div_rows)
if not div_df.empty:
    div_df["_reg_ord"] = div_df["Region"].map(lambda r: REGION_ORDER.get(r, 999))
    div_df = div_df.sort_values(
        by=["Completion (%)", "_reg_ord", "Division (SDO)"],
        ascending=[False, True, True],
    ).drop(columns=["_reg_ord"])

divs_ready_cnt = int((div_df["Billing Status"] == "✅ READY FOR BILLING").sum()) if not div_df.empty else 0
divs_total_cnt = len(div_df)

total_recipients = len(unique_rec)
total_rem_pkgs = int(compute_remaining_packages(filtered, ["Fund_Clean"]).sum())
cnt_ongoing = int((rec_stage["Delivery_Stage"] == "Ongoing Delivery").sum())
cnt_rebid = int((rec_stage["Delivery_Stage"] == "Successful Rebid").sum())
cnt_deliv_pend = int((rec_stage["Delivery_Stage"] == "Delivered (Pending Docs)").sum())
cnt_comp = int((rec_stage["Delivery_Stage"] == "Completed").sum())

pct_comp = (cnt_comp / total_recipients * 100) if total_recipients > 0 else 0.0
pct_deliv_pend = (cnt_deliv_pend / total_recipients * 100) if total_recipients > 0 else 0.0
pct_ongoing = (cnt_ongoing / total_recipients * 100) if total_recipients > 0 else 0.0

cols = st.columns(7)
kpi_specs = [
    ("Total Recipients", f"{total_recipients:,}", "100%", "#eff6ff", "#1d4ed8", "Unique schools/offices"),
    ("Remaining Pkgs", f"{total_rem_pkgs:,}", "Pending", "#f0f9ff", "#0284c7", "Undelivered packages"),
    ("Ongoing Delivery", f"{cnt_ongoing:,}", f"{pct_ongoing:.1f}%", "#eff6ff", "#2563eb", "Awarded / in transit"),
    ("Successful Rebid", f"{cnt_rebid:,}", "Rebid", "#f5f3ff", "#7c3aed", "Rebid allocation"),
    ("Delivered (Pending Docs)", f"{cnt_deliv_pend:,}", f"{pct_deliv_pend:.1f}%", "#fef3c7", "#b45309", "Awaiting DR/IAR validation"),
    ("Completed", f"{cnt_comp:,}", f"{pct_comp:.1f}%", "#ecfdf5", "#059669", "Delivered & docs validated"),
    ("Divisions Billable", f"{divs_ready_cnt:,} / {divs_total_cnt:,}", "100% Comp.", "#dcfce7", "#15803d", "SDOs ready for billing"),
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

st.markdown("<div style='height: 0.9rem;'></div>", unsafe_allow_html=True)

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
    deliv_r = int((r_sub["Delivery_Stage"] == "Delivered (Pending Docs)").sum())
    comp_rate = (comp_r / tot_r * 100) if tot_r > 0 else 0.0

    reg_rows.append({
        "Region": str(reg),
        "Total Recipients": tot_r,
        "Remaining Packages": int(rem_by_reg.get(reg, 0)),
        "Ongoing Delivery": int((r_sub["Delivery_Stage"] == "Ongoing Delivery").sum()),
        "Successful Rebid": int((r_sub["Delivery_Stage"] == "Successful Rebid").sum()),
        "Delivered (Pending Docs)": deliv_r,
        "Completed": comp_r,
        "Completion (%)": round(comp_rate, 1),
    })

reg_df = pd.DataFrame(reg_rows)
reg_total = pd.DataFrame([{
    "Region": "TOTAL",
    "Total Recipients": reg_df["Total Recipients"].sum(),
    "Remaining Packages": reg_df["Remaining Packages"].sum(),
    "Ongoing Delivery": reg_df["Ongoing Delivery"].sum(),
    "Successful Rebid": reg_df["Successful Rebid"].sum(),
    "Delivered (Pending Docs)": reg_df["Delivered (Pending Docs)"].sum(),
    "Completed": reg_df["Completed"].sum(),
    "Completion (%)": round(pct_comp, 1),
}])
reg_display = pd.concat([reg_df, reg_total], ignore_index=True)

if not div_df.empty:
    tot_r_div = div_df["Total Recipients"].sum()
    div_total = pd.DataFrame([{
        "Region": "TOTAL",
        "Division (SDO)": f"{divs_ready_cnt} Ready for Billing (of {divs_total_cnt} SDOs)",
        "Total Recipients": tot_r_div,
        "Remaining Packages": div_df["Remaining Packages"].sum(),
        "Ongoing Delivery": div_df["Ongoing Delivery"].sum(),
        "Successful Rebid": div_df["Successful Rebid"].sum(),
        "Delivered (Pending Docs)": div_df["Delivered (Pending Docs)"].sum(),
        "Completed": div_df["Completed"].sum(),
        "With DR Date": div_df["With DR Date"].sum(),
        "With IAR Date": div_df["With IAR Date"].sum(),
        "Completion (%)": round(div_df["Completed"].sum() / tot_r_div * 100, 1) if tot_r_div > 0 else 0.0,
        "Billing Status": f"{divs_ready_cnt} READY",
    }])
    div_display = pd.concat([div_df, div_total], ignore_index=True)
else:
    div_display = pd.DataFrame()

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
    "💰 Division (SDO) Billing Status",
    "🗺️ Regional Table",
    "📦 Packages & Components",
    "🔎 Lookup & Status Simulator",
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
            value_vars=["Completed", "Delivered (Pending Docs)", "Ongoing Delivery", "Successful Rebid"],
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
        fig_bar.update_xaxes(type="category", title="", showgrid=False, tickfont=dict(color="#334155"))
        fig_bar.update_layout(
            plot_bgcolor="rgba(0,0,0,0)",
            paper_bgcolor="rgba(0,0,0,0)",
            font=dict(color="#0f172a"),
            margin=dict(l=10, r=10, t=15, b=10),
            legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="left", x=0, title_text=""),
            yaxis=dict(title="Recipients", gridcolor="#e2e8f0", tickfont=dict(color="#334155")),
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
            "Stage": ["Completed", "Delivered (Pending Docs)", "Ongoing Delivery", "Successful Rebid"],
            "Count": [cnt_comp, cnt_deliv_pend, cnt_ongoing, cnt_rebid],
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
                    textfont=dict(color="#0f172a"),
                    hovertemplate="<b>%{label}</b><br>Recipients: %{value:,}<br>Share: %{percent}<extra></extra>",
                )
            ]
        )
        fig_donut.update_layout(
            height=390,
            plot_bgcolor="rgba(0,0,0,0)",
            paper_bgcolor="rgba(0,0,0,0)",
            font=dict(color="#0f172a"),
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
        st.markdown('<div class="section-header">📈 Delivery Completion (%) per Region</div>', unsafe_allow_html=True)
        rate_sorted = reg_df.sort_values("Completion (%)", ascending=True).copy()
        rate_sorted["Region"] = rate_sorted["Region"].astype(str)
        rate_sorted["Label_Text"] = rate_sorted.apply(
            lambda r: f"{r['Completion (%)']:.1f}% ({int(r['Completed']):,}/{int(r['Total Recipients']):,})", axis=1
        )
        fig_rate = px.bar(
            rate_sorted,
            x="Completion (%)",
            y="Region",
            orientation="h",
            text="Label_Text",
            color="Completion (%)",
            color_continuous_scale=["#93c5fd", "#2563eb", "#10b981"],
            range_x=[0, 125],
            height=max(380, len(rate_sorted) * 26),
        )
        fig_rate.update_yaxes(type="category", title="", tickfont=dict(color="#334155"))
        fig_rate.update_traces(textposition="outside", textfont=dict(color="#0f172a"))
        fig_rate.update_layout(
            plot_bgcolor="rgba(0,0,0,0)",
            paper_bgcolor="rgba(0,0,0,0)",
            font=dict(color="#0f172a"),
            coloraxis_showscale=False,
            margin=dict(l=10, r=30, t=10, b=10),
            xaxis=dict(title="Completed (%)", gridcolor="#e2e8f0", tickfont=dict(color="#334155")),
        )
        st.plotly_chart(fig_rate, use_container_width=True)

    with c_bot2:
        st.markdown('<div class="section-header">📦 Remaining Packages by DCP Package</div>', unsafe_allow_html=True)
        pkg_sorted = pkg_df.sort_values("Remaining Packages", ascending=True)
        max_rem = pkg_sorted["Remaining Packages"].max() if not pkg_sorted.empty else 100
        fig_pkg = px.bar(
            pkg_sorted,
            x="Remaining Packages",
            y="DCP Package",
            orientation="h",
            text="Remaining Packages",
            color_discrete_sequence=["#1d4ed8"],
            range_x=[0, max_rem * 1.22 if max_rem > 0 else 10],
            height=max(380, len(rate_sorted) * 26),
        )
        fig_pkg.update_traces(texttemplate="%{text:,}", textposition="outside", textfont=dict(color="#0f172a"))
        fig_pkg.update_layout(
            plot_bgcolor="rgba(0,0,0,0)",
            paper_bgcolor="rgba(0,0,0,0)",
            font=dict(color="#0f172a"),
            margin=dict(l=10, r=40, t=10, b=10),
            xaxis=dict(title="Remaining Packages", gridcolor="#e2e8f0", tickfont=dict(color="#334155")),
            yaxis=dict(title="", tickfont=dict(color="#334155")),
        )
        st.plotly_chart(fig_pkg, use_container_width=True)

with tab2:
    st.markdown('<div class="section-header">💰 Schools Division (SDO) Billing Status</div>', unsafe_allow_html=True)
    st.caption(
        f"**{divs_total_cnt:,} unique Division Offices** sa dataset • "
        f"**{divs_ready_cnt:,} Ready for Billing** (100% Completed — delivered at validated na ang docs)."
    )

    opt_c1, opt_c2 = st.columns([2.4, 1.0])
    with opt_c1:
        pay_filter = st.radio(
            "Filter Billing Status",
            ["All", "✅ READY FOR BILLING", "🟡 PENDING DOCS", "❌ INCOMPLETE DELIVERY"],
            horizontal=True,
        )
    with opt_c2:
        st.markdown("<div style='height: 1.55rem;'></div>", unsafe_allow_html=True)
        st.checkbox(
            "Include Regional Offices (RO)",
            key="include_ro_in_div",
            help="Naka-uncheck by default para pure Schools Division Offices (SDO) lang ang bilangin.",
        )

    div_view = div_display.copy()
    if not div_view.empty and pay_filter != "All":
        div_view = div_view[div_view["Billing Status"] == pay_filter]

    if not div_view.empty:
        dyn_height = min(620, max(120, (len(div_view) + 1) * 36 + 10))
        st.dataframe(
            div_view.style.format({
                "Total Recipients": "{:,.0f}",
                "Remaining Packages": "{:,.0f}",
                "Ongoing Delivery": "{:,.0f}",
                "Successful Rebid": "{:,.0f}",
                "Delivered (Pending Docs)": "{:,.0f}",
                "Completed": "{:,.0f}",
                "With DR Date": "{:,.0f}",
                "With IAR Date": "{:,.0f}",
                "Completion (%)": "{:.1f}%",
            }),
            use_container_width=True,
            hide_index=True,
            height=dyn_height,
            column_config={
                "Region": st.column_config.TextColumn("Region", pinned=True, width="small"),
                "Division (SDO)": st.column_config.TextColumn("Division (SDO)", pinned=True, width="medium"),
                "Total Recipients": st.column_config.NumberColumn("Recipients", width="small"),
                "Remaining Packages": st.column_config.NumberColumn("Rem. Pkgs", width="small"),
                "Ongoing Delivery": st.column_config.NumberColumn("Ongoing", width="small"),
                "Successful Rebid": st.column_config.NumberColumn("Rebid", width="small"),
                "Delivered (Pending Docs)": st.column_config.NumberColumn("Delivered (Pending Docs)", width="medium"),
                "Completed": st.column_config.NumberColumn("Completed", width="small"),
                "With DR Date": st.column_config.NumberColumn("DR Date", width="small"),
                "With IAR Date": st.column_config.NumberColumn("IAR Date", width="small"),
                "Completion (%)": st.column_config.NumberColumn("Comp. %", width="small"),
                "Billing Status": st.column_config.TextColumn("Billing Status", width="medium"),
            },
        )
    else:
        st.info("No divisions match the selected billing status filter.")

with tab3:
    st.markdown('<div class="section-header">🗺️ Regional Delivery Monitoring Breakdown</div>', unsafe_allow_html=True)
    st.dataframe(
        reg_display.style.format({
            "Total Recipients": "{:,.0f}",
            "Remaining Packages": "{:,.0f}",
            "Ongoing Delivery": "{:,.0f}",
            "Successful Rebid": "{:,.0f}",
            "Delivered (Pending Docs)": "{:,.0f}",
            "Completed": "{:,.0f}",
            "Completion (%)": "{:.1f}%",
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
    st.markdown('<div class="section-header">🔎 School / Division Lookup & Status Simulator</div>', unsafe_allow_html=True)
    st.caption(
        "Mag-search ng kahit anong School ID, Division (SDO), LOT, o Supplier. Puwede ring mag-simulate ng status updates dito nang hindi nababago ang live Google Sheet."
    )

    s_left, s_right = st.columns([3, 1])
    with s_left:
        search_q = st.text_input(
            "Search School ID, Division (SDO), Recipient Office, LOT, or Supplier",
            placeholder="e.g. 502655, BACOLOD CITY, SAN CARLOS CITY, or Columbia Technologies...",
        )
    with s_right:
        st.markdown("<div style='height: 1.75rem;'></div>", unsafe_allow_html=True)
        if st.button("↩️ Reset Simulated Edits", use_container_width=True):
            st.session_state["stage_overrides"] = {}
            st.session_state["dr_overrides"] = {}
            st.session_state["iar_overrides"] = {}
            st.rerun()

    lookup_cols = [
        "Row_ID", "Fund_Clean", "LOT", "Region_Clean", "Division_Clean",
        "School_ID", "Package_Display", "Pkg_Count", "DR_Date", "IAR_Date", "Delivery_Stage"
    ]
    lookup_df = filtered[lookup_cols].copy()
    lookup_df.columns = [
        "Row_ID", "Fund", "LOT", "Region", "Division (SDO)",
        "School ID", "DCP Package", "Packages", "DR Date", "IAR Date", "Status"
    ]
    if search_q.strip():
        q = search_q.strip().lower()
        mask = lookup_df.drop(columns=["Row_ID"]).apply(
            lambda col: col.astype(str).str.lower().str.contains(q, na=False)
        ).any(axis=1)
        lookup_df = lookup_df[mask]

    st.caption(f"Showing {len(lookup_df):,} records (double-click DR Date, IAR Date, or Status to simulate)")
    edited_df = st.data_editor(
        lookup_df.head(500),
        use_container_width=True,
        hide_index=True,
        height=520,
        disabled=["Row_ID", "Fund", "LOT", "Region", "Division (SDO)", "School ID", "DCP Package", "Packages"],
        column_config={
            "Row_ID": None,
            "DR Date": st.column_config.TextColumn("DR Date (Editable)"),
            "IAR Date": st.column_config.TextColumn("IAR Date (Editable)"),
            "Status": st.column_config.SelectboxColumn(
                "Status (Editable)",
                options=["Completed", "Delivered (Pending Docs)", "Ongoing Delivery", "Successful Rebid"],
                required=True,
            ),
        },
        key="school_status_editor",
    )

    orig_slice = lookup_df.head(500)
    orig_dr = dict(zip(orig_slice["Row_ID"], orig_slice["DR Date"]))
    orig_iar = dict(zip(orig_slice["Row_ID"], orig_slice["IAR Date"]))
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
        if str(row["Status"]) != str(orig_st.get(rid, "")):
            st.session_state["stage_overrides"][rid] = str(row["Status"])
            changed = True

    if changed:
        st.rerun()

# ==========================================
# EXPORT BAR
# ==========================================
excel_bytes = build_styled_excel(reg_display, div_display, pkg_display, comp_display)

st.markdown("<hr style='border:0; border-top:1px solid #e2e8f0; margin: 1.5rem 0 0.8rem 0;'>", unsafe_allow_html=True)
dl_c1, dl_c2 = st.columns([3, 1])
with dl_c1:
    st.caption("Kasama sa Excel export ang **Regional Summary**, **Division Billing Status**, **Package Distribution**, at **Component Breakdown**.")
with dl_c2:
    st.download_button(
        label="📥 Export Full Excel Report (.xlsx)",
        data=excel_bytes,
        file_name=f"DCP_Executive_Billing_Report_{st.session_state['sel_fund']}_{st.session_state['sel_region']}.xlsx".replace(" ", "_"),
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
        type="primary",
    )
