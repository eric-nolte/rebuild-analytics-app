
import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
from datetime import datetime
from pathlib import Path
import base64
import json
import re
import urllib.request

import altair as alt
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

APP_VERSION = "V16.2.3"
APP_LAST_UPDATED = "2026-06-30"
METHODOLOGY_VERSION = "2026.06-Cost-IQR-v2"
OUTLIER_RULE_VERSION = "Cost Log-IQR by Machine + CCR TYPE"
DEALER_RATE_VERSION = "Built-in Expanded Dealer Rates 2016-2026"
SECURITY_CONTROL_VERSION = "Phase 1 Security Controls"
EXPORT_FORMAT_VERSION = "V16.2.3 No Removed Label Export"
CONFIDENTIALITY_LABEL = ""
MAX_UPLOAD_MB = 50
DEFAULT_MAX_ROWS_WARNING = 25000

st.set_page_config(layout="wide", page_title="Rebuild Analytics Platform")

# =====================================================
# CAT-INSPIRED BRANDING / THEME
# =====================================================
def find_logo_file():
    for name in ["cat_logo.png", "caterpillar_logo.png", "CAT_logo.png", "logo.png"]:
        path = Path(name)
        if path.exists():
            return path
    return None


def img_to_base64(path):
    try:
        return base64.b64encode(path.read_bytes()).decode("utf-8")
    except Exception:
        return None


def inject_cat_theme():
    st.markdown(
        """
        <style>
        :root { --cat-yellow:#FFC500; --cat-black:#000000; --cat-charcoal:#1F1F1F; --cat-gray:#7A7A7A; --cat-light:#F4F4F4; --cat-border:#D9D9D9; }
        .stApp { background: var(--cat-light); color: var(--cat-charcoal); }
        .block-container { padding-top: 1rem; padding-left: 2rem; padding-right: 2rem; max-width: 1550px; }
        .cat-header { background: linear-gradient(90deg,#000000 0%,#1F1F1F 72%,#FFC500 72%,#FFC500 100%); padding: 22px 28px; border-radius: 4px; margin-bottom: 20px; box-shadow: 0 4px 14px rgba(0,0,0,.22); border-left: 8px solid #FFC500; display: flex; justify-content: space-between; align-items: center; gap: 24px; }
        .cat-header-title { color: #FFFFFF; font-size: 34px; font-weight: 900; letter-spacing: .6px; line-height: 1.05; margin: 0; }
        .cat-header-subtitle { color: #D9D9D9; font-size: 14px; margin-top: 9px; font-weight: 600; }
        .cat-version-pill { display: inline-block; background: #FFC500; color: #000000; font-weight: 900; padding: 5px 12px; border-radius: 2px; margin-top: 12px; font-size: 12px; }
        .cat-logo-box { background:#FFC500; color:#000000; font-weight: 1000; font-size: 34px; letter-spacing: -1px; padding: 12px 18px; border: 3px solid #000000; min-width: 95px; text-align:center; }
        .cat-logo-img { max-height: 72px; max-width: 210px; background:#FFFFFF; padding:6px; border-radius:3px; border:2px solid #000000; }
        .method-card { background:#FFFFFF; border-left:8px solid #FFC500; padding:15px 18px; margin:12px 0 18px 0; box-shadow:0 2px 8px rgba(0,0,0,.08); border-radius:4px; }
        .method-title { font-size:18px; font-weight:900; color:#000000; margin-bottom:5px; }
        .method-body { font-size:14px; color:#333333; line-height:1.45; }
        h1,h2,h3 { color:#1F1F1F; font-weight:850; }
        h2 { border-left:6px solid #FFC500; padding-left:10px; }
        [data-testid="stMetric"] { background:#FFFFFF; border:1px solid #D9D9D9; border-top:5px solid #FFC500; border-radius:4px; padding:15px 17px; box-shadow:0 2px 8px rgba(0,0,0,.08); }
        [data-testid="stMetricLabel"] { color:#4A4A4A; font-weight:800; }
        [data-testid="stMetricValue"] { color:#000000; font-weight:950; }
        button[data-baseweb="tab"] { background:#FFFFFF; color:#1F1F1F; border-radius:2px 2px 0 0; font-weight:800; border:1px solid #D9D9D9; margin-right:4px; padding:10px 14px; }
        button[data-baseweb="tab"][aria-selected="true"] { background:#000000; color:#FFC500; border-bottom:4px solid #FFC500; }
        .stButton > button { background:#FFC500; color:#000000; border:2px solid #000000; border-radius:3px; font-weight:900; }
        .stButton > button:hover { background:#000000; color:#FFC500; border:2px solid #FFC500; }
        .stDownloadButton > button { background:#000000; color:#FFC500; border:2px solid #FFC500; border-radius:3px; font-weight:900; }
        .stDownloadButton > button:hover { background:#FFC500; color:#000000; border:2px solid #000000; }
        label { font-weight:800 !important; color:#1F1F1F !important; }
        div[data-testid="stAlert"] { border-radius:4px; border-left:6px solid #FFC500; }
        [data-testid="stDataFrame"] { background:#FFFFFF; border:1px solid #D9D9D9; border-radius:4px; box-shadow:0 2px 8px rgba(0,0,0,.06); }
        details { background:#FFFFFF !important; border:1px solid #D9D9D9 !important; border-radius:4px !important; }
        summary { font-weight:900 !important; color:#111111 !important; }
        [data-testid="stVegaLiteChart"] { background:#FFFFFF; border:1px solid #D9D9D9; border-radius:4px; padding:10px; box-shadow:0 2px 8px rgba(0,0,0,.06); }
        div[data-baseweb="select"] > div { border:2px solid #1F1F1F !important; border-radius:3px !important; background:#FFFFFF !important; }
        div[data-baseweb="select"]:focus-within > div { border-color:#FFC500 !important; box-shadow:0 0 0 2px rgba(255,197,0,.35) !important; }
        span[data-baseweb="tag"], div[data-baseweb="tag"] { background:#FFC500 !important; color:#000000 !important; border:1px solid #000000 !important; font-weight:800 !important; }
        span[data-baseweb="tag"] span, div[data-baseweb="tag"] span { color:#000000 !important; font-weight:800 !important; }
        div[role="option"][aria-selected="true"], ul[role="listbox"] li[aria-selected="true"] { background:#FFC500 !important; color:#000000 !important; font-weight:900 !important; }
        div[role="option"]:hover, ul[role="listbox"] li:hover { background:#1F1F1F !important; color:#FFC500 !important; }
        div[data-testid="stCheckbox"] [data-baseweb="checkbox"] > div:first-child { border-color:#000000 !important; }
        div[data-testid="stCheckbox"] input:checked + div { background-color:#FFC500 !important; border-color:#000000 !important; color:#000000 !important; }
        .cat-footer { margin-top:35px; padding:15px; background:#1F1F1F; color:#D9D9D9; font-size:12px; border-top:5px solid #FFC500; border-radius:3px; }
        .cat-footer strong { color:#FFC500; }
        </style>
        """,
        unsafe_allow_html=True,
    )


def render_header():
    logo = find_logo_file()
    if logo:
        b64 = img_to_base64(logo)
        logo_html = f'<img class="cat-logo-img" src="data:image/png;base64,{b64}" />' if b64 else '<div class="cat-logo-box">CAT</div>'
    else:
        logo_html = '<div class="cat-logo-box">CAT</div>'
    st.markdown(
        f"""
        <div class="cat-header">
          <div>
            <div class="cat-header-title">REBUILD ANALYTICS PLATFORM</div>
            <div class="cat-header-subtitle">USD-normalized rebuild analytics | Dealer-year labor rates | BLS CPI-U inflation | Outlier and exception intelligence</div>
            <div class="cat-version-pill">{APP_VERSION} &nbsp; | &nbsp; Updated {APP_LAST_UPDATED}</div>
          </div>
          <div>{logo_html}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

inject_cat_theme()
render_header()

# =====================================================
# CONFIG
# =====================================================
CERTIFIED_REBUILD_TYPES = {
    "CER": "COMMERCIAL ENGINE REBUILD",
    "CGR-E": "CERTIFIED GENSET CAPTIVE ENGINE",
    "CHR": "CERTIFIED HYDRAULICS REBUILD",
    "CMCR A-C": "CERTIFIED MACHINE COMPONENT REBUILD AXLE (CENTER)",
    "CMCR A-F": "CERTIFIED MACHINE COMPONENT REBUILD AXLE (FRONT)",
    "CMCR A-R": "CERTIFIED MACHINE COMPONENT REBUILD AXLE (REAR)",
    "CMCR D-C": "CERTIFIED MACHINE COMPONENT REBUILD DIFFERENTIAL (CENTER)",
    "CMCR D-F": "CERTIFIED MACHINE COMPONENT REBUILD DIFFERENTIAL (FRONT)",
    "CMCR D-R": "CERTIFIED MACHINE COMPONENT REBUILD DIFFERENTIAL (REAR)",
    "CMCR-E": "CERTIFIED MACHINE COMPONENT REBUILD - ENGINE",
    "CMCR-F": "CERTIFIED MACHINE COMPONENT REBUILD FINAL DRIVE",
    "CMCR-HFP": "CERTIFIED MACHINE COMPONENT REBUILD HYDRAULIC FAN PUMP AND MOTOR",
    "CMCR-HIP": "CERTIFIED MACHINE COMPONENT REBUILD HYDRAULIC IMPLEMENT PUMP",
    "CMCR-HP": "CERTIFIED MACHINE COMPONENT REBUILD - HYDRAULIC PUMP",
    "CMCR-HSP": "CERTIFIED MACHINE COMPONENT REBUILD HYDRAULIC STEERING PUMP",
    "CMCR-T": "CERTIFIED MACHINE COMPONENT REBUILD - TRANSMISSION AND TORQUE CONVERTOR",
    "CMR": "CERTIFIED MACHINE REBUILD",
    "CMR-U": "CERTIFIED MACHINE REBUILD UPGRADE",
    "CPT+H": "CERTIFIED POWERTRAIN + HYDRAULICS",
    "CPT-O": "CERTIFIED POWERTRAIN",
}
VALID_REGIONS = ["AFRICA", "ANZP/INDONESIA", "CHINA", "EASTERN US", "EUROPE", "INDIA", "JAPAN/ASIA", "LATIN AMERICA", "MIDDLE EAST & EURASIA", "WESTERN US & CANADA"]
REBUILD_TYPE_ALIASES = {"CERTIFIED MACHINE REBUILD": "CMR", "CMR - CERTIFIED MACHINE REBUILD": "CMR", "CERTIFIED MACHINE REBUILD UPGRADE": "CMR-U", "CERTIFIED POWERTRAIN + HYDRAULICS": "CPT+H", "CPT PLUS H": "CPT+H", "CERTIFIED POWERTRAIN": "CPT-O", "CPT 777": "CPT-O"}
METHOD_LOCK_TEXT = """
**Active Methodology Lock**

- **Cost basis:** Adjusted Total Cost USD = PARTS DN USD + Labor Cost USD. PLUS PARTS DN is ignored entirely.
- **Labor:** Labor Cost USD = REBUILD WORK HRS × dealer service-year labor rate converted to USD when applicable.
- **Currency:** Source currency is converted to USD before inflation and analysis.
- **Inflation:** BLS CPI-U All Items, U.S. city average, not seasonally adjusted. Inflation is applied by component.
- **Outliers:** Cost-only log(cost) + IQR by Machine + CCR TYPE.
- **Sample-size rules:** <5 = no removal; 5–7 = 2.0×IQR; 8+ = 1.5×IQR.
- **SMU:** SMU 0 or 1 is a data-quality flag only; SMU is not used as a statistical outlier basis.
- **Cross-type review:** CPT+H rows are flagged when above machine-level valid CMR median × 1.10, only when at least 3 valid CMR rows exist.
- **Adjusted CPT+H:** Adjusted averages exclude cross-type flagged CPT+H rows only; CMR and CPT-O remain unchanged.
- **V16.1 controls:** Strict Mode and configurable thresholds are optional governance controls; defaults preserve the locked methodology unless manually changed.
"""

if "run_clicked" not in st.session_state:
    st.session_state.run_clicked = False
if "analysis" not in st.session_state:
    st.session_state.analysis = None

# =====================================================
# CHART HELPERS
# =====================================================
CAT_DOMAIN = list(CERTIFIED_REBUILD_TYPES.keys()) + ["CPT+H Adjusted"]
CAT_RANGE = ["#000000", "#FFC500", "#7A7A7A", "#4D4D4D", "#FFCD00", "#2B2B2B", "#A6A6A6", "#6B5B00", "#C49700", "#595959", "#B38F00", "#333333", "#D9A300", "#808080", "#F2B705", "#1F1F1F", "#C0C0C0", "#8A6F00", "#E0B000", "#666666", "#FFC500"]
CAT_COLOR_SCALE = alt.Scale(domain=CAT_DOMAIN, range=CAT_RANGE)

def cat_scatter_chart(data, x, y, color="CCR TYPE", tooltip=None, height=420):
    if data.empty:
        st.info("No chart data available.")
        return
    tooltip = tooltip or [x, y, color]
    chart = (alt.Chart(data).mark_circle(size=78, opacity=0.82, stroke="#000000", strokeWidth=0.35)
        .encode(
            x=alt.X(f"{x}:Q", title=x, axis=alt.Axis(labelColor="#1F1F1F", titleColor="#1F1F1F", gridColor="#E6E6E6")),
            y=alt.Y(f"{y}:Q", title=y, axis=alt.Axis(format="$,.0f", labelColor="#1F1F1F", titleColor="#1F1F1F", gridColor="#E6E6E6")),
            color=alt.Color(f"{color}:N", scale=CAT_COLOR_SCALE, legend=alt.Legend(title=color, orient="right")),
            tooltip=tooltip)
        .properties(height=height).configure_view(stroke="#D9D9D9").configure_axis(labelFontSize=12, titleFontSize=13).configure_legend(labelFontSize=12, titleFontSize=13))
    st.altair_chart(chart, use_container_width=True)

def cat_bar_chart(data, x, y, color=None, tooltip=None, height=360):
    if data.empty:
        st.info("No chart data available.")
        return
    tooltip = tooltip or [x, y]
    color_encoding = alt.value("#FFC500") if color is None else alt.Color(f"{color}:N", scale=CAT_COLOR_SCALE, legend=alt.Legend(title=color))
    chart = (alt.Chart(data).mark_bar(cornerRadiusTopLeft=2, cornerRadiusTopRight=2, stroke="#000000", strokeWidth=0.45)
        .encode(
            x=alt.X(f"{x}:N", title=x, sort="-y", axis=alt.Axis(labelAngle=-35, labelColor="#1F1F1F", titleColor="#1F1F1F")),
            y=alt.Y(f"{y}:Q", title=y, axis=alt.Axis(format="$,.0f", labelColor="#1F1F1F", titleColor="#1F1F1F", gridColor="#E6E6E6")),
            color=color_encoding,
            tooltip=tooltip)
        .properties(height=height).configure_view(stroke="#D9D9D9").configure_axis(labelFontSize=12, titleFontSize=13).configure_legend(labelFontSize=12, titleFontSize=13))
    st.altair_chart(chart, use_container_width=True)

# =====================================================
# GENERAL HELPERS
# =====================================================
def money(x):
    try:
        if pd.isna(x):
            return ""
        return f"${x:,.0f}"
    except Exception:
        return x

def sample_confidence(n):
    try:
        n = int(n)
    except Exception:
        return "Unknown"
    if n < 5:
        return "Limited (<5)"
    if n < 8:
        return "Moderate (5–7)"
    return "Strong (8+)"

def is_numeric_series(series):
    cleaned = series.dropna()
    if cleaned.empty:
        return False
    converted = pd.to_numeric(cleaned, errors="coerce")
    return converted.notna().all()

def display_table(df, currency_cols=None, percent_cols=None, number_cols=None, year_cols=None, smu_cols=None, decimal_cols=None):
    currency_cols = currency_cols or []
    percent_cols = percent_cols or []
    number_cols = number_cols or []
    year_cols = year_cols or []
    smu_cols = smu_cols or []
    decimal_cols = decimal_cols or []
    for col in df.columns:
        upper = str(col).upper()
        if "YEAR" in upper and col not in year_cols:
            year_cols.append(col)
        if "SMU" in upper and col not in smu_cols:
            smu_cols.append(col)
    fmt = {}
    def add_format(col, pattern):
        if col in df.columns and is_numeric_series(df[col]):
            fmt[col] = pattern
    for col in currency_cols:
        add_format(col, "${:,.0f}")
    for col in percent_cols:
        add_format(col, "{:.1f}%")
    for col in number_cols + year_cols + smu_cols:
        add_format(col, "{:,.0f}")
    for col in decimal_cols:
        add_format(col, "{:,.4f}")
    if fmt:
        st.dataframe(df.style.format(fmt), use_container_width=True)
    else:
        st.dataframe(df, use_container_width=True)
    return None


def sanitize_excel_value(value):
    """Prevent Excel formula injection in exported workbooks."""
    if isinstance(value, str):
        stripped = value.lstrip()
        if stripped.startswith(("=", "+", "-", "@")):
            return "'" + value
    return value


def sanitize_excel_df(df):
    """Sanitize object/string columns before Excel export."""
    if df is None or not isinstance(df, pd.DataFrame):
        return df
    safe_df = df.copy()
    for col in safe_df.columns:
        if safe_df[col].dtype == "object" or pd.api.types.is_string_dtype(safe_df[col]):
            safe_df[col] = safe_df[col].map(sanitize_excel_value)
    return safe_df


if not hasattr(pd.DataFrame, "_cat_original_to_excel"):
    pd.DataFrame._cat_original_to_excel = pd.DataFrame.to_excel

    def _secure_to_excel(self, *args, **kwargs):
        return pd.DataFrame._cat_original_to_excel(sanitize_excel_df(self), *args, **kwargs)

    pd.DataFrame.to_excel = _secure_to_excel


def file_size_mb(uploaded_file):
    """Return uploaded file size in MB without permanently moving the file pointer."""
    try:
        current_pos = uploaded_file.tell()
        uploaded_file.seek(0, 2)
        size_bytes = uploaded_file.tell()
        uploaded_file.seek(current_pos)
        return size_bytes / (1024 * 1024)
    except Exception:
        return None


def validate_uploaded_file_security(uploaded_file, max_upload_mb=MAX_UPLOAD_MB):
    """Enforce .xlsx-only and max file-size checks."""
    if uploaded_file is None:
        return True
    if not uploaded_file.name.lower().endswith(".xlsx"):
        st.error("Only .xlsx files are allowed for security and compatibility reasons.")
        st.stop()
    uploaded_size_mb = file_size_mb(uploaded_file)
    if uploaded_size_mb is not None and uploaded_size_mb > max_upload_mb:
        st.error(f"Uploaded file is {uploaded_size_mb:.1f} MB. Maximum allowed size is {max_upload_mb} MB.")
        st.stop()
    return True


def reset_app_state():
    """Clear analysis state from Streamlit session."""
    st.session_state.run_clicked = False
    st.session_state.analysis = None


def render_export_acknowledgement(key="export_ack"):
    """Return True only when user acknowledges export authorization."""
    return st.checkbox(
        "I am authorized to download this export.",
        key=key,
    )


def safe_sheet_name(name):
    invalid = ["\\", "/", "*", "[", "]", ":", "?"]
    name = str(name)
    for ch in invalid:
        name = name.replace(ch, "-")
    return name[:31]

def standardize_text(value):
    if pd.isna(value):
        return ""
    return re.sub(r"\s+", " ", str(value).strip().upper())

def normalize_ccr_type(value, fallback_text=""):
    raw = standardize_text(value)
    fallback = standardize_text(fallback_text)
    combined = f"{raw} {fallback}".strip()
    if raw in CERTIFIED_REBUILD_TYPES:
        return raw
    if raw in REBUILD_TYPE_ALIASES:
        return REBUILD_TYPE_ALIASES[raw]
    for code in sorted(CERTIFIED_REBUILD_TYPES.keys(), key=len, reverse=True):
        code_pattern = re.escape(code).replace("\\ ", r"[\s_-]*")
        if re.search(rf"(^|[^A-Z0-9]){code_pattern}([^A-Z0-9]|$)", combined):
            return code
    for alias, code in REBUILD_TYPE_ALIASES.items():
        if alias in combined:
            return code
    return raw if raw else "UNKNOWN"

def normalize_region(value):
    raw = standardize_text(value)
    if not raw:
        return "UNKNOWN"
    aliases = {"WESTERN US AND CANADA": "WESTERN US & CANADA", "WESTERN US & CANADA": "WESTERN US & CANADA", "EASTERN US": "EASTERN US", "LATAM": "LATIN AMERICA", "MIDDLE EAST AND EURASIA": "MIDDLE EAST & EURASIA", "MIDDLE EAST & EURASIA": "MIDDLE EAST & EURASIA", "JAPAN ASIA": "JAPAN/ASIA", "JAPAN/ASIA": "JAPAN/ASIA", "ANZP INDONESIA": "ANZP/INDONESIA", "ANZP/INDONESIA": "ANZP/INDONESIA"}
    if raw in aliases:
        return aliases[raw]
    if raw in VALID_REGIONS:
        return raw
    for region in VALID_REGIONS:
        if region in raw:
            return region
    return raw

def add_vs_cmr(table, type_col="CCR TYPE", avg_col="Avg_Cost"):
    table = table.copy()
    if table.empty or type_col not in table.columns or avg_col not in table.columns:
        return table
    cmr_rows = table[table[type_col] == "CMR"]
    if not cmr_rows.empty:
        cmr_avg = cmr_rows[avg_col].iloc[0]
        if cmr_avg and not pd.isna(cmr_avg):
            table["Vs CMR %"] = ((table[avg_col] - cmr_avg) / cmr_avg) * 100
        else:
            table["Vs CMR %"] = np.nan
    else:
        table["Vs CMR %"] = np.nan
    return table

def add_vs_section_avg(table, avg_col="Avg_Cost"):
    table = table.copy()
    if table.empty or avg_col not in table.columns:
        return table
    section_avg = table[avg_col].mean()
    if section_avg and not pd.isna(section_avg):
        table["Vs Section Avg %"] = ((table[avg_col] - section_avg) / section_avg) * 100
    else:
        table["Vs Section Avg %"] = np.nan
    return table


# =====================================================
# V16 DECISION SUPPORT HELPERS
# =====================================================
def has_both_cmr_cpth(data, type_col="CCR TYPE"):
    if data is None or data.empty or type_col not in data.columns:
        return False
    types = set(data[type_col].dropna().astype(str).str.upper())
    return {"CMR", "CPT+H"}.issubset(types)


def build_dealer_rate_coverage_summary(processed_df):
    if processed_df is None or processed_df.empty or "Rate Source" not in processed_df.columns:
        return pd.DataFrame({"Metric": ["Dealer Rate Coverage Score"], "Value": ["No processed data available"]})
    total = max(len(processed_df), 1)
    direct = int((processed_df["Rate Source"] == "Dealer-Year Rate").sum())
    fallback = total - direct
    missing_dealer = int(processed_df["Dealer Code"].isna().sum()) if "Dealer Code" in processed_df.columns else 0
    exc = int((processed_df["Dealer Rate Exception Flag"] != "").sum()) if "Dealer Rate Exception Flag" in processed_df.columns else 0
    pct = round(direct / total * 100, 1)
    score = "Strong" if pct >= 95 else ("Moderate" if pct >= 80 else "Needs Review")
    return pd.DataFrame({"Metric": ["Dealer Rate Coverage Score", "Dealer-Year Match Rate %", "Dealer-Year Direct Match Rows", "Fallback Rate Rows", "Missing Dealer Code Rows", "Dealer Rate Exception Rows"], "Value": [score, pct, direct, fallback, missing_dealer, exc]})


def build_data_quality_score_summary(processed_df, outlier_count=0, insufficient_count=0):
    if processed_df is None or processed_df.empty:
        return pd.DataFrame({"Metric": ["Data Quality Score"], "Value": ["No processed data available"]})
    total = max(len(processed_df), 1)
    missing_dealer = int(processed_df["Dealer Code"].isna().sum()) if "Dealer Code" in processed_df.columns else 0
    rate_exc = int((processed_df["Dealer Rate Exception Flag"] != "").sum()) if "Dealer Rate Exception Flag" in processed_df.columns else 0
    fx_fallback = int((processed_df["FX Source"] == "Embedded fallback FX table").sum()) if "FX Source" in processed_df.columns else 0
    smu_flags = int(processed_df["Data Quality Exception Flag"].eq("SMU 0 or 1").sum()) if "Data Quality Exception Flag" in processed_df.columns else 0
    other_regions = int((processed_df["Region Display"] == "OTHER").sum()) if "Region Display" in processed_df.columns else 0
    score = 100
    score -= min(25, missing_dealer / total * 100)
    score -= min(25, rate_exc / total * 75)
    score -= min(15, fx_fallback / total * 50)
    score -= min(15, smu_flags / total * 50)
    score -= min(10, other_regions / total * 50)
    score -= min(10, outlier_count / total * 25)
    score -= min(10, insufficient_count / total * 10)
    score = round(max(0, min(100, score)), 1)
    label = "Strong" if score >= 90 else ("Moderate" if score >= 75 else "Needs Review")
    return pd.DataFrame({"Metric": ["Data Quality Score", "Data Quality Label", "Missing Dealer Code Rows", "Dealer Rate Exception Rows", "Fallback FX Rows", "SMU 0 or 1 Rows", "Unknown/Other Region Rows", "Cost Outlier Rows", "Insufficient Sample Rows"], "Value": [score, label, missing_dealer, rate_exc, fx_fallback, smu_flags, other_regions, int(outlier_count), int(insufficient_count)]})


def build_run_readiness_summary(processed_df, rate_coverage_summary, data_quality_score_summary):
    rows = []
    row_count = 0 if processed_df is None else len(processed_df)
    rows.append({"Readiness Check": "Rows After Filters", "Status": "Ready" if row_count > 0 else "Cannot Run", "Details": f"{row_count:,} processed row(s) available."})
    try:
        pct = float(rate_coverage_summary.loc[rate_coverage_summary["Metric"] == "Dealer-Year Match Rate %", "Value"].iloc[0])
        rows.append({"Readiness Check": "Dealer Rate Coverage", "Status": "Ready" if pct >= float(globals().get("dealer_rate_coverage_warning_threshold", 95.0)) else "Needs Review", "Details": f"Dealer-year match rate is {pct}%."})
    except Exception:
        rows.append({"Readiness Check": "Dealer Rate Coverage", "Status": "Needs Review", "Details": "Coverage could not be calculated."})
    try:
        dq = data_quality_score_summary.loc[data_quality_score_summary["Metric"] == "Data Quality Label", "Value"].iloc[0]
        rows.append({"Readiness Check": "Data Quality", "Status": "Ready" if dq == "Strong" else "Needs Review", "Details": f"Data quality label is {dq}."})
    except Exception:
        rows.append({"Readiness Check": "Data Quality", "Status": "Needs Review", "Details": "Score could not be calculated."})
    overall = "Cannot Run" if any(r["Status"] == "Cannot Run" for r in rows) else ("Needs Review" if any(r["Status"] == "Needs Review" for r in rows) else "Ready")
    rows.insert(0, {"Readiness Check": "Overall Run Readiness", "Status": overall, "Details": "Review detailed checks before relying on outputs."})
    return pd.DataFrame(rows)


def build_machine_benchmark_ranking(valid_df, processed_df, cost_col):
    if valid_df is None or valid_df.empty:
        return pd.DataFrame()
    base = valid_df.groupby("SALES MODEL").agg(Avg_Cost=(cost_col, "mean"), Avg_SMU=("SMU AT REBUILD", "mean"), Valid_Rows=(cost_col, "count"), Cross_Type_Flags=("Cross-Type Exception Flag", lambda x: (x != "").sum())).reset_index()
    full = processed_df.groupby("SALES MODEL").agg(Total_Rows=(cost_col, "count"), Outlier_Rows=("Outlier", "sum"), Dealer_Rate_Exception_Rows=("Dealer Rate Exception Flag", lambda x: (x != "").sum())).reset_index()
    out = base.merge(full, on="SALES MODEL", how="left")
    out["Outlier Rate %"] = (out["Outlier_Rows"] / out["Total_Rows"] * 100).fillna(0)
    out["Dealer Rate Exception Rate %"] = (out["Dealer_Rate_Exception_Rows"] / out["Total_Rows"] * 100).fillna(0)
    overall = out["Avg_Cost"].mean()
    out["Vs Overall Avg %"] = ((out["Avg_Cost"] - overall) / overall * 100) if overall else np.nan
    out["Priority Score"] = (out["Vs Overall Avg %"].clip(lower=0).fillna(0) * .5 + out["Outlier Rate %"] * .25 + out["Dealer Rate Exception Rate %"] * .15 + out["Cross_Type_Flags"] * .10).round(1)
    out["Priority Label"] = np.where(out["Priority Score"] >= 25, "High", np.where(out["Priority Score"] >= 10, "Monitor", "Lower"))
    return out.sort_values(["Priority Score", "Avg_Cost"], ascending=False)


def build_executive_narrative(valid_df, summary_df, cost_col, outlier_count, cross_count, rate_cov, dq_score, show_adjusted):
    lines = []
    if valid_df is None or valid_df.empty:
        return ["No valid rows are available for executive narrative generation."]
    lines.append(f"The analysis includes {len(valid_df):,} valid rows after filters and cost-outlier removal.")
    lines.append(f"The average USD analysis cost is {money(valid_df[cost_col].mean())}.")
    if summary_df is not None and not summary_df.empty and "Avg_Cost" in summary_df.columns:
        top = summary_df.sort_values("Avg_Cost", ascending=False).iloc[0]
        lines.append(f"The highest average rebuild type cost is {top['CCR TYPE']} at {money(top['Avg_Cost'])}.")
    try:
        cov = rate_cov.loc[rate_cov["Metric"] == "Dealer-Year Match Rate %", "Value"].iloc[0]
        lines.append(f"Dealer labor-rate coverage is {cov}% direct dealer-year match.")
    except Exception:
        pass
    try:
        score = dq_score.loc[dq_score["Metric"] == "Data Quality Score", "Value"].iloc[0]
        label = dq_score.loc[dq_score["Metric"] == "Data Quality Label", "Value"].iloc[0]
        lines.append(f"The data quality score is {score}, rated {label}.")
    except Exception:
        pass
    lines.append(f"Cost outliers excluded from core averages: {outlier_count:,}.")
    lines.append(f"CPT+H cross-type review flags: {cross_count:,}.")
    lines.append("Adjusted CPT+H comparison is shown because both CMR and CPT+H records are present." if show_adjusted else "Adjusted CPT+H comparison is hidden because both CMR and CPT+H records are required.")
    return lines


def build_cover_sheet(metadata, readiness, dq_score, rate_cov, export_mode="Full Analysis Workbook", export_reason="Not provided", scenario_name_value=None, user_role_view_value="Analyst", strict_mode_value=False):
    """Create a cover sheet without relying on UI globals, so template/export utilities can reuse it safely."""
    def look(table, key, default=""):
        try:
            return table.loc[table.iloc[:, 0] == key, table.columns[1]].iloc[0]
        except Exception:
            return default
    safe_scenario = scenario_name_value if scenario_name_value else "Not provided"
    return pd.DataFrame({
        "Field": ["App Version", "Methodology Version", "Outlier Rule Version", "Dealer Rate Version", "Security Control Version", "Export Format Version", "Export Timestamp", "Export Mode", "Export Reason", "Scenario Name", "User Role View", "Strict Mode", "Overall Run Readiness", "Data Quality Score", "Data Quality Label", "Dealer Rate Coverage Score", "Dealer-Year Match Rate %", "Highlight Legend"],
        "Value": [APP_VERSION, METHODOLOGY_VERSION, OUTLIER_RULE_VERSION, DEALER_RATE_VERSION, SECURITY_CONTROL_VERSION, EXPORT_FORMAT_VERSION, datetime.now().strftime("%Y-%m-%d %H:%M:%S"), export_mode, export_reason, safe_scenario, user_role_view_value, "Yes" if strict_mode_value else "No", look(readiness, "Overall Run Readiness", "Unknown"), look(dq_score, "Data Quality Score", "Unknown"), look(dq_score, "Data Quality Label", "Unknown"), look(rate_cov, "Dealer Rate Coverage Score", "Unknown"), look(rate_cov, "Dealer-Year Match Rate %", "Unknown"), "Red = cost outlier; orange = cross-type exception; yellow = insufficient sample group."]
    })



# =====================================================
# V16.1 GOVERNANCE / FUTURE-PROOFING HELPERS
# =====================================================
def build_known_limitations_table():
    return pd.DataFrame({
        "Known Limitation": [
            "Adjusted CPT+H dependency",
            "Dealer code dependency",
            "Labor-hour dependency",
            "SMU handling",
            "FX fallback",
            "Built-in dealer rate workbook dependency",
            "Outlier scope",
            "Cross-type scope",
            "Security scope",
        ],
        "Details": [
            "Adjusted CPT+H outputs require both CMR and CPT+H records in the relevant scope.",
            "Dealer labor-rate matching depends on detecting a valid four-character dealer code from the DEALER field.",
            "Rows with missing REBUILD WORK HRS cannot be used for adjusted labor-cost analysis.",
            "SMU values of 0 or 1 are data-quality flags only; SMU is not used for statistical outlier removal.",
            "If live FX retrieval fails, embedded fallback FX rates may be used and flagged.",
            "The expanded built-in dealer-rate workbook must be deployed next to app.py; otherwise emergency generic rates are used.",
            "Cost outliers are detected by Machine + CCR TYPE using log(cost) IQR logic only.",
            "CPT+H cross-type review requires at least 3 valid CMR rows for the same machine.",
            " labeling and acknowledgements support governance but do not replace enterprise authentication/authorization.",
        ],
    })


def build_data_dictionary_table():
    return pd.DataFrame({
        "Field": [
            "Adjusted Total Cost USD",
            "Inflation-Adjusted Adjusted Total Cost USD",
            "PARTS DN USD",
            "Labor Cost USD",
            "Rate Source",
            "Dealer Rate Exception Flag",
            "Outlier",
            "Outlier Reason",
            "Insufficient Sample Group",
            "Cross-Type Exception Flag",
            "CMR Benchmark Cost",
            "Data Quality Exception Flag",
            "Dealer-Year Match Rate %",
            "Data Quality Score",
            "Machine Priority Score",
        ],
        "Definition": [
            "PARTS DN converted to USD plus labor cost converted/calculated in USD.",
            "Adjusted Total Cost USD after component-level CPI inflation adjustment to the selected base year.",
            "PARTS DN multiplied by the source-currency FX-to-USD rate.",
            "REBUILD WORK HRS multiplied by the dealer service-year labor rate in USD.",
            "Indicates whether a direct dealer-year rate or fallback rate was used.",
            "Flags rows where the dealer-year labor-rate match was missing or fallback logic was required.",
            "Boolean flag identifying rows excluded from average calculations by cost outlier logic.",
            "Text explanation for statistical cost outlier or insufficient-sample handling.",
            "TRUE when Machine + CCR TYPE has fewer than 5 rows; no statistical cost outlier removal is applied.",
            "Flags CPT+H rows above machine-level valid CMR median × configured threshold when enough CMR rows exist.",
            "Machine-level valid non-outlier CMR median cost used for CPT+H review.",
            "Flags data-quality issues such as SMU 0 or 1.",
            "Percent of processed rows that used a direct dealer-year labor-rate match.",
            "Composite score summarizing major data-quality risks for the run.",
            "Decision-support score ranking machines by cost position, outliers, dealer-rate exceptions, and cross-type flags.",
        ],
    })


def build_parameter_summary_table():
    return pd.DataFrame({
        "Parameter": [
            "Scenario Name",
            "User Role View",
            "Strict Mode Enabled",
            "Maximum Row Warning Threshold",
            "Dealer Rate Coverage Warning Threshold %",
            "Dealer Rate Coverage Strict Threshold %",
            "Data Quality Strict Minimum Score",
            "Cross-Type Threshold Multiplier",
            "Minimum CMR Rows for CPT+H Benchmark",
            "Machine Filter",
            "Start Year",
            "End Year",
            "Base Year",
            "Rebuild Type Filter",
            "Region Filter",
            "Inflation Applied",
            "Dealer Rate Source",
            "Export Mode",
            "Export Reason",
        ],
        "Value": [
            scenario_name if scenario_name else "Not provided",
            user_role_view,
            "Yes" if strict_mode else "No",
            max_rows_warning_threshold,
            dealer_rate_coverage_warning_threshold,
            dealer_rate_coverage_strict_threshold,
            data_quality_strict_min_score,
            cross_type_threshold_multiplier,
            min_cmr_rows_for_benchmark,
            machine_input if machine_input else "All",
            start_year,
            end_year,
            base_year,
            ", ".join(rebuild_filter),
            ", ".join(region_filter),
            "Yes" if apply_inflation else "No",
            "Built-in Expanded Dealer Rates" if use_default else "Uploaded Custom Dealer Rates",
            export_mode,
            export_reason_final,
        ],
    })


def build_role_policy_table():
    return pd.DataFrame({
        "Role View": ["Viewer", "Analyst", "Admin"],
        "Intended Capability": [
            "Dashboard and read-only review. Export controls should be limited in a secured deployment.",
            "Analysis review plus selected-machine and standard export workflows.",
            "Full export and advanced configuration review. Enterprise authentication should enforce this in production.",
        ],
        "Current App Behavior": [
            "Role is advisory only unless enterprise authentication is added.",
            "Role is advisory only unless enterprise authentication is added.",
            "Role is advisory only unless enterprise authentication is added.",
        ],
    })


def build_performance_safeguard_summary(source_rows, processed_rows):
    warnings = []
    if source_rows > max_rows_warning_threshold:
        warnings.append({"Safeguard": "Source Row Count", "Status": "Review", "Details": f"Source workbook contains {source_rows:,} rows, above threshold {max_rows_warning_threshold:,}. Consider Summary Only export if performance is slow."})
    else:
        warnings.append({"Safeguard": "Source Row Count", "Status": "OK", "Details": f"Source workbook contains {source_rows:,} rows."})
    if processed_rows > max_rows_warning_threshold:
        warnings.append({"Safeguard": "Processed Row Count", "Status": "Review", "Details": f"Processed data contains {processed_rows:,} rows. Large raw-data export may take longer."})
    else:
        warnings.append({"Safeguard": "Processed Row Count", "Status": "OK", "Details": f"Processed data contains {processed_rows:,} rows."})
    if export_mode == "Full Analysis Workbook" and processed_rows > max_rows_warning_threshold:
        warnings.append({"Safeguard": "Export Mode", "Status": "Review", "Details": "Full Analysis Workbook selected for a large run. Summary Only or Exceptions Only may be more efficient."})
    else:
        warnings.append({"Safeguard": "Export Mode", "Status": "OK", "Details": f"Selected export mode: {export_mode}."})
    return pd.DataFrame(warnings)


def apply_confidential_watermark(workbook, scenario_name_value=None):
    """No-op retained for compatibility. V16.2.3 removes confidentiality label rows from all exports."""
    return None


def evaluate_strict_mode(run_readiness_summary, dealer_rate_coverage_summary, data_quality_score_summary, processed_rows):
    """Raise a clear error when Strict Mode is enabled and configured thresholds are not met."""
    if not strict_mode:
        return
    issues = []
    try:
        coverage = float(dealer_rate_coverage_summary.loc[dealer_rate_coverage_summary["Metric"] == "Dealer-Year Match Rate %", "Value"].iloc[0])
        if coverage < float(dealer_rate_coverage_strict_threshold):
            issues.append(f"Dealer-year match rate is {coverage:.1f}%, below strict threshold {dealer_rate_coverage_strict_threshold:.1f}%.")
    except Exception:
        issues.append("Dealer-year match rate could not be evaluated.")
    try:
        dq_score = float(data_quality_score_summary.loc[data_quality_score_summary["Metric"] == "Data Quality Score", "Value"].iloc[0])
        if dq_score < float(data_quality_strict_min_score):
            issues.append(f"Data Quality Score is {dq_score:.1f}, below strict threshold {data_quality_strict_min_score:.1f}.")
    except Exception:
        issues.append("Data Quality Score could not be evaluated.")
    if processed_rows > max_rows_warning_threshold * 2:
        issues.append(f"Processed rows ({processed_rows:,}) exceed 2× the row warning threshold ({max_rows_warning_threshold:,}).")
    if issues:
        raise ValueError("Strict Mode stopped the analysis: " + " ".join(issues))


# =====================================================
# EXPORT FORMATTING HELPERS
# =====================================================
def _truthy_excel_value(value):
    """Treat Excel boolean/string/numeric values as true when appropriate."""
    if value is True:
        return True
    if value is False or value is None:
        return False
    text = str(value).strip().upper()
    return text in {"TRUE", "YES", "Y", "1"}


def apply_excel_brand_formatting(workbook):
    """
    Caterpillar-inspired workbook formatting plus row highlighting.

    Highlight rules:
    - Red fill: Outlier == TRUE
    - Orange fill: Cross-Type Exception Flag is populated
    - Yellow fill: Insufficient Sample Group == TRUE
    """
    header_fill = PatternFill(start_color="1F1F1F", end_color="1F1F1F", fill_type="solid")
    header_font = Font(color="FFC500", bold=True)
    thin = Side(style="thin", color="D9E2F3")
    outlier_fill = PatternFill(start_color="F4CCCC", end_color="F4CCCC", fill_type="solid")
    cross_type_fill = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")
    limited_sample_fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
    currency_keywords = ["COST", "RATE", "PARTS DN", "LABOR", "BENCHMARK", "AVG"]
    percent_keywords = ["%"]
    decimal_keywords = ["CPI", "FACTOR", "FX", "SCORE"]

    for ws in workbook.worksheets:
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions
        for cell in ws[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center")

        header_map = {str(cell.value).upper().strip() if cell.value is not None else "": cell.column for cell in ws[1]}
        outlier_col = header_map.get("OUTLIER")
        cross_type_col = header_map.get("CROSS-TYPE EXCEPTION FLAG")
        insufficient_col = header_map.get("INSUFFICIENT SAMPLE GROUP")

        for col_idx in range(1, ws.max_column + 1):
            header = str(ws.cell(row=1, column=col_idx).value).upper() if ws.cell(row=1, column=col_idx).value is not None else ""
            col_letter = get_column_letter(col_idx)
            ws.column_dimensions[col_letter].width = min(max(len(header) + 4, 14), 42)
            for cell in ws[col_letter][1:]:
                cell.border = Border(bottom=thin)
                if "YEAR" in header or "SMU" in header or header in ["COUNT", "TOTAL_ROWS", "OUTLIER_ROWS", "VALID_ROWS", "VALUE"]:
                    cell.number_format = "#,##0"
                elif any(keyword in header for keyword in percent_keywords):
                    cell.number_format = "0.0%"
                elif any(keyword in header for keyword in decimal_keywords):
                    cell.number_format = "0.0000"
                elif any(keyword in header for keyword in currency_keywords):
                    cell.number_format = "$#,##0"

        for row_idx in range(2, ws.max_row + 1):
            row_fill = None
            if outlier_col and _truthy_excel_value(ws.cell(row=row_idx, column=outlier_col).value):
                row_fill = outlier_fill
            elif cross_type_col:
                cross_value = ws.cell(row=row_idx, column=cross_type_col).value
                if cross_value is not None and str(cross_value).strip() != "":
                    row_fill = cross_type_fill
            elif insufficient_col and _truthy_excel_value(ws.cell(row=row_idx, column=insufficient_col).value):
                row_fill = limited_sample_fill
            if row_fill is not None:
                for cell in ws[row_idx]:
                    cell.fill = row_fill

        for cell in ws[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center")


# =====================================================
# FX / CPI / RATE HELPERS
# =====================================================
COMMON_CURRENCY_FALLBACK_TO_USD = {"USD": 1.00, "CAD": 0.74, "EUR": 1.09, "GBP": 1.27, "AUD": 0.66, "NZD": 0.61, "MXN": 0.058, "BRL": 0.19, "CLP": 0.0011, "COP": 0.00025, "PEN": 0.27, "JPY": 0.0065, "CNY": 0.14, "INR": 0.012, "ZAR": 0.055}

def normalize_currency_code(value, default_code="USD"):
    if pd.isna(value):
        return default_code.upper()
    code = str(value).strip().upper()
    if code in ["$", "US$", "USDOLLAR", "US DOLLAR", "U.S. DOLLAR", ""]:
        return "USD"
    if len(code) >= 3:
        return code[:3]
    return default_code.upper()

def detect_currency_column(df):
    candidates = ["CURRENCY", "Currency", "currency", "CURR", "Curr", "curr", "CURRENCY CODE", "Currency Code", "currency_code", "CCY"]
    for col in candidates:
        if col in df.columns:
            return col
    for col in df.columns:
        col_upper = str(col).upper()
        if "CURRENCY" in col_upper or col_upper in ["CURR", "CCY"]:
            return col
    return None

def frankfurter_year_average_to_usd(currency, year):
    currency = normalize_currency_code(currency)
    year = int(year)
    if currency == "USD":
        return 1.0, "USD"
    url = f"https://api.frankfurter.dev/v1/{year}-01-01..{year}-12-31?base={currency}&symbols=USD"
    with urllib.request.urlopen(url, timeout=20) as response:
        data = json.loads(response.read().decode("utf-8"))
    rates = []
    for values in data.get("rates", {}).values():
        if "USD" in values:
            rates.append(float(values["USD"]))
    if not rates:
        raise RuntimeError(f"No FX rates returned for {currency} {year}")
    return float(np.mean(rates)), "Frankfurter annual average"

def build_fx_lookup(currencies, years):
    rows = []
    cache = {}
    for currency in sorted(set([normalize_currency_code(c) for c in currencies if pd.notna(c)])):
        for year in sorted(set([int(y) for y in years if pd.notna(y)])):
            key = (currency, year)
            if key in cache:
                fx, source = cache[key]
            else:
                try:
                    fx, source = frankfurter_year_average_to_usd(currency, year)
                except Exception:
                    fx = COMMON_CURRENCY_FALLBACK_TO_USD.get(currency, 1.0)
                    source = "Embedded fallback FX table"
                cache[key] = (fx, source)
            rows.append({"Currency": currency, "Service Year": year, "FX to USD": fx, "FX Source": source})
    return pd.DataFrame(rows)

def fallback_bls_cpi_table(start_year, end_year):
    fallback = {2010: 218.056, 2011: 224.939, 2012: 229.594, 2013: 232.957, 2014: 236.736, 2015: 237.017, 2016: 240.007, 2017: 245.120, 2018: 251.107, 2019: 255.657, 2020: 258.811, 2021: 270.970, 2022: 292.655, 2023: 304.702, 2024: 313.689, 2025: 322.000, 2026: 330.080, 2027: 339.982, 2028: 350.181, 2029: 360.686, 2030: 371.506}
    return {y: fallback[y] for y in range(int(start_year), int(end_year) + 1) if y in fallback}

def fetch_bls_cpi_annual(start_year, end_year):
    url = "https://api.bls.gov/publicAPI/v2/timeseries/data/"
    payload = {"seriesid": ["CUUR0000SA0"], "startyear": str(int(start_year)), "endyear": str(int(end_year))}
    request = urllib.request.Request(url, data=json.dumps(payload).encode("utf-8"), headers={"Content-Type": "application/json"}, method="POST")
    with urllib.request.urlopen(request, timeout=20) as response:
        result = json.loads(response.read().decode("utf-8"))
    if result.get("status") != "REQUEST_SUCCEEDED":
        raise RuntimeError("BLS API request did not succeed")
    data = result["Results"]["series"][0]["data"]
    by_year = {}
    annual_m13 = {}
    for row in data:
        year = int(row["year"])
        period = row["period"]
        value = float(row["value"])
        if period == "M13":
            annual_m13[year] = value
        elif period.startswith("M") and period != "M13":
            by_year.setdefault(year, []).append(value)
    cpi = {}
    for year in range(int(start_year), int(end_year) + 1):
        if year in annual_m13:
            cpi[year] = annual_m13[year]
        elif year in by_year and len(by_year[year]) > 0:
            cpi[year] = float(np.mean(by_year[year]))
    return cpi

def get_cpi_table(start_year, end_year, base_year):
    min_year = min(int(start_year), int(base_year))
    max_year = max(int(end_year), int(base_year))
    source = "BLS API: CPI-U All Items, U.S. city average, not seasonally adjusted (CUUR0000SA0)"
    try:
        cpi = fetch_bls_cpi_annual(min_year, max_year)
        if int(base_year) not in cpi:
            raise RuntimeError("Base year CPI unavailable from BLS API")
        return cpi, source
    except Exception:
        return fallback_bls_cpi_table(min_year, max_year), "Fallback embedded CPI table because BLS API was unavailable"

def build_default_rate_table(start=2010, end=2030):
    """
    Return built-in expanded dealer-year labor rates from the bundled workbook.
    The workbook must sit in the same folder as app.py when deployed.
    Falls back to the old generic yearly rate table only if the workbook is missing/unreadable.
    """
    built_in_rate_path = Path("expanded_dealer_base_rate_by_year_2016_2026.xlsx")
    if built_in_rate_path.exists():
        try:
            rate_df = parse_multisheet_rate_workbook(built_in_rate_path)
            rate_df = clean_rate_table(rate_df)
            if not rate_df.empty:
                if start is not None and end is not None:
                    rate_df = rate_df[(rate_df["Service Year"] >= int(start)) & (rate_df["Service Year"] <= int(end))].copy()
                rate_df["Rate File Format"] = "Built-in Expanded Dealer Rates 2016-2026"
                rate_df["Notes"] = "Built-in expanded dealer base rate workbook: 2016-2026"
                return rate_df
        except Exception:
            pass

    # Emergency fallback only if the built-in workbook is unavailable.
    years = list(range(start, end + 1))
    return pd.DataFrame({
        "Dealer Code": ["DEFAULT"] * len(years),
        "Service Year": years,
        "Rate": [115 + (year - 2016) * 3 for year in years],
        "Rate Currency": ["USD"] * len(years),
        "Notes": ["Emergency generic fallback because built-in dealer workbook was not found/readable"] * len(years),
        "Rate File Sheet": ["Emergency Fallback"] * len(years),
        "Rate File Format": ["Emergency Generic Default"] * len(years),
    })


def normalize_dealer_code(value):
    if pd.isna(value):
        return ""
    text = str(value).strip().upper()
    match = re.search(r"([A-Z][A-Z0-9]{3})", text)
    return match.group(1) if match else text


def normalize_rate_columns(df):
    """Normalize common dealer-rate column names to Dealer Code, Service Year, Rate, Rate Currency, Notes."""
    temp = df.copy()
    temp.columns = [str(col).strip() for col in temp.columns]
    rename_map = {}
    for col in temp.columns:
        cu = str(col).strip().upper().replace("_", " ")
        if cu in ["DEALER", "DEALER CODE", "DEALER ID", "CODE"] or ("DEALER" in cu and "CODE" in cu):
            rename_map[col] = "Dealer Code"
        elif cu in ["YEAR", "SERVICE YEAR", "RATE YEAR", "BASE RATE YEAR"] or ("YEAR" in cu and "RATE" in cu):
            rename_map[col] = "Service Year"
        elif cu in ["RATE", "LABOR RATE", "BASE RATE", "AVG BASE RATE", "AVERAGE BASE RATE", "DEALER AVERAGE BASE RATE"] or ("RATE" in cu and "SOURCE" not in cu and "CURRENCY" not in cu):
            rename_map[col] = "Rate"
        elif cu in ["CURRENCY", "RATE CURRENCY", "CURRENCY CODE", "CCY"]:
            rename_map[col] = "Rate Currency"
        elif cu in ["NOTE", "NOTES", "COMMENTS", "COMMENT"]:
            rename_map[col] = "Notes"
    temp = temp.rename(columns=rename_map)
    return temp


def parse_flat_rate_table(rate_file, sheet_name=None):
    """Parse a preferred flat dealer-rate table from one sheet."""
    sheets = pd.read_excel(rate_file, sheet_name=None)
    candidate_items = []
    if sheet_name and sheet_name in sheets:
        candidate_items = [(sheet_name, sheets[sheet_name])]
    else:
        candidate_items = list(sheets.items())

    for sname, sheet in candidate_items:
        temp = normalize_rate_columns(sheet)
        if {"Dealer Code", "Service Year", "Rate"}.issubset(set(temp.columns)):
            out = temp.copy()
            out["Rate File Sheet"] = sname
            out["Rate File Format"] = "Flat Table"
            if "Rate Currency" not in out.columns:
                out["Rate Currency"] = "USD"
            if "Notes" not in out.columns:
                out["Notes"] = ""
            return out[["Dealer Code", "Service Year", "Rate", "Rate Currency", "Notes", "Rate File Sheet", "Rate File Format"]]
    return pd.DataFrame(columns=["Dealer Code", "Service Year", "Rate", "Rate Currency", "Notes", "Rate File Sheet", "Rate File Format"])


def parse_multisheet_rate_workbook(rate_file):
    """Parse legacy dealer-rate workbook where each sheet represents a dealer."""
    sheets = pd.read_excel(rate_file, sheet_name=None)
    rows = []
    for sheet_name, sheet in sheets.items():
        if str(sheet_name).strip().lower() == "summary":
            continue
        temp = normalize_rate_columns(sheet)
        dealer_code = normalize_dealer_code(str(sheet_name).split("_")[0].strip())
        if "Dealer Code" not in temp.columns:
            temp["Dealer Code"] = dealer_code
        if "Rate Currency" not in temp.columns:
            temp["Rate Currency"] = "USD"
        if "Notes" not in temp.columns:
            temp["Notes"] = ""
        if "Service Year" in temp.columns and "Rate" in temp.columns:
            keep = temp[["Dealer Code", "Service Year", "Rate", "Rate Currency", "Notes"]].copy()
            keep["Rate File Sheet"] = sheet_name
            keep["Rate File Format"] = "Multi-Sheet Dealer Workbook"
            rows.append(keep)
    if not rows:
        return pd.DataFrame(columns=["Dealer Code", "Service Year", "Rate", "Rate Currency", "Notes", "Rate File Sheet", "Rate File Format"])
    return pd.concat(rows, ignore_index=True)


def clean_rate_table(rate_df):
    """Clean, type, and de-duplicate dealer rate table."""
    if rate_df is None or rate_df.empty:
        return pd.DataFrame(columns=["Dealer Code", "Service Year", "Rate", "Rate Currency", "Notes", "Rate File Sheet", "Rate File Format"])
    out = rate_df.copy()
    out["Dealer Code"] = out["Dealer Code"].apply(normalize_dealer_code)
    out["Service Year"] = pd.to_numeric(out["Service Year"], errors="coerce")
    out["Rate"] = pd.to_numeric(out["Rate"], errors="coerce")
    out["Rate Currency"] = out.get("Rate Currency", "USD")
    out["Rate Currency"] = out["Rate Currency"].apply(lambda x: normalize_currency_code(x, "USD"))
    if "Notes" not in out.columns:
        out["Notes"] = ""
    if "Rate File Sheet" not in out.columns:
        out["Rate File Sheet"] = "Unknown"
    if "Rate File Format" not in out.columns:
        out["Rate File Format"] = "Unknown"
    out = out.dropna(subset=["Service Year", "Rate"])
    out = out[out["Dealer Code"].astype(str).str.strip() != ""]
    out = out[out["Rate"] > 0]
    out["Service Year"] = out["Service Year"].astype(int)
    out = out.sort_values(["Dealer Code", "Service Year", "Rate File Sheet"]).drop_duplicates(["Dealer Code", "Service Year"], keep="last")
    return out[["Dealer Code", "Service Year", "Rate", "Rate Currency", "Notes", "Rate File Sheet", "Rate File Format"]]


def build_rate_table_from_workbook(rate_file, rate_format="Auto-detect"):
    """Build dealer rate table from either the preferred flat table or legacy multi-sheet workbook."""
    if rate_file is None:
        return build_default_rate_table()
    rate_file.seek(0)
    if rate_format == "Flat Table":
        rate_df = parse_flat_rate_table(rate_file)
    elif rate_format == "Multi-Sheet Dealer Workbook":
        rate_file.seek(0)
        rate_df = parse_multisheet_rate_workbook(rate_file)
    else:
        rate_df = parse_flat_rate_table(rate_file)
        if rate_df.empty:
            rate_file.seek(0)
            rate_df = parse_multisheet_rate_workbook(rate_file)
    return clean_rate_table(rate_df)


def validate_dealer_rate_table(rate_df, selected_start_year=None, selected_end_year=None):
    """Create validation summary for uploaded dealer labor rates."""
    if rate_df is None or rate_df.empty:
        return pd.DataFrame({"Check": ["Rate rows parsed"], "Value": [0], "Status": ["Review"]})
    checks = []
    checks.append({"Check": "Rate rows parsed", "Value": len(rate_df), "Status": "OK" if len(rate_df) > 0 else "Review"})
    checks.append({"Check": "Unique dealers", "Value": rate_df["Dealer Code"].nunique(), "Status": "OK"})
    checks.append({"Check": "Minimum service year", "Value": int(rate_df["Service Year"].min()) if not rate_df.empty else "", "Status": "OK"})
    checks.append({"Check": "Maximum service year", "Value": int(rate_df["Service Year"].max()) if not rate_df.empty else "", "Status": "OK"})
    checks.append({"Check": "Missing dealer codes", "Value": int((rate_df["Dealer Code"].astype(str).str.strip() == "").sum()), "Status": "OK" if int((rate_df["Dealer Code"].astype(str).str.strip() == "").sum()) == 0 else "Review"})
    checks.append({"Check": "Non-positive rates", "Value": int((rate_df["Rate"] <= 0).sum()), "Status": "OK" if int((rate_df["Rate"] <= 0).sum()) == 0 else "Review"})
    duplicate_count = int(rate_df.duplicated(["Dealer Code", "Service Year"]).sum())
    checks.append({"Check": "Duplicate dealer-year rows after cleaning", "Value": duplicate_count, "Status": "OK" if duplicate_count == 0 else "Review"})
    if selected_start_year is not None and selected_end_year is not None:
        outside = int(((rate_df["Service Year"] < selected_start_year) | (rate_df["Service Year"] > selected_end_year)).sum())
        checks.append({"Check": "Rows outside selected analysis year range", "Value": outside, "Status": "OK" if outside == 0 else "Info"})
    checks.append({"Check": "Detected rate format(s)", "Value": ", ".join(sorted(rate_df["Rate File Format"].dropna().unique())), "Status": "OK"})
    return pd.DataFrame(checks)


def create_dealer_rate_template():
    """Return a downloadable preferred flat dealer-rate template workbook."""
    template = pd.DataFrame({
        "Dealer Code": ["A000", "A000", "B123"],
        "Service Year": [2024, 2025, 2025],
        "Rate": [150.00, 155.00, 162.50],
        "Rate Currency": ["USD", "USD", "USD"],
        "Notes": ["Example only", "Example only", "Example only"],
    })
    instructions = pd.DataFrame({
        "Field": ["Dealer Code", "Service Year", "Rate", "Rate Currency", "Notes"],
        "Requirement": ["Required", "Required", "Required", "Optional", "Optional"],
        "Description": [
            "Dealer code used in rebuild file, such as A000 or B123.",
            "Calendar/service year for the labor rate.",
            "Dealer labor rate for that service year. Must be positive numeric.",
            "Currency code for the rate. Default expected value is USD.",
            "Optional user notes for auditability.",
        ]
    })
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        template.to_excel(writer, sheet_name="Dealer Rates", index=False)
        instructions.to_excel(writer, sheet_name="Instructions", index=False)
        apply_excel_brand_formatting(writer.book)
        apply_confidential_watermark(writer.book)
    output.seek(0)
    return output.getvalue()


# =====================================================
# PRE-RUN VALIDATION / UI
# =====================================================
def read_rebuild_workbook(uploaded_file):
    uploaded_file.seek(0)
    return pd.read_excel(uploaded_file, sheet_name=None)

def validate_uploaded_workbook(uploaded_file):
    try:
        rebuild = read_rebuild_workbook(uploaded_file)
    except Exception as exc:
        return None, pd.DataFrame({"Issue": [f"Unable to read workbook: {exc}"]})
    required = ["DELIVERED DATE", "ENROLL DATE", "IN SERVICE DATE", "SALES MODEL", "DEALER", "PARTS DN", "SMU AT REBUILD", "REBUILD WORK HRS"]
    issues = []
    rows = []
    all_frames = []
    for sheet_name, sheet in rebuild.items():
        sheet = sheet.copy()
        missing = [col for col in required if col not in sheet.columns]
        detected_type = normalize_ccr_type(sheet_name, sheet_name)
        rows.append({"Sheet": sheet_name, "Rows": len(sheet), "Detected CCR TYPE": detected_type, "Missing Required Columns": ", ".join(missing) if missing else "None"})
        if missing:
            issues.append({"Issue Type": "Missing Columns", "Sheet": sheet_name, "Details": ", ".join(missing)})
        sheet["_Detected CCR TYPE"] = detected_type
        all_frames.append(sheet)
    profile = pd.DataFrame(rows)
    if all_frames:
        combined = pd.concat(all_frames, ignore_index=True)
        if "REGION" in combined.columns:
            observed_regions = sorted(combined["REGION"].dropna().apply(normalize_region).unique())
            for region in [region for region in observed_regions if region not in VALID_REGIONS]:
                issues.append({"Issue Type": "Unconfigured Region", "Sheet": "All", "Details": region})
        observed_types = sorted(combined["_Detected CCR TYPE"].dropna().unique())
        for rebuild_type in [rtype for rtype in observed_types if rtype not in CERTIFIED_REBUILD_TYPES]:
            issues.append({"Issue Type": "Unknown Rebuild Type", "Sheet": "All", "Details": rebuild_type})
    issue_df = pd.DataFrame(issues) if issues else pd.DataFrame(columns=["Issue Type", "Sheet", "Details"])
    return profile, issue_df

# =====================================================
# SINGLE MACHINE EXPORT - V15.4
# =====================================================
def dealer_performance_for_df(machine_valid, machine_all, cost_col):
    if machine_valid.empty:
        return pd.DataFrame()
    dealer_summary = machine_valid.groupby("DEALER").agg(
        Avg_Cost=(cost_col, "mean"),
        Avg_SMU=("SMU AT REBUILD", "mean"),
        Count=(cost_col, "count"),
        Cross_Type_Flags=("Cross-Type Exception Flag", lambda x: (x != "").sum()),
    ).reset_index().sort_values("Avg_Cost", ascending=False)
    dealer_summary = add_vs_section_avg(dealer_summary)
    dealer_outliers = machine_all.groupby("DEALER").agg(
        Total_Rows=(cost_col, "count"),
        Outlier_Rows=("Outlier", "sum"),
        DQ_Rows=("Data Quality Exception Flag", lambda x: (x != "").sum()),
    ).reset_index()
    dealer_summary = dealer_summary.merge(dealer_outliers, on="DEALER", how="left")
    dealer_summary["Outlier Rate %"] = (dealer_summary["Outlier_Rows"] / dealer_summary["Total_Rows"]) * 100
    dealer_summary["DQ Rate %"] = (dealer_summary["DQ_Rows"] / dealer_summary["Total_Rows"]) * 100
    dealer_summary["Cross Flag Rate %"] = (dealer_summary["Cross_Type_Flags"] / dealer_summary["Count"]) * 100
    dealer_summary["Performance Score"] = (
        100
        - dealer_summary["Vs Section Avg %"].clip(lower=0).fillna(0) * 0.50
        - dealer_summary["Outlier Rate %"].fillna(0) * 0.30
        - dealer_summary["Cross Flag Rate %"].fillna(0) * 0.20
        - dealer_summary["DQ Rate %"].fillna(0) * 0.20
    ).clip(lower=0, upper=100)
    dealer_summary["Performance Label"] = np.where(dealer_summary["Performance Score"] >= 85, "Strong", np.where(dealer_summary["Performance Score"] >= 70, "Monitor", "Review Needed"))
    return dealer_summary

def region_performance_for_df(machine_valid, cost_col):
    if machine_valid.empty:
        return pd.DataFrame()
    region_summary = machine_valid.groupby("Region").agg(
        Avg_Cost=(cost_col, "mean"),
        Avg_SMU=("SMU AT REBUILD", "mean"),
        Count=(cost_col, "count"),
        Cross_Type_Flags=("Cross-Type Exception Flag", lambda x: (x != "").sum()),
    ).reset_index().sort_values("Avg_Cost", ascending=False)
    region_summary = add_vs_section_avg(region_summary)
    return region_summary

def build_machine_export(selected_machine, analysis):
    df = analysis["df"]
    valid = analysis["valid"]
    cost_col = analysis["cost_col"]
    metadata = analysis["metadata"].copy()
    fx_lookup = analysis["fx_lookup"].copy()
    cpi_table = analysis["cpi_table"]

    machine_all = df[df["SALES MODEL"] == selected_machine].copy()
    machine_valid = valid[valid["SALES MODEL"] == selected_machine].copy()
    machine_adjusted_valid = machine_valid[~((machine_valid["CCR TYPE"] == "CPT+H") & (machine_valid["Cross-Type Exception Flag"] == "CPT+H Cost Above Typical CMR"))].copy()

    standard_avg = machine_valid[cost_col].mean() if not machine_valid.empty else np.nan
    adjusted_avg = machine_adjusted_valid[cost_col].mean() if not machine_adjusted_valid.empty else np.nan
    adjusted_cpth_avg = machine_adjusted_valid[machine_adjusted_valid["CCR TYPE"] == "CPT+H"][cost_col].mean() if not machine_adjusted_valid.empty else np.nan

    machine_summary = pd.DataFrame({
        "Metric": [
            "Selected Machine", "Total Rows", "Valid Rows", "Standard Avg Cost", "Adjusted Avg Cost",
            "Adjusted CPT+H Avg", "Average SMU", "Cost Outliers", "Cross-Type Flags",
            "SMU Data Quality Flags", "Fallback Labor Rate Rows", "Fallback FX Rows"
        ],
        "Value": [
            selected_machine, len(machine_all), len(machine_valid), standard_avg, adjusted_avg, adjusted_cpth_avg,
            machine_valid["SMU AT REBUILD"].mean() if not machine_valid.empty else np.nan,
            int(machine_all["Outlier"].sum()) if "Outlier" in machine_all.columns else 0,
            int((machine_valid["Cross-Type Exception Flag"] != "").sum()) if "Cross-Type Exception Flag" in machine_valid.columns else 0,
            int((machine_all["Data Quality Exception Flag"] != "").sum()) if "Data Quality Exception Flag" in machine_all.columns else 0,
            int(machine_all["Rate Source"].isin(["Global Yearly Fallback Rate", "Overall Average Fallback Rate"]).sum()) if "Rate Source" in machine_all.columns else 0,
            int((machine_all["FX Source"] == "Embedded fallback FX table").sum()) if "FX Source" in machine_all.columns else 0,
        ]
    })

    
    if "Metric" in machine_summary.columns:
        machine_summary = machine_summary[~machine_summary["Metric"].astype(str).str.contains("Confidential|Removed Label", case=False, na=False)].copy()
    if "Value" in machine_summary.columns:
        machine_summary = machine_summary[~machine_summary["Value"].astype(str).str.contains("Confidential|Yellow", case=False, na=False)].copy()
    rebuild_summary = machine_valid.groupby("CCR TYPE").agg(
        Avg_Cost=(cost_col, "mean"),
        Avg_SMU=("SMU AT REBUILD", "mean"),
        Count=(cost_col, "count"),
    ).reset_index() if not machine_valid.empty else pd.DataFrame(columns=["CCR TYPE", "Avg_Cost", "Avg_SMU", "Count"])
    rebuild_summary = add_vs_cmr(rebuild_summary)
    if not rebuild_summary.empty:
        rebuild_summary["Sample Confidence"] = rebuild_summary["Count"].apply(sample_confidence)

    adjusted_summary = machine_adjusted_valid.groupby("CCR TYPE").agg(
        Adjusted_Avg_Cost=(cost_col, "mean"),
        Count=(cost_col, "count"),
    ).reset_index() if not machine_adjusted_valid.empty else pd.DataFrame(columns=["CCR TYPE", "Adjusted_Avg_Cost", "Count"])
    if not adjusted_summary.empty:
        adjusted_summary["CCR TYPE"] = adjusted_summary["CCR TYPE"].replace({"CPT+H": "CPT+H Adjusted"})
        adjusted_summary["Sample Confidence"] = adjusted_summary["Count"].apply(sample_confidence)

    standard_cpth = machine_valid[machine_valid["CCR TYPE"] == "CPT+H"][cost_col].mean() if not machine_valid.empty else np.nan
    adjusted_cpth = adjusted_cpth_avg
    cpth_diff = adjusted_cpth - standard_cpth if pd.notna(adjusted_cpth) and pd.notna(standard_cpth) else np.nan
    cpth_diff_pct = (cpth_diff / standard_cpth) if pd.notna(cpth_diff) and standard_cpth else np.nan
    adjusted_comparison = pd.DataFrame({
        "Metric": ["Standard CPT+H Avg", "Adjusted CPT+H Avg", "Difference $", "Difference %", "Cross-Type Flags Removed"],
        "Value": [standard_cpth, adjusted_cpth, cpth_diff, cpth_diff_pct, int(((machine_valid["CCR TYPE"] == "CPT+H") & (machine_valid["Cross-Type Exception Flag"] != "")).sum()) if not machine_valid.empty else 0]
    })

    dealer_summary = dealer_performance_for_df(machine_valid, machine_all, cost_col)
    region_summary = region_performance_for_df(machine_valid, cost_col)

    exceptions_summary = pd.DataFrame({
        "Metric": ["Cost Outliers", "CPT+H Cross-Type Flags", "SMU 0 or 1", "Insufficient Sample Rows", "Rows Using Fallback FX", "Rows Using Global Yearly Fallback Rate", "Rows Using Overall Average Fallback Rate"],
        "Value": [
            int(machine_all["Outlier"].sum()) if "Outlier" in machine_all.columns else 0,
            int((machine_valid["Cross-Type Exception Flag"] != "").sum()) if "Cross-Type Exception Flag" in machine_valid.columns else 0,
            int(machine_all["Data Quality Exception Flag"].eq("SMU 0 or 1").sum()) if "Data Quality Exception Flag" in machine_all.columns else 0,
            int(machine_all["Insufficient Sample Group"].sum()) if "Insufficient Sample Group" in machine_all.columns else 0,
            int((machine_all["FX Source"] == "Embedded fallback FX table").sum()) if "FX Source" in machine_all.columns else 0,
            int((machine_all["Rate Source"] == "Global Yearly Fallback Rate").sum()) if "Rate Source" in machine_all.columns else 0,
            int((machine_all["Rate Source"] == "Overall Average Fallback Rate").sum()) if "Rate Source" in machine_all.columns else 0,
        ]
    })

    outlier_perf = machine_all.groupby("CCR TYPE").agg(
        Total_Rows=(cost_col, "count"),
        Outlier_Rows=("Outlier", "sum"),
        Avg_Cost_All=(cost_col, "mean"),
    ).reset_index() if not machine_all.empty else pd.DataFrame(columns=["CCR TYPE", "Total_Rows", "Outlier_Rows", "Avg_Cost_All"])
    valid_counts = machine_valid.groupby("CCR TYPE").size().reset_index(name="Valid_Rows") if not machine_valid.empty else pd.DataFrame(columns=["CCR TYPE", "Valid_Rows"])
    outlier_perf = outlier_perf.merge(valid_counts, on="CCR TYPE", how="left").fillna({"Valid_Rows": 0})
    if not outlier_perf.empty:
        outlier_perf["Outlier Rate %"] = (outlier_perf["Outlier_Rows"] / outlier_perf["Total_Rows"])

    cross_rows = machine_valid[machine_valid["Cross-Type Exception Flag"] != ""].copy() if "Cross-Type Exception Flag" in machine_valid.columns else pd.DataFrame()
    if cross_rows.empty:
        cross_rows = pd.DataFrame({"Message": ["No cross-type exceptions detected for this machine."]})

    run_metadata = pd.concat([
        pd.DataFrame({"Field": ["Selected Machine"], "Value": [selected_machine]}),
        metadata,
    ], ignore_index=True)

    methodology = pd.DataFrame({
        "Methodology Item": [
            "Cost Basis", "Labor Cost", "Currency", "Inflation", "Outliers", "Sample Size", "SMU", "Cross-Type Rule", "Adjusted CPT+H", "Dealer Score"
        ],
        "Description": [
            "Adjusted Total Cost USD = PARTS DN USD + Labor Cost USD. PLUS PARTS DN is ignored.",
            "Labor Cost USD = REBUILD WORK HRS × dealer service-year labor rate converted to USD when applicable.",
            "Source currency is converted to USD before inflation and analysis.",
            "BLS CPI-U All Items, U.S. city average, not seasonally adjusted. Inflation applied by component.",
            "Cost-only log(cost) + IQR by Machine + CCR TYPE.",
            "<5 rows: no removal; 5–7 rows: 2.0×IQR; 8+ rows: 1.5×IQR.",
            "SMU 0 or 1 is a data-quality flag only and is not a statistical outlier basis.",
            "CPT+H flagged when above machine-level valid CMR median × 1.10, if at least 3 valid CMR rows exist.",
            "Adjusted CPT+H excludes cross-type flagged CPT+H rows only; CMR and CPT-O stay unchanged.",
            "Score subtracts weighted impacts from above-average cost, outlier rate, cross-type rate, and data-quality rate.",
        ]
    })

    cpi_export = pd.DataFrame({"Year": list(cpi_table.keys()), "CPI": list(cpi_table.values())}).sort_values("Year")

    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        machine_summary.to_excel(writer, sheet_name="Machine Summary", index=False)
        rebuild_summary.to_excel(writer, sheet_name="Rebuild Type Summary", index=False)
        adjusted_summary.to_excel(writer, sheet_name="Adjusted Cost Summary", index=False)
        adjusted_comparison.to_excel(writer, sheet_name="Adjusted CPT-H Impact", index=False)
        dealer_summary.to_excel(writer, sheet_name="Dealer Performance", index=False)
        region_summary.to_excel(writer, sheet_name="Region Performance", index=False)
        machine_all.to_excel(writer, sheet_name="Raw Data", index=False)
        machine_valid.to_excel(writer, sheet_name="Valid Raw Data", index=False)
        exceptions_summary.to_excel(writer, sheet_name="Exceptions", index=False)
        outlier_perf.to_excel(writer, sheet_name="Outlier Performance", index=False)
        cross_rows.to_excel(writer, sheet_name="Cross Type Flags", index=False)
        machine_rate_exceptions = machine_all[machine_all.get("Dealer Rate Exception Flag", "") != ""].copy() if "Dealer Rate Exception Flag" in machine_all.columns else pd.DataFrame()
        if machine_rate_exceptions.empty:
            machine_rate_exceptions = pd.DataFrame({"Message": ["No dealer rate exceptions detected for this machine."]})
        machine_rate_exceptions.to_excel(writer, sheet_name="Rate Exceptions", index=False)
        run_metadata.to_excel(writer, sheet_name="Run Metadata", index=False)
        methodology.to_excel(writer, sheet_name="Methodology", index=False)
        cpi_export.to_excel(writer, sheet_name="BLS CPI", index=False)
        fx_lookup.to_excel(writer, sheet_name="FX Rates", index=False)
        apply_excel_brand_formatting(writer.book)
        apply_confidential_watermark(writer.book, globals().get("scenario_name", None))
    output.seek(0)
    return output.getvalue()


# =====================================================
# POWER BI DATASET EXPORT HELPERS - V16.2
# =====================================================
def _clean_powerbi_columns(df):
    """Return a clean copy with Power BI-friendly column names and data types preserved."""
    if df is None or not isinstance(df, pd.DataFrame):
        return pd.DataFrame()
    out = df.copy()
    # Power BI handles spaces, but removing line breaks and repeated whitespace makes the model cleaner.
    out.columns = [re.sub(r"\s+", " ", str(c)).strip() for c in out.columns]
    return out


def _add_run_columns(df, run_id, scenario_name_value):
    out = _clean_powerbi_columns(df)
    if out.empty:
        return out
    if "Run ID" not in out.columns:
        out.insert(0, "Run ID", run_id)
    if "Scenario Name" not in out.columns:
        out.insert(1, "Scenario Name", scenario_name_value if scenario_name_value else "Not provided")
    return out


def _percent_from_counts(numerator, denominator):
    try:
        denominator = float(denominator)
        if denominator == 0:
            return np.nan
        return float(numerator) / denominator * 100
    except Exception:
        return np.nan


def build_powerbi_exception_rows(processed_df, cost_col, run_id, scenario_name_value):
    """Create one clean exception table for Power BI visuals."""
    rows = []
    if processed_df is None or processed_df.empty:
        return pd.DataFrame(columns=["Run ID", "Scenario Name", "SALES MODEL", "DEALER", "Dealer Code", "Region", "CCR TYPE", "Service Year", "Exception Type", "Exception Detail", "Cost", "SMU AT REBUILD"])

    base_cols = ["SALES MODEL", "DEALER", "Dealer Code", "Region", "CCR TYPE", "Service Year", "SMU AT REBUILD"]
    for _, r in processed_df.iterrows():
        base = {col: r.get(col, np.nan) for col in base_cols}
        base["Cost"] = r.get(cost_col, np.nan)
        if bool(r.get("Outlier", False)):
            rec = base.copy()
            rec["Exception Type"] = "Cost Outlier"
            rec["Exception Detail"] = r.get("Outlier Reason", "Cost outlier")
            rows.append(rec)
        if str(r.get("Cross-Type Exception Flag", "")).strip():
            rec = base.copy()
            rec["Exception Type"] = "Cross-Type Exception"
            rec["Exception Detail"] = r.get("Cross-Type Exception Flag", "")
            rows.append(rec)
        if str(r.get("Data Quality Exception Flag", "")).strip():
            rec = base.copy()
            rec["Exception Type"] = "Data Quality"
            rec["Exception Detail"] = r.get("Data Quality Exception Flag", "")
            rows.append(rec)
        if bool(r.get("Insufficient Sample Group", False)):
            rec = base.copy()
            rec["Exception Type"] = "Insufficient Sample"
            rec["Exception Detail"] = "Machine + CCR TYPE group has fewer than 5 records"
            rows.append(rec)
        if str(r.get("Dealer Rate Exception Flag", "")).strip():
            rec = base.copy()
            rec["Exception Type"] = "Dealer Rate Exception"
            rec["Exception Detail"] = r.get("Dealer Rate Exception Flag", "")
            rows.append(rec)
        if str(r.get("FX Source", "")).strip() == "Embedded fallback FX table":
            rec = base.copy()
            rec["Exception Type"] = "FX Fallback"
            rec["Exception Detail"] = "Embedded fallback FX table used"
            rows.append(rec)

    out = pd.DataFrame(rows)
    return _add_run_columns(out, run_id, scenario_name_value)


def build_powerbi_dealer_performance(valid_df, processed_df, cost_col, run_id, scenario_name_value):
    if valid_df is None or valid_df.empty:
        return pd.DataFrame()
    group_cols = ["SALES MODEL", "Dealer Code", "DEALER", "Region"]
    valid_cols = [c for c in group_cols if c in valid_df.columns]
    proc_cols = [c for c in group_cols if c in processed_df.columns]
    perf = valid_df.groupby(valid_cols, dropna=False).agg(
        Avg_Cost=(cost_col, "mean"),
        Avg_SMU=("SMU AT REBUILD", "mean"),
        Count=(cost_col, "count"),
        Cross_Type_Flags=("Cross-Type Exception Flag", lambda x: (x.astype(str).str.strip() != "").sum()),
        Dealer_Rate_Exception_Rows=("Dealer Rate Exception Flag", lambda x: (x.astype(str).str.strip() != "").sum()),
        Data_Quality_Rows=("Data Quality Exception Flag", lambda x: (x.astype(str).str.strip() != "").sum()),
    ).reset_index()
    full = processed_df.groupby(proc_cols, dropna=False).agg(
        Total_Rows=(cost_col, "count"),
        Outlier_Rows=("Outlier", "sum"),
    ).reset_index()
    out = perf.merge(full, on=valid_cols, how="left")
    out["Outlier Rate %"] = out.apply(lambda r: _percent_from_counts(r.get("Outlier_Rows", 0), r.get("Total_Rows", 0)), axis=1)
    out["Cross Flag Rate %"] = out.apply(lambda r: _percent_from_counts(r.get("Cross_Type_Flags", 0), r.get("Count", 0)), axis=1)
    out["Dealer Rate Exception Rate %"] = out.apply(lambda r: _percent_from_counts(r.get("Dealer_Rate_Exception_Rows", 0), r.get("Count", 0)), axis=1)
    out["Data Quality Rate %"] = out.apply(lambda r: _percent_from_counts(r.get("Data_Quality_Rows", 0), r.get("Count", 0)), axis=1)
    section_avg = out["Avg_Cost"].mean()
    out["Vs Overall Avg %"] = ((out["Avg_Cost"] - section_avg) / section_avg * 100) if section_avg else np.nan
    out["Performance Score"] = (
        100
        - out["Vs Overall Avg %"].clip(lower=0).fillna(0) * 0.50
        - out["Outlier Rate %"].fillna(0) * 0.30
        - out["Cross Flag Rate %"].fillna(0) * 0.20
        - out["Dealer Rate Exception Rate %"].fillna(0) * 0.20
        - out["Data Quality Rate %"].fillna(0) * 0.20
    ).clip(lower=0, upper=100).round(1)
    out["Performance Label"] = np.where(out["Performance Score"] >= 85, "Strong", np.where(out["Performance Score"] >= 70, "Monitor", "Review Needed"))
    return _add_run_columns(out, run_id, scenario_name_value)


def build_powerbi_region_performance(valid_df, processed_df, cost_col, run_id, scenario_name_value):
    if valid_df is None or valid_df.empty:
        return pd.DataFrame()
    group_cols = ["SALES MODEL", "Region"]
    perf = valid_df.groupby(group_cols, dropna=False).agg(
        Avg_Cost=(cost_col, "mean"),
        Avg_SMU=("SMU AT REBUILD", "mean"),
        Count=(cost_col, "count"),
        Cross_Type_Flags=("Cross-Type Exception Flag", lambda x: (x.astype(str).str.strip() != "").sum()),
        Data_Quality_Rows=("Data Quality Exception Flag", lambda x: (x.astype(str).str.strip() != "").sum()),
    ).reset_index()
    full = processed_df.groupby(group_cols, dropna=False).agg(
        Total_Rows=(cost_col, "count"),
        Outlier_Rows=("Outlier", "sum"),
    ).reset_index()
    out = perf.merge(full, on=group_cols, how="left")
    out["Outlier Rate %"] = out.apply(lambda r: _percent_from_counts(r.get("Outlier_Rows", 0), r.get("Total_Rows", 0)), axis=1)
    overall = out["Avg_Cost"].mean()
    out["Vs Overall Avg %"] = ((out["Avg_Cost"] - overall) / overall * 100) if overall else np.nan
    return _add_run_columns(out, run_id, scenario_name_value)


def build_powerbi_dataset_tables(analysis, scenario_name_value, export_reason_value, export_mode_value):
    """Build clean, flat Power BI-ready fact/dimension tables.

    This export intentionally avoids watermarks, merged cells, visual headers, and decorative formatting.
    Confidentiality and methodology values are included as fields in metadata tables instead.
    """
    processed_df = analysis["df"].copy()
    valid_df = analysis["valid"].copy()
    adjusted_valid = analysis.get("adjusted_valid", pd.DataFrame()).copy()
    cost_col = analysis["cost_col"]
    run_id = datetime.now().strftime("RUN_%Y%m%d_%H%M%S")
    scenario_label = scenario_name_value if scenario_name_value else "Not provided"

    fact_rebuild = _add_run_columns(processed_df, run_id, scenario_label)
    fact_valid_rebuild = _add_run_columns(valid_df, run_id, scenario_label)

    dim_machine = pd.DataFrame({"SALES MODEL": sorted(processed_df["SALES MODEL"].dropna().unique())}) if "SALES MODEL" in processed_df.columns else pd.DataFrame()
    if not dim_machine.empty:
        dim_machine["Machine Family"] = dim_machine["SALES MODEL"].astype(str).str.extract(r"(\d+)")[0].fillna(dim_machine["SALES MODEL"].astype(str))
        dim_machine = _add_run_columns(dim_machine, run_id, scenario_label)

    dealer_cols = [c for c in ["Dealer Code", "DEALER", "Region"] if c in processed_df.columns]
    dim_dealer = processed_df[dealer_cols].drop_duplicates().sort_values(dealer_cols).reset_index(drop=True) if dealer_cols else pd.DataFrame()
    dim_dealer = _add_run_columns(dim_dealer, run_id, scenario_label)

    dim_rebuild_type = analysis.get("rebuild_reference", pd.DataFrame()).rename(columns={"Description": "Rebuild Description"})
    dim_rebuild_type = _add_run_columns(dim_rebuild_type, run_id, scenario_label)

    dim_region = analysis.get("region_reference", pd.DataFrame()).rename(columns={"Configured Region": "Region"})
    if not dim_region.empty:
        dim_region["Configured Region Flag"] = True
    dim_region = _add_run_columns(dim_region, run_id, scenario_label)

    years = sorted(processed_df["Service Year"].dropna().astype(int).unique()) if "Service Year" in processed_df.columns else []
    dim_date = pd.DataFrame({"Service Year": years})
    if not dim_date.empty:
        dim_date["Year"] = dim_date["Service Year"]
        dim_date["Year Label"] = dim_date["Service Year"].astype(str)
    dim_date = _add_run_columns(dim_date, run_id, scenario_label)

    fact_exceptions = build_powerbi_exception_rows(processed_df, cost_col, run_id, scenario_label)
    fact_outliers = _add_run_columns(processed_df[processed_df["Outlier"] == True].copy(), run_id, scenario_label) if "Outlier" in processed_df.columns else pd.DataFrame()
    fact_cross_flags = _add_run_columns(valid_df[valid_df["Cross-Type Exception Flag"].astype(str).str.strip() != ""].copy(), run_id, scenario_label) if "Cross-Type Exception Flag" in valid_df.columns else pd.DataFrame()

    fact_machine_summary = _add_run_columns(analysis.get("machine_benchmark_ranking", pd.DataFrame()).copy(), run_id, scenario_label)
    fact_dealer_performance = build_powerbi_dealer_performance(valid_df, processed_df, cost_col, run_id, scenario_label)
    fact_region_performance = build_powerbi_region_performance(valid_df, processed_df, cost_col, run_id, scenario_label)

    fact_rate_coverage = _add_run_columns(analysis.get("dealer_rate_coverage_summary", pd.DataFrame()).copy(), run_id, scenario_label)
    fact_data_quality = _add_run_columns(analysis.get("data_quality_score_summary", pd.DataFrame()).copy(), run_id, scenario_label)
    data_quality_summary = _add_run_columns(analysis.get("data_quality_summary", pd.DataFrame()).copy(), run_id, scenario_label)

    run_metadata = analysis.get("metadata", pd.DataFrame()).copy()
    additional_metadata = pd.DataFrame({
        "Field": [
            "Run ID", "Scenario Name", "Export Mode", "Export Reason", "Power BI Export Format",
            "Power BI Notes", "Generated Timestamp"
        ],
        "Value": [
            run_id, scenario_label, export_mode_value, export_reason_value, "V16.2.2 Power BI Dataset Export",
            "Clean flat tables; no watermark rows, merged cells, decorative headers, or  marker rows. Use Run ID to relate scenario-specific tables.",
            datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        ]
    })
    run_metadata = pd.concat([additional_metadata, run_metadata], ignore_index=True)
    if "Field" in run_metadata.columns:
        run_metadata = run_metadata[~run_metadata["Field"].astype(str).str.contains("Confidential|Removed Label", case=False, na=False)].copy()
    if "Value" in run_metadata.columns:
        run_metadata = run_metadata[~run_metadata["Value"].astype(str).str.contains("", case=False, na=False)].copy()

    relationship_guide = pd.DataFrame({
        "From Table": [
            "Fact_RebuildRows", "Fact_RebuildRows", "Fact_RebuildRows", "Fact_RebuildRows", "Fact_RebuildRows",
            "Fact_MachineSummary", "Fact_DealerPerformance", "Fact_RegionPerformance", "Fact_ExceptionRows"
        ],
        "From Column": [
            "SALES MODEL", "Dealer Code", "CCR TYPE", "Region", "Service Year",
            "SALES MODEL", "Dealer Code", "Region", "SALES MODEL"
        ],
        "To Table": [
            "Dim_Machine", "Dim_Dealer", "Dim_RebuildType", "Dim_Region", "Dim_Date",
            "Dim_Machine", "Dim_Dealer", "Dim_Region", "Dim_Machine"
        ],
        "To Column": [
            "SALES MODEL", "Dealer Code", "CCR TYPE", "Region", "Service Year",
            "SALES MODEL", "Dealer Code", "Region", "SALES MODEL"
        ],
        "Relationship Notes": [
            "Many-to-one suggested", "Many-to-one suggested", "Many-to-one suggested", "Many-to-one suggested", "Many-to-one suggested",
            "Many-to-one suggested", "Many-to-one suggested", "Many-to-one suggested", "Many-to-one suggested"
        ]
    })

    dax_starter = pd.DataFrame({
        "Measure Name": [
            "Average Cost", "Valid Rows", "Total Rows", "Outlier Rows", "Outlier Rate %",
            "Cross-Type Flags", "Dealer Rate Exceptions", "Average SMU"
        ],
        "DAX Expression": [
            f"Average Cost = AVERAGE(Fact_RebuildRows[{cost_col}])",
            "Valid Rows = CALCULATE(COUNTROWS(Fact_RebuildRows), Fact_RebuildRows[Outlier] = FALSE())",
            "Total Rows = COUNTROWS(Fact_RebuildRows)",
            "Outlier Rows = CALCULATE(COUNTROWS(Fact_RebuildRows), Fact_RebuildRows[Outlier] = TRUE())",
            "Outlier Rate % = DIVIDE([Outlier Rows], [Total Rows])",
            "Cross-Type Flags = CALCULATE(COUNTROWS(Fact_RebuildRows), Fact_RebuildRows[Cross-Type Exception Flag] <> \"\")",
            "Dealer Rate Exceptions = CALCULATE(COUNTROWS(Fact_RebuildRows), Fact_RebuildRows[Dealer Rate Exception Flag] <> \"\")",
            "Average SMU = AVERAGE(Fact_RebuildRows[SMU AT REBUILD])",
        ]
    })

    tables = {
        "Fact_RebuildRows": fact_rebuild,
        "Fact_ValidRows": fact_valid_rebuild,
        "Dim_Machine": dim_machine,
        "Dim_Dealer": dim_dealer,
        "Dim_RebuildType": dim_rebuild_type,
        "Dim_Region": dim_region,
        "Dim_Date": dim_date,
        "Fact_ExceptionRows": fact_exceptions,
        "Fact_OutlierRows": fact_outliers,
        "Fact_CrossTypeFlags": fact_cross_flags,
        "Fact_MachineSummary": fact_machine_summary,
        "Fact_DealerPerf": fact_dealer_performance,
        "Fact_RegionPerf": fact_region_performance,
        "Fact_RateCoverage": fact_rate_coverage,
        "Fact_DataQuality": fact_data_quality,
        "DataQualitySummary": data_quality_summary,
        "Run_Metadata": _clean_powerbi_columns(run_metadata),
        "Parameters": _add_run_columns(analysis.get("parameter_summary", pd.DataFrame()).copy(), run_id, scenario_label),
        "Data_Dictionary": _clean_powerbi_columns(analysis.get("data_dictionary", pd.DataFrame()).copy()),
        "Known_Limitations": _clean_powerbi_columns(analysis.get("known_limitations", pd.DataFrame()).copy()),
        "Relationship_Guide": relationship_guide,
        "DAX_Starter": dax_starter,
    }
    return {name: table for name, table in tables.items() if table is not None and isinstance(table, pd.DataFrame)}


def _strip_powerbi_confidential_marker_rows(df):
    """Remove any accidental label marker rows before writing Power BI tables."""
    out = _clean_powerbi_columns(df)
    if out.empty:
        return out
    first_col = out.columns[0]
    marker_mask = out[first_col].astype(str).str.contains("Confidential|Yellow|Removed Label", case=False, na=False)
    return out.loc[~marker_mask].copy()


def write_powerbi_dataset_workbook(writer, analysis, scenario_name_value, export_reason_value, export_mode_value):
    tables = build_powerbi_dataset_tables(analysis, scenario_name_value, export_reason_value, export_mode_value)
    for sheet_name, table in tables.items():
        safe_name = safe_sheet_name(sheet_name)
        clean_table = _strip_powerbi_confidential_marker_rows(table)
        # Explicit startrow=0 ensures row 1 is always the actual column header row for Power BI.
        clean_table.to_excel(writer, sheet_name=safe_name, index=False, startrow=0)
        ws = writer.book[safe_name]
        # Freeze on row 2 for convenience only; no inserted rows, merged cells, or watermarking.
        ws.freeze_panes = "A2"

# =====================================================
# MAIN UI INPUTS
# =====================================================
st.markdown(
    """
    <div class="method-card">
        <div class="method-title">Active Analysis Logic</div>
        <div class="method-body">Cost = PARTS DN USD + Labor Cost USD. Inflation is applied by component. Outliers use log(cost) + IQR by Machine + CCR TYPE. CPT+H cross-type flags compare against machine-level CMR benchmark.</div>
    </div>
    """,
    unsafe_allow_html=True,
)

rebuild_file = st.file_uploader("Upload Rebuild File", type=["xlsx"])
if st.button("Clear uploaded data / reset analysis"):
    reset_app_state()
    st.success("Session analysis state cleared. Refresh the page if you also want to clear the file uploader selection.")
    st.stop()
validate_uploaded_file_security(rebuild_file)
with st.expander("Methodology Lock", expanded=False):
    st.markdown(METHOD_LOCK_TEXT)

st.subheader("Currency Conversion")
auto_currency = st.checkbox("Automatically convert source currency to USD", True)
default_source_currency = st.text_input("Default source currency if no currency column exists", "USD").upper().strip()
dealer_rate_currency_mode = st.selectbox("Dealer labor rate currency", ["USD", "Same as source currency"], index=0)

st.subheader("Dealer Labor Rates")
use_default = st.checkbox("Use Built-in Default Dealer Rates", True)
rate_file = None
rate_file_format = "Auto-detect"
rate_fallback_behavior = "Use yearly average fallback (recommended)"

st.download_button(
    "Download Dealer Rate Template",
    data=create_dealer_rate_template(),
    file_name="Dealer_Rate_Template.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)

if use_default:
    built_in_rate_file_available = Path("expanded_dealer_base_rate_by_year_2016_2026.xlsx").exists()
    if built_in_rate_file_available:
        st.warning("Built-in expanded dealer base rates 2016-2026 are being used from the bundled workbook. You can still upload a custom dealer labor rate workbook if needed.")
    else:
        st.error("Built-in dealer-rate workbook is missing from the app folder. The app can still run using emergency generic fallback rates, but production analysis should not rely on that fallback.")
else:
    rate_file_format = st.selectbox(
        "Dealer rate file format",
        ["Auto-detect", "Flat Table", "Multi-Sheet Dealer Workbook"],
        index=0,
        help="Recommended format is Flat Table with Dealer Code, Service Year, Rate, optional Rate Currency, and Notes.",
    )
    rate_fallback_behavior = st.selectbox(
        "Missing dealer-year rate fallback behavior",
        [
            "Use yearly average fallback (recommended)",
            "Use overall average fallback",
            "Use built-in default fallback",
            "Stop analysis if dealer-year rate is missing",
        ],
        index=0,
        help="Recommended default uses the uploaded rate file's yearly average when a dealer-year match is missing, and flags the row.",
    )
    rate_file = st.file_uploader("Upload Custom Dealer Labor Rate Workbook", type=["xlsx"])
    validate_uploaded_file_security(rate_file)
    if rate_file:
        try:
            preview_rates = build_rate_table_from_workbook(rate_file, rate_file_format)
            st.write("### Dealer Rate Upload Validation")
            display_table(validate_dealer_rate_table(preview_rates), number_cols=["Value"])
            st.write("### Dealer Rate Preview")
            display_table(preview_rates.head(50), currency_cols=["Rate"], number_cols=["Service Year"])
            st.caption("Preferred flat table columns: Dealer Code, Service Year, Rate, Rate Currency, Notes. Legacy multi-sheet dealer workbooks are still supported.")
        except Exception as exc:
            st.error(f"Unable to parse dealer rate workbook: {exc}")
            st.stop()

apply_inflation = st.checkbox("Apply BLS CPI-U Inflation", True)
base_year = st.number_input("Base Year for Inflation", 2010, 2030, 2026)
machine_input = st.text_input("Machines (optional, comma-separated; leave blank for all)")
start_year = st.number_input("Start Year", 2010, 2030, 2016)
end_year = st.number_input("End Year", 2010, 2030, 2026)
rebuild_filter = st.multiselect("Filter Rebuild Types", options=list(CERTIFIED_REBUILD_TYPES.keys()), default=["CMR", "CPT+H", "CPT-O"])
region_filter = st.multiselect("Filter Regions", options=VALID_REGIONS + ["UNKNOWN", "OTHER"], default=VALID_REGIONS + ["UNKNOWN", "OTHER"])

st.subheader("Scenario, Governance, and Future-Proof Controls")
scenario_name = st.text_input("Analysis Scenario Name", "")
user_role_view = st.selectbox("Role view / intended user type", ["Analyst", "Viewer", "Admin"], index=0, help="Advisory only unless enterprise authentication is added.")
strict_mode = st.checkbox("Strict Mode: stop analysis when configured quality thresholds are not met", False)
with st.expander("Advanced Threshold Configuration", expanded=False):
    st.caption("Defaults preserve the current methodology. Change these only when there is a business-approved reason.")
    cross_type_threshold_multiplier = st.number_input("CPT+H cross-type threshold multiplier", min_value=1.00, max_value=2.00, value=1.10, step=0.01)
    min_cmr_rows_for_benchmark = st.number_input("Minimum valid CMR rows for CPT+H benchmark", min_value=1, max_value=20, value=3, step=1)
    dealer_rate_coverage_warning_threshold = st.number_input("Dealer rate coverage warning threshold %", min_value=0.0, max_value=100.0, value=95.0, step=1.0)
    dealer_rate_coverage_strict_threshold = st.number_input("Dealer rate coverage strict threshold %", min_value=0.0, max_value=100.0, value=80.0, step=1.0)
    data_quality_strict_min_score = st.number_input("Strict Mode minimum data quality score", min_value=0.0, max_value=100.0, value=75.0, step=1.0)
    max_rows_warning_threshold = st.number_input("Large-run row warning threshold", min_value=1000, max_value=250000, value=DEFAULT_MAX_ROWS_WARNING, step=1000)

st.subheader("Export Controls")
export_mode = st.selectbox("Full workbook export mode", ["Full Analysis Workbook", "Summary Only", "Exceptions Only", "Dealer Rate Audit", "Power BI Dataset Export"], index=0)
export_reason = st.selectbox("Export reason", ["Manager review", "Dealer review", "Cost benchmarking", "Data validation", "Presentation support", "Other"], index=0)
export_reason_other = ""
if export_reason == "Other":
    export_reason_other = st.text_input("Describe export reason", "")
export_reason_final = export_reason_other.strip() if export_reason == "Other" and export_reason_other.strip() else export_reason
if user_role_view == "Viewer" and export_mode not in ["Summary Only", "Power BI Dataset Export"]:
    st.warning("Viewer role view is intended for limited distribution. Export mode has been changed to Summary Only for this session.")
    export_mode = "Summary Only"

if rebuild_file:
    with st.expander("Pre-Run Data Validation", expanded=True):
        profile, issue_df = validate_uploaded_workbook(rebuild_file)
        if profile is not None:
            st.write("### Workbook Profile")
            display_table(profile, number_cols=["Rows"])
        st.write("### Validation Issues")
        if issue_df.empty:
            st.success("No pre-run validation issues detected.")
        else:
            st.warning("Review validation issues before running analysis.")
            st.dataframe(issue_df, use_container_width=True)

if rebuild_file:
    st.write("### Preliminary Run Readiness Check")
    prelim_rows = [{"Readiness Check": "Workbook Uploaded", "Status": "Ready", "Details": "A rebuild workbook is loaded."}, {"Readiness Check": "Dealer Rate Source", "Status": "Ready", "Details": "Built-in expanded rates selected." if use_default else "Custom dealer rate workbook selected. Review upload validation."}]
    display_table(pd.DataFrame(prelim_rows))

if st.button("Run Analysis"):
    st.session_state.run_clicked = True
    st.session_state.analysis = None

# =====================================================
# ANALYSIS CORE
# =====================================================
def run_analysis(rebuild_file, rate_file):
    rebuild = read_rebuild_workbook(rebuild_file)
    frames = []
    for sheet_name, sheet in rebuild.items():
        sheet = sheet.copy()
        existing_ccr = None
        for col in sheet.columns:
            if str(col).upper().strip() in ["CCR TYPE", "CCRTYPE", "CCR_TYPE"]:
                existing_ccr = col
                break
        if existing_ccr:
            sheet["CCR TYPE"] = sheet[existing_ccr].apply(lambda value: normalize_ccr_type(value, sheet_name))
        else:
            sheet["CCR TYPE"] = normalize_ccr_type(sheet_name, sheet_name)
        frames.append(sheet)

    df = pd.concat(frames, ignore_index=True)
    required = ["DELIVERED DATE", "ENROLL DATE", "IN SERVICE DATE", "SALES MODEL", "DEALER", "PARTS DN", "SMU AT REBUILD", "REBUILD WORK HRS"]
    missing = [col for col in required if col not in df.columns]
    if missing:
        raise ValueError(f"Missing required columns: {missing}")

    rows_uploaded = len(df)
    df = df[df["CCR TYPE"].isin(rebuild_filter)].copy()
    df["Service Date"] = df["DELIVERED DATE"].fillna(df["ENROLL DATE"]).fillna(df["IN SERVICE DATE"])
    df["Service Year"] = pd.to_datetime(df["Service Date"], errors="coerce").dt.year
    missing_service_date_count = int(df["Service Year"].isna().sum())
    df = df[(df["Service Year"] >= start_year) & (df["Service Year"] <= end_year)].copy()
    if machine_input:
        machine_list = [machine.strip() for machine in machine_input.split(",")]
        df = df[df["SALES MODEL"].isin(machine_list)].copy()
    if df.empty:
        raise ValueError("No rows remain after filters. Adjust machine, year, region, or rebuild type filters.")

    df["Dealer Code"] = df["DEALER"].astype(str).str.upper().str.extract(r"([A-Z][A-Z0-9]{3})")
    missing_dealer_code_count = int(df["Dealer Code"].isna().sum())
    df["Region"] = df["REGION"].apply(normalize_region) if "REGION" in df.columns else "UNKNOWN"
    df["Region Display"] = df["Region"].where(df["Region"].isin(VALID_REGIONS + ["UNKNOWN"]), "OTHER")
    unknown_region_count = int((df["Region Display"] == "OTHER").sum())
    df = df[df["Region Display"].isin(region_filter)].copy()
    if df.empty:
        raise ValueError("No rows remain after region filter.")

    currency_col = detect_currency_column(df)
    if auto_currency and currency_col:
        df["Source Currency"] = df[currency_col].apply(lambda value: normalize_currency_code(value, default_source_currency))
    else:
        df["Source Currency"] = normalize_currency_code(default_source_currency)
    df["Service Year"] = pd.to_numeric(df["Service Year"], errors="coerce").astype("Int64")
    fx_lookup = build_fx_lookup(df["Source Currency"].dropna().unique(), df["Service Year"].dropna().unique())
    df = df.merge(fx_lookup, how="left", left_on=["Source Currency", "Service Year"], right_on=["Currency", "Service Year"])
    df["FX to USD"] = df["FX to USD"].fillna(1.0)
    df["FX Source"] = df["FX Source"].fillna("USD/no FX conversion")
    fallback_fx_count = int((df["FX Source"] == "Embedded fallback FX table").sum())

    if use_default or rate_file is None:
        rate_df = build_default_rate_table(2010, 2030)
        rate_file_mode = "Built-in Expanded Dealer Rates 2016-2026"
        effective_fallback_behavior = "Built-in expanded dealer rates; missing dealer-years use yearly average fallback, then overall average if needed"
    else:
        rate_df = build_rate_table_from_workbook(rate_file, rate_file_format)
        if rate_df.empty:
            raise ValueError("Custom dealer rate workbook did not contain any valid dealer-year rates. Use the downloadable template or select the correct rate format.")
        rate_file_mode = rate_file_format
        effective_fallback_behavior = rate_fallback_behavior

    rate_df["Service Year"] = pd.to_numeric(rate_df["Service Year"], errors="coerce").astype("Int64")
    rate_df["Rate"] = pd.to_numeric(rate_df["Rate"], errors="coerce")
    df = df.merge(rate_df, how="left", on=["Dealer Code", "Service Year"])
    df["Rate Source"] = "Dealer-Year Rate"
    df["Dealer Rate Exception Flag"] = ""
    missing_dealer_year = df["Rate"].isna()
    df.loc[missing_dealer_year, "Dealer Rate Exception Flag"] = "Missing dealer-year labor rate"

    if missing_dealer_year.any() and effective_fallback_behavior == "Stop analysis if dealer-year rate is missing":
        missing_count = int(missing_dealer_year.sum())
        raise ValueError(f"{missing_count} rows are missing dealer-year labor rates. Upload a complete rate file or choose a fallback behavior.")

    if missing_dealer_year.any():
        if effective_fallback_behavior == "Use built-in default fallback":
            default_rates = build_default_rate_table(2010, 2030).set_index("Service Year")["Rate"]
            df.loc[missing_dealer_year, "Rate"] = df.loc[missing_dealer_year, "Service Year"].map(default_rates)
            df.loc[missing_dealer_year, "Rate Source"] = "Built-in Default Fallback Rate"
        elif effective_fallback_behavior == "Use overall average fallback":
            overall_avg = rate_df["Rate"].mean()
            df.loc[missing_dealer_year, "Rate"] = overall_avg
            df.loc[missing_dealer_year, "Rate Source"] = "Overall Average Fallback Rate"
        else:
            yearly_avg = rate_df.groupby("Service Year")["Rate"].mean()
            df.loc[missing_dealer_year, "Rate"] = df.loc[missing_dealer_year, "Service Year"].map(yearly_avg)
            df.loc[missing_dealer_year, "Rate Source"] = "Global Yearly Fallback Rate"
            still_missing = df["Rate"].isna()
            if still_missing.any():
                overall_avg = rate_df["Rate"].mean()
                df.loc[still_missing, "Rate"] = overall_avg
                df.loc[still_missing, "Rate Source"] = "Overall Average Fallback Rate"
                df.loc[still_missing, "Dealer Rate Exception Flag"] = "Missing dealer-year and service-year average labor rate"

    missing_overall = df["Rate"].isna()
    if missing_overall.any():
        fallback_rate = build_default_rate_table(2010, 2030)["Rate"].mean()
        df.loc[missing_overall, "Rate"] = fallback_rate
        df.loc[missing_overall, "Rate Source"] = "Emergency Built-in Overall Fallback Rate"
        df.loc[missing_overall, "Dealer Rate Exception Flag"] = "Emergency fallback labor rate used"

    df["Base Rate Year Used"] = df["Service Year"]
    df["Avg Base Rate Used"] = df["Rate"]

    df["PARTS DN"] = pd.to_numeric(df["PARTS DN"], errors="coerce")
    df["REBUILD WORK HRS"] = pd.to_numeric(df["REBUILD WORK HRS"], errors="coerce")
    df["SMU AT REBUILD"] = pd.to_numeric(df["SMU AT REBUILD"], errors="coerce")
    missing_parts_count = int(df["PARTS DN"].isna().sum())
    missing_labor_hours_count = int(df["REBUILD WORK HRS"].isna().sum())
    df = df[(df["PARTS DN"] > 0) & (df["REBUILD WORK HRS"].notna())].copy()

    df["PARTS DN USD"] = df["PARTS DN"] * df["FX to USD"]
    df["Rate FX to USD"] = df["FX to USD"] if dealer_rate_currency_mode == "Same as source currency" else 1.0
    df["Rate USD"] = df["Rate"] * df["Rate FX to USD"]
    df["Labor Cost USD"] = df["REBUILD WORK HRS"] * df["Rate USD"]
    df["Adjusted Total Cost USD"] = df["PARTS DN USD"] + df["Labor Cost USD"]

    cpi_table, cpi_source = get_cpi_table(start_year, end_year, base_year)
    base_cpi = cpi_table.get(int(base_year), np.nan)
    if apply_inflation and pd.notna(base_cpi):
        df["Service Year CPI"] = df["Service Year"].apply(lambda year: cpi_table.get(int(year), np.nan) if pd.notnull(year) else np.nan)
        df["Inflation Factor"] = (base_cpi / df["Service Year CPI"]).fillna(1.0)
    else:
        df["Service Year CPI"] = np.nan
        df["Inflation Factor"] = 1.0
    df["Inflation-Adjusted Parts DN USD"] = df["PARTS DN USD"] * df["Inflation Factor"]
    df["Inflation-Adjusted Base Rate USD"] = df["Rate USD"] * df["Inflation Factor"]
    df["Inflation-Adjusted Labor Cost USD"] = df["REBUILD WORK HRS"] * df["Inflation-Adjusted Base Rate USD"]
    df["Inflation-Adjusted Adjusted Total Cost USD"] = df["Inflation-Adjusted Parts DN USD"] + df["Inflation-Adjusted Labor Cost USD"]
    cost_col = "Inflation-Adjusted Adjusted Total Cost USD" if apply_inflation else "Adjusted Total Cost USD"
    df = df[df[cost_col] > 0].copy()
    if df.empty:
        raise ValueError("No valid cost rows remain after processing.")

    df["Data Quality Exception Flag"] = ""
    df.loc[df["SMU AT REBUILD"].isin([0, 1]), "Data Quality Exception Flag"] = "SMU 0 or 1"
    df["Outlier"] = False
    df["Outlier Reason"] = ""
    df["Insufficient Sample Group"] = False
    for (machine, ccr_type), group in df.groupby(["SALES MODEL", "CCR TYPE"], dropna=False):
        idx = group.index
        n = len(group)
        if n < 5:
            df.loc[idx, "Insufficient Sample Group"] = True
            df.loc[idx, "Outlier Reason"] = "Insufficient sample; no statistical outlier removal"
            continue
        log_vals = np.log(group[cost_col])
        q1, q3 = np.percentile(log_vals, [25, 75])
        iqr = q3 - q1
        mult = 2.0 if n < 8 else 1.5
        low = q1 - mult * iqr
        high = q3 + mult * iqr
        mask = (log_vals < low) | (log_vals > high)
        df.loc[idx, "Outlier"] = mask
        df.loc[idx[mask], "Outlier Reason"] = "Cost outlier: log(cost) IQR rule"

    valid = df[df["Outlier"] == False].copy()
    if valid.empty:
        raise ValueError("All rows were excluded as cost outliers. Review data or filters.")

    df["Cross-Type Exception Flag"] = ""
    df["CMR Benchmark Cost"] = np.nan
    for machine, group in df.groupby("SALES MODEL", dropna=False):
        cmr_valid = group[(group["CCR TYPE"] == "CMR") & (group["Outlier"] == False)]
        if len(cmr_valid) >= int(min_cmr_rows_for_benchmark):
            benchmark = cmr_valid[cost_col].median()
            df.loc[group.index, "CMR Benchmark Cost"] = benchmark
            mask = group["CCR TYPE"].eq("CPT+H") & group["Outlier"].eq(False) & (group[cost_col] > benchmark * float(cross_type_threshold_multiplier))
            df.loc[group.index[mask], "Cross-Type Exception Flag"] = "CPT+H Cost Above Typical CMR"

    valid = df[df["Outlier"] == False].copy()
    summary = valid.groupby("CCR TYPE").agg(Avg_Cost=(cost_col, "mean"), Avg_SMU=("SMU AT REBUILD", "mean"), Count=(cost_col, "count")).reset_index()
    summary = add_vs_cmr(summary)
    summary["Sample Confidence"] = summary["Count"].apply(sample_confidence)
    adjusted_valid = valid[~((valid["CCR TYPE"] == "CPT+H") & (valid["Cross-Type Exception Flag"] == "CPT+H Cost Above Typical CMR"))].copy()
    adjusted_summary = adjusted_valid.groupby("CCR TYPE").agg(Adjusted_Avg_Cost=(cost_col, "mean"), Count=(cost_col, "count")).reset_index()
    adjusted_summary["CCR TYPE"] = adjusted_summary["CCR TYPE"].replace({"CPT+H": "CPT+H Adjusted"})
    adjusted_summary["Sample Confidence"] = adjusted_summary["Count"].apply(sample_confidence)
    rebuild_reference = pd.DataFrame([{"CCR TYPE": key, "Description": value} for key, value in CERTIFIED_REBUILD_TYPES.items()])
    region_reference = pd.DataFrame({"Configured Region": VALID_REGIONS})
    metadata = pd.DataFrame({"Field": ["App Version", "Run Timestamp", "Rows Uploaded", "Rows After Filters", "Valid Rows", "Start Year", "End Year", "Base Year", "Machine Filter", "Rebuild Type Filter", "Region Filter", "Default Source Currency", "Dealer Rate Currency Mode", "Currency Column Detected", "BLS CPI Source", "FX Source", "Analysis Cost Used", "Dealer Rate Mode", "Dealer Rate Format", "Dealer Rate Fallback Behavior", "Dealer Rate Rows", "Dealer Rate Unique Dealers", "Scenario Name", "User Role View", "Strict Mode", "Methodology Version", "Outlier Rule Version", "Dealer Rate Version", "Security Control Version", "Export Format Version", "CPT+H Threshold Multiplier", "Minimum CMR Benchmark Rows", "Dealer Rate Coverage Warning Threshold %", "Dealer Rate Coverage Strict Threshold %", "Data Quality Strict Minimum Score", "Large-Run Row Warning Threshold"], "Value": [APP_VERSION, datetime.now().strftime("%Y-%m-%d %H:%M:%S"), rows_uploaded, len(df), len(valid), start_year, end_year, base_year, machine_input if machine_input else "All", ", ".join(rebuild_filter), ", ".join(region_filter), default_source_currency, dealer_rate_currency_mode, currency_col if currency_col else "None", cpi_source, "Frankfurter annual average; embedded fallback if unavailable", cost_col, "Built-in Expanded 2016-2026" if use_default or rate_file is None else "Uploaded Custom", rate_file_mode, effective_fallback_behavior, len(rate_df), rate_df["Dealer Code"].nunique(), scenario_name if scenario_name else "Not provided", user_role_view, "Yes" if strict_mode else "No", METHODOLOGY_VERSION, OUTLIER_RULE_VERSION, DEALER_RATE_VERSION, SECURITY_CONTROL_VERSION, EXPORT_FORMAT_VERSION, cross_type_threshold_multiplier, min_cmr_rows_for_benchmark, dealer_rate_coverage_warning_threshold, dealer_rate_coverage_strict_threshold, data_quality_strict_min_score, max_rows_warning_threshold]})
    
    if "Field" in metadata.columns:
        metadata = metadata[~metadata["Field"].astype(str).str.contains("Confidential|Removed Label", case=False, na=False)].copy()
    if "Value" in metadata.columns:
        metadata = metadata[~metadata["Value"].astype(str).str.contains("Confidential|Yellow", case=False, na=False)].copy()
    data_quality_summary = pd.DataFrame({"Metric": ["Missing Service Date", "Missing Dealer Code", "Missing Parts DN", "Missing Labor Hours", "SMU 0 or 1", "Unknown/Other Regions", "Rows Using Fallback FX", "Rows Using Global Yearly Fallback Rate", "Rows Using Overall Average Fallback Rate", "Rows Using Built-in Default Fallback Rate", "Dealer Rate Exception Rows"], "Value": [missing_service_date_count, missing_dealer_code_count, missing_parts_count, missing_labor_hours_count, int(df["Data Quality Exception Flag"].eq("SMU 0 or 1").sum()), unknown_region_count, fallback_fx_count, int((df["Rate Source"] == "Global Yearly Fallback Rate").sum()), int((df["Rate Source"] == "Overall Average Fallback Rate").sum()), int((df["Rate Source"] == "Built-in Default Fallback Rate").sum()), int((df["Dealer Rate Exception Flag"] != "").sum())]})
    dealer_rate_coverage_summary = build_dealer_rate_coverage_summary(df)
    data_quality_score_summary = build_data_quality_score_summary(df, outlier_count=int(df["Outlier"].sum()), insufficient_count=int(df["Insufficient Sample Group"].sum()))
    run_readiness_summary = build_run_readiness_summary(df, dealer_rate_coverage_summary, data_quality_score_summary)
    parameter_summary = build_parameter_summary_table()
    known_limitations = build_known_limitations_table()
    data_dictionary = build_data_dictionary_table()
    role_policy = build_role_policy_table()
    performance_safeguards = build_performance_safeguard_summary(rows_uploaded, len(df))
    evaluate_strict_mode(run_readiness_summary, dealer_rate_coverage_summary, data_quality_score_summary, len(df))
    show_adjusted_cpth = has_both_cmr_cpth(valid)
    machine_benchmark_ranking = build_machine_benchmark_ranking(valid, df, cost_col)
    executive_narrative = build_executive_narrative(valid, summary, cost_col, int(df["Outlier"].sum()), int((valid["Cross-Type Exception Flag"] != "").sum()), dealer_rate_coverage_summary, data_quality_score_summary, show_adjusted_cpth)

    return {"df": df, "valid": valid, "adjusted_valid": adjusted_valid, "summary": summary, "adjusted_summary": adjusted_summary, "cost_col": cost_col, "cpi_table": cpi_table, "cpi_source": cpi_source, "base_cpi": base_cpi, "fx_lookup": fx_lookup, "currency_col": currency_col, "rebuild_reference": rebuild_reference, "region_reference": region_reference, "metadata": metadata, "data_quality_summary": data_quality_summary, "outlier_count": int(df["Outlier"].sum()), "cross_count": int((valid["Cross-Type Exception Flag"] != "").sum()), "dq_count": int(df["Data Quality Exception Flag"].eq("SMU 0 or 1").sum()), "insufficient_count": int(df["Insufficient Sample Group"].sum()), "global_year_fallback_count": int((df["Rate Source"] == "Global Yearly Fallback Rate").sum()), "overall_fallback_count": int((df["Rate Source"] == "Overall Average Fallback Rate").sum()), "rate_df": rate_df, "dealer_rate_validation": validate_dealer_rate_table(rate_df, start_year, end_year), "dealer_rate_exception_rows": df[df["Dealer Rate Exception Flag"] != ""].copy(), "rate_file_mode": rate_file_mode, "effective_fallback_behavior": effective_fallback_behavior, "dealer_rate_coverage_summary": dealer_rate_coverage_summary, "data_quality_score_summary": data_quality_score_summary, "run_readiness_summary": run_readiness_summary, "show_adjusted_cpth": show_adjusted_cpth, "machine_benchmark_ranking": machine_benchmark_ranking, "executive_narrative": executive_narrative, "parameter_summary": parameter_summary, "known_limitations": known_limitations, "data_dictionary": data_dictionary, "role_policy": role_policy, "performance_safeguards": performance_safeguards}

# =====================================================
# RENDER RESULTS
# =====================================================
if st.session_state.run_clicked and rebuild_file:
    if st.session_state.analysis is None:
        try:
            st.session_state.analysis = run_analysis(rebuild_file, rate_file)
        except Exception as exc:
            st.error(str(exc))
            st.stop()

    analysis = st.session_state.analysis
    df = analysis["df"]
    valid = analysis["valid"]
    summary = analysis["summary"]
    adjusted_summary = analysis["adjusted_summary"]
    cost_col = analysis["cost_col"]
    cpi_table = analysis["cpi_table"]
    cpi_source = analysis["cpi_source"]
    base_cpi = analysis["base_cpi"]
    fx_lookup = analysis["fx_lookup"]
    currency_col = analysis["currency_col"]
    rebuild_reference = analysis["rebuild_reference"]
    region_reference = analysis["region_reference"]
    metadata = analysis["metadata"]
    data_quality_summary = analysis["data_quality_summary"]
    outlier_count = analysis["outlier_count"]
    cross_count = analysis["cross_count"]
    dq_count = analysis["dq_count"]
    insufficient_count = analysis["insufficient_count"]
    global_year_fallback_count = analysis["global_year_fallback_count"]
    overall_fallback_count = analysis["overall_fallback_count"]
    rate_df = analysis.get("rate_df", pd.DataFrame())
    dealer_rate_validation = analysis.get("dealer_rate_validation", pd.DataFrame())
    dealer_rate_exception_rows = analysis.get("dealer_rate_exception_rows", pd.DataFrame())
    rate_file_mode = analysis.get("rate_file_mode", "Unknown")
    effective_fallback_behavior = analysis.get("effective_fallback_behavior", "Unknown")
    dealer_rate_coverage_summary = analysis.get("dealer_rate_coverage_summary", pd.DataFrame())
    data_quality_score_summary = analysis.get("data_quality_score_summary", pd.DataFrame())
    run_readiness_summary = analysis.get("run_readiness_summary", pd.DataFrame())
    show_adjusted_cpth = analysis.get("show_adjusted_cpth", False)
    machine_benchmark_ranking = analysis.get("machine_benchmark_ranking", pd.DataFrame())
    executive_narrative = analysis.get("executive_narrative", [])
    parameter_summary = analysis.get("parameter_summary", pd.DataFrame())
    known_limitations = analysis.get("known_limitations", pd.DataFrame())
    data_dictionary = analysis.get("data_dictionary", pd.DataFrame())
    role_policy = analysis.get("role_policy", pd.DataFrame())
    performance_safeguards = analysis.get("performance_safeguards", pd.DataFrame())

    tabs = st.tabs(["Executive Dashboard", "Machine Detail", "Dealer Performance", "Region Performance", "Exceptions & Data Quality", "Executive Insights", "How to Use", "Methodology", "Governance & Dictionary", "Reference"])
    tab1, tab2, tab3, tab4, tab5, tab6, tab7, tab8, tab9, tab10 = tabs

    with tab1:
        st.subheader("Run Summary")
        display_table(metadata)
        st.write("### Run Readiness Check")
        display_table(run_readiness_summary)
        st.write("### Saved Parameter Summary")
        display_table(parameter_summary)
        st.write("### Performance Safeguards")
        display_table(performance_safeguards)
        st.write("### Dealer Rate Coverage Score")
        display_table(dealer_rate_coverage_summary)
        st.write("### Data Quality Score")
        display_table(data_quality_score_summary)
        st.subheader("Executive Summary")
        col1, col2, col3, col4 = st.columns(4)
        col1.metric("Valid Rows", f"{len(valid):,}")
        col2.metric("Cost Outliers", f"{outlier_count:,}")
        col3.metric("Avg USD Analysis Cost", money(valid[cost_col].mean()))
        col4.metric("Cross-Type Flags", f"{cross_count:,}")
        st.write("### Standard Rebuild Type Average Cost")
        display_table(summary, currency_cols=["Avg_Cost"], percent_cols=["Vs CMR %"], smu_cols=["Avg_SMU"], number_cols=["Count"])
        if show_adjusted_cpth:
            st.write("### Adjusted Rebuild Type Average Cost")
            display_table(adjusted_summary, currency_cols=["Adjusted_Avg_Cost"], number_cols=["Count"])
        else:
            st.info("Adjusted CPT+H comparison is hidden because both CMR and CPT+H records are required for this comparison.")
        st.write("### Machine Benchmark Ranking")
        display_table(machine_benchmark_ranking, currency_cols=["Avg_Cost"], percent_cols=["Outlier Rate %", "Dealer Rate Exception Rate %", "Vs Overall Avg %"], smu_cols=["Avg_SMU"], number_cols=["Valid_Rows", "Total_Rows", "Outlier_Rows", "Cross_Type_Flags", "Dealer_Rate_Exception_Rows", "Priority Score"])
        st.write("### SMU vs USD Analysis Cost by Rebuild Type")
        scatter = valid[["SMU AT REBUILD", cost_col, "CCR TYPE"]].dropna().rename(columns={"SMU AT REBUILD": "SMU", cost_col: "Cost"})
        cat_scatter_chart(scatter, "SMU", "Cost", "CCR TYPE", tooltip=["CCR TYPE", "SMU", "Cost"])
        st.write("### Currency and Inflation Settings")
        settings = pd.DataFrame({"Setting": ["Currency Column Detected", "Default Source Currency", "Dealer Rate Currency Mode", "FX Source", "CPI Source", "BLS Series", "Inflation Applied", "Base Year", "Base Year CPI", "Cost Source", "Labor Cost", "PLUS PARTS DN", "Analysis Cost Used"], "Value": [str(currency_col) if currency_col else "None - default source currency used", default_source_currency, dealer_rate_currency_mode, "Frankfurter annual average; embedded fallback if unavailable", cpi_source, "CUUR0000SA0 - CPI-U All Items, U.S. city average, not seasonally adjusted", "Yes" if apply_inflation else "No", str(base_year), f"{base_cpi:,.3f}" if pd.notna(base_cpi) else "N/A", "PARTS DN USD + Labor Cost USD", "Labor Hours × Dealer Service-Year Rate converted to USD where applicable", "Ignored entirely", cost_col]})
        st.dataframe(settings, use_container_width=True)

    with tab2:
        st.subheader("Machine Detail")
        machine_list = sorted(valid["SALES MODEL"].dropna().unique())
        selected_machine = st.selectbox("Select Machine", machine_list, key="machine_detail_selector")
        dfm = valid[valid["SALES MODEL"] == selected_machine].copy()
        dfm_all = df[df["SALES MODEL"] == selected_machine].copy()
        machine_adjusted_valid = dfm[~((dfm["CCR TYPE"] == "CPT+H") & (dfm["Cross-Type Exception Flag"] == "CPT+H Cost Above Typical CMR"))].copy()
        st.write(f"### Machine: {selected_machine}")
        mcol1, mcol2, mcol3, mcol4 = st.columns(4)
        machine_show_adjusted_cpth = has_both_cmr_cpth(dfm)
        mcol1.metric("Standard Avg Cost", money(dfm[cost_col].mean()))
        if machine_show_adjusted_cpth:
            mcol2.metric("Adjusted Avg Cost", money(machine_adjusted_valid[cost_col].mean()))
            adj_cpth = machine_adjusted_valid[machine_adjusted_valid["CCR TYPE"] == "CPT+H"][cost_col].mean()
            mcol3.metric("Adjusted CPT+H Avg", money(adj_cpth) if pd.notna(adj_cpth) else "N/A")
        else:
            mcol2.metric("Adjusted Avg Cost", "N/A")
            mcol3.metric("Adjusted CPT+H Avg", "N/A")
        mcol4.metric("Machine Outliers", f"{int(dfm_all['Outlier'].sum()):,}")

        machine_export_bytes = build_machine_export(selected_machine, analysis)
        safe_machine = str(selected_machine).replace("/", "-").replace("\\", "-").replace(" ", "_")
        if render_export_acknowledgement("machine_export_ack"):
            st.download_button(
                "Download Selected Machine Workbook",
                data=machine_export_bytes,
                file_name=f"Rebuild_Analysis_{safe_machine}_{datetime.now().strftime('%Y-%m-%d')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        else:
            st.info("Confirm export authorization to enable the selected-machine workbook download.")

        machine_type_summary = dfm.groupby("CCR TYPE").agg(Avg_Cost=(cost_col, "mean"), Avg_SMU=("SMU AT REBUILD", "mean"), Count=(cost_col, "count")).reset_index()
        machine_type_summary = add_vs_cmr(machine_type_summary)
        machine_type_summary["Sample Confidence"] = machine_type_summary["Count"].apply(sample_confidence)
        st.write("### Average by Rebuild Type")
        display_table(machine_type_summary, currency_cols=["Avg_Cost"], percent_cols=["Vs CMR %"], smu_cols=["Avg_SMU"], number_cols=["Count"])
        st.write("### Bar Chart: Average Rebuild Cost by Type")
        cat_bar_chart(machine_type_summary.rename(columns={"Avg_Cost": "Average Cost"}), "CCR TYPE", "Average Cost", "CCR TYPE", tooltip=["CCR TYPE", "Average Cost"])

        machine_adjusted_summary = machine_adjusted_valid.groupby("CCR TYPE").agg(Adjusted_Avg_Cost=(cost_col, "mean"), Count=(cost_col, "count")).reset_index()
        machine_adjusted_summary["CCR TYPE"] = machine_adjusted_summary["CCR TYPE"].replace({"CPT+H": "CPT+H Adjusted"})
        machine_adjusted_summary["Sample Confidence"] = machine_adjusted_summary["Count"].apply(sample_confidence)
        if machine_show_adjusted_cpth:
            st.write("### Adjusted Average Rebuild Cost by Type")
            display_table(machine_adjusted_summary, currency_cols=["Adjusted_Avg_Cost"], number_cols=["Count"])
            if not machine_adjusted_summary.empty:
                cat_bar_chart(machine_adjusted_summary.rename(columns={"Adjusted_Avg_Cost": "Adjusted Average Cost"}), "CCR TYPE", "Adjusted Average Cost", "CCR TYPE", tooltip=["CCR TYPE", "Adjusted Average Cost"])
            else:
                st.info("No adjusted rebuild cost data available for this machine.")
        else:
            st.info("Adjusted CPT+H comparison is hidden for this machine because both CMR and CPT+H records are required.")

        st.write("### Machine SMU vs USD Analysis Cost")
        machine_chart = dfm[["SMU AT REBUILD", cost_col, "CCR TYPE"]].dropna().rename(columns={"SMU AT REBUILD": "SMU", cost_col: "Cost"})
        cat_scatter_chart(machine_chart, "SMU", "Cost", "CCR TYPE", tooltip=["CCR TYPE", "SMU", "Cost"])

        st.write("### Cross-Type Exception Review")
        benchmark_value = dfm_all["CMR Benchmark Cost"].dropna().iloc[0] if not dfm_all["CMR Benchmark Cost"].dropna().empty else np.nan
        machine_cross = dfm[dfm["Cross-Type Exception Flag"] != ""]
        st.dataframe(pd.DataFrame({"Metric": ["CMR Benchmark Cost", "CPT+H Cross-Type Flags", "Machine Outliers"], "Value": [money(benchmark_value) if not pd.isna(benchmark_value) else "N/A", f"{len(machine_cross):,}", f"{int(dfm_all['Outlier'].sum()):,}"]}), use_container_width=True)
        if not machine_cross.empty:
            display_table(machine_cross[["DEALER", "Region", "CCR TYPE", "SMU AT REBUILD", cost_col, "CMR Benchmark Cost", "Cross-Type Exception Flag"]], currency_cols=[cost_col, "CMR Benchmark Cost"], smu_cols=["SMU AT REBUILD"])
        else:
            st.write("No CPT+H cross-type exceptions for this machine.")

        st.write("### Outlier Performance")
        outlier_perf = dfm_all.groupby("CCR TYPE").agg(Total_Rows=(cost_col, "count"), Outlier_Rows=("Outlier", "sum"), Avg_Cost_All=(cost_col, "mean")).reset_index()
        valid_counts = dfm.groupby("CCR TYPE").size().reset_index(name="Valid_Rows")
        outlier_perf = outlier_perf.merge(valid_counts, on="CCR TYPE", how="left").fillna({"Valid_Rows": 0})
        outlier_perf["Outlier Rate %"] = (outlier_perf["Outlier_Rows"] / outlier_perf["Total_Rows"]) * 100
        display_table(outlier_perf, currency_cols=["Avg_Cost_All"], percent_cols=["Outlier Rate %"], number_cols=["Total_Rows", "Outlier_Rows", "Valid_Rows"])

    with tab3:
        st.subheader("Dealer Analysis")
        dealer_machine_options = ["All Machines"] + sorted(valid["SALES MODEL"].dropna().unique().tolist())
        dealer_machine = st.selectbox("Focus Dealer Analysis on Machine", dealer_machine_options, key="dealer_machine_selector")
        dealer_df = valid.copy() if dealer_machine == "All Machines" else valid[valid["SALES MODEL"] == dealer_machine]
        dealer_full_df = df.copy() if dealer_machine == "All Machines" else df[df["SALES MODEL"] == dealer_machine]
        dealer_summary = dealer_performance_for_df(dealer_df, dealer_full_df, cost_col)
        st.write(f"### Dealer Summary - {dealer_machine}")
        display_table(dealer_summary, currency_cols=["Avg_Cost"], percent_cols=["Vs Section Avg %", "Outlier Rate %", "DQ Rate %", "Cross Flag Rate %"], smu_cols=["Avg_SMU"], number_cols=["Count", "Cross_Type_Flags", "Total_Rows", "Outlier_Rows", "DQ_Rows", "Performance Score"])
        dealer_type = dealer_df.groupby(["DEALER", "CCR TYPE"]).agg(Avg_Cost=(cost_col, "mean"), Avg_SMU=("SMU AT REBUILD", "mean"), Count=(cost_col, "count")).reset_index().sort_values(["DEALER", "CCR TYPE"])
        st.write("### Dealer by Rebuild Type")
        display_table(dealer_type, currency_cols=["Avg_Cost"], smu_cols=["Avg_SMU"], number_cols=["Count"])
        st.write("### Dealer Average Cost Chart")
        if not dealer_summary.empty:
            cat_bar_chart(dealer_summary.rename(columns={"Avg_Cost": "Average Cost"}), "DEALER", "Average Cost", None, tooltip=["DEALER", "Average Cost"])
        else:
            st.info("No dealer summary data available.")

    with tab4:
        st.subheader("Region Analysis")
        region_machine_options = ["All Machines"] + sorted(valid["SALES MODEL"].dropna().unique().tolist())
        region_machine = st.selectbox("Focus Region Analysis on Machine", region_machine_options, key="region_machine_selector")
        region_df = valid.copy() if region_machine == "All Machines" else valid[valid["SALES MODEL"] == region_machine]
        region_summary = region_performance_for_df(region_df, cost_col)
        st.write(f"### Region Summary - {region_machine}")
        display_table(region_summary, currency_cols=["Avg_Cost"], percent_cols=["Vs Section Avg %"], smu_cols=["Avg_SMU"], number_cols=["Count", "Cross_Type_Flags"])
        region_type = region_df.groupby(["Region", "CCR TYPE"]).agg(Avg_Cost=(cost_col, "mean"), Avg_SMU=("SMU AT REBUILD", "mean"), Count=(cost_col, "count")).reset_index().sort_values(["Region", "CCR TYPE"])
        st.write("### Region by Rebuild Type")
        display_table(region_type, currency_cols=["Avg_Cost"], smu_cols=["Avg_SMU"], number_cols=["Count"])
        st.write("### Region Average Cost Chart")
        if not region_summary.empty:
            cat_bar_chart(region_summary.rename(columns={"Avg_Cost": "Average Cost"}), "Region", "Average Cost", None, tooltip=["Region", "Average Cost"])
        else:
            st.info("No region summary data available.")

    with tab5:
        st.subheader("Exception Summary")
        exception_summary = pd.DataFrame({"Metric": ["Cost Outliers", "CPT+H Cross-Type Flags", "SMU 0 or 1", "Insufficient Sample Rows", "Rows Using Global Yearly Fallback Rate", "Rows Using Overall Average Fallback Rate"], "Value": [outlier_count, cross_count, dq_count, insufficient_count, global_year_fallback_count, overall_fallback_count]})
        st.dataframe(exception_summary, use_container_width=True)
        st.write("### Data Quality Summary")
        st.dataframe(data_quality_summary, use_container_width=True)
        st.write("### Data Quality Score")
        display_table(data_quality_score_summary)
        st.write("### Performance Safeguards")
        display_table(performance_safeguards)
        st.write("### Dealer Rate Coverage Score")
        display_table(dealer_rate_coverage_summary)
        st.write("### Dealer Rate Upload Validation")
        if not dealer_rate_validation.empty:
            display_table(dealer_rate_validation, number_cols=["Value"])
        else:
            st.write("No dealer rate validation table available.")
        st.write("### Dealer Rate Exception Rows")
        if not dealer_rate_exception_rows.empty:
            display_table(dealer_rate_exception_rows, currency_cols=["Rate", "Rate USD", "Labor Cost USD", "Adjusted Total Cost USD", "Inflation-Adjusted Adjusted Total Cost USD"], smu_cols=["SMU AT REBUILD"], year_cols=["Service Year"])
        else:
            st.write("No dealer rate exception rows detected.")
        st.write("### Uploaded Dealer Rates Used")
        if not rate_df.empty:
            display_table(rate_df.head(200), currency_cols=["Rate"], number_cols=["Service Year"])
        else:
            st.write("No uploaded/custom dealer rate table available.")
        st.write("### Outlier Performance by Machine + Rebuild Type")
        outlier_perf_all = df.groupby(["SALES MODEL", "CCR TYPE"]).agg(Total_Rows=(cost_col, "count"), Outlier_Rows=("Outlier", "sum"), Avg_Cost_All=(cost_col, "mean")).reset_index()
        outlier_perf_all["Outlier Rate %"] = (outlier_perf_all["Outlier_Rows"] / outlier_perf_all["Total_Rows"]) * 100
        display_table(outlier_perf_all, currency_cols=["Avg_Cost_All"], percent_cols=["Outlier Rate %"], number_cols=["Total_Rows", "Outlier_Rows"])
        st.write("### Cross-Type Exception Rows")
        cross_rows = valid[valid["Cross-Type Exception Flag"] != ""]
        if not cross_rows.empty:
            display_table(cross_rows[["SALES MODEL", "DEALER", "Region", "CCR TYPE", "SMU AT REBUILD", cost_col, "CMR Benchmark Cost", "Cross-Type Exception Flag"]], currency_cols=[cost_col, "CMR Benchmark Cost"], smu_cols=["SMU AT REBUILD"])
        else:
            st.write("No cross-type exceptions detected.")
        st.write("### Outlier Rows")
        outlier_rows = df[df["Outlier"] == True]
        if not outlier_rows.empty:
            display_table(outlier_rows, currency_cols=["PARTS DN", "PARTS DN USD", "Rate", "Rate USD", "Labor Cost USD", "Adjusted Total Cost USD", "Inflation-Adjusted Adjusted Total Cost USD"], smu_cols=["SMU AT REBUILD"], year_cols=["Service Year"])
        else:
            st.write("No cost outliers detected.")

    with tab6:
        st.subheader("Executive Insights")
        st.write("### Executive Summary Narrative")
        for line in executive_narrative:
            st.write("•", line)
        st.write("### Supporting Insights")
        insights = []
        if not summary.empty:
            top_type = summary.sort_values("Avg_Cost", ascending=False).iloc[0]
            insights.append(f"Highest average rebuild type cost is {top_type['CCR TYPE']} at {money(top_type['Avg_Cost'])}.")
        machine_avg = valid.groupby("SALES MODEL")[cost_col].mean().sort_values(ascending=False)
        if len(machine_avg) > 0:
            insights.append(f"Highest average machine cost is {machine_avg.index[0]} at {money(machine_avg.iloc[0])}.")
        try:
            review_dealers = dealer_summary[dealer_summary["Performance Label"] == "Review Needed"] if "Performance Label" in dealer_summary.columns else pd.DataFrame()
            if not review_dealers.empty:
                insights.append(f"{len(review_dealers)} dealer(s) are marked Review Needed based on score, cost position, outlier rate, cross-type flags, and data quality.")
        except Exception:
            pass
        insights.append(f"Cost outliers excluded from core averages: {outlier_count:,}.")
        insights.append(f"CPT+H cross-type review flags: {cross_count:,}.")
        insights.append(f"Rows using fallback labor rates: {global_year_fallback_count + overall_fallback_count:,}.")
        for insight in insights:
            st.write("•", insight)

    with tab7:
        st.subheader("How to Use This App")
        st.markdown("""
        1. Upload the rebuild workbook.
        2. Select the dealer labor-rate source. Built-in expanded 2016–2026 rates are selected by default.
        3. Review validation and readiness checks.
        4. Run the analysis.
        5. Review the Executive Dashboard, dealer rate coverage score, and data quality score.
        6. Use Machine Detail for selected-machine review.
        7. Review Exceptions & Data Quality before sharing results.
        8. If needed, manually change advanced thresholds; defaults preserve the standard methodology.
        9. Use Strict Mode only for formal review where the configured thresholds should stop weak runs.
        10. Choose the least-detailed export mode and provide an export reason. Use Power BI Dataset Export when building Power BI reports.
        11. Interpret Excel highlights: red = cost outlier, orange = cross-type exception, yellow = insufficient sample group.
        """)

    with tab8:
        st.subheader("Methodology")
        st.markdown(METHOD_LOCK_TEXT)
        st.markdown("""**Visual style:** Charts, checkboxes, filter controls, tabs, and workbook headers use a Caterpillar-inspired black, yellow, and gray palette.  
**Important:** Official Caterpillar logo usage should follow internal brand/asset approval rules.  
**V16.2.3 fix:** Removes the confidentiality label from the app UI, export acknowledgements, workbook watermark rows, workbook metadata, machine summaries, and Power BI dataset exports. Power BI Dataset Export keeps column headers directly on row 1 for every sheet.""")

    with tab9:
        st.subheader("Governance & Data Dictionary")
        st.write("### Known Limitations")
        display_table(known_limitations)
        st.write("### Data Dictionary")
        display_table(data_dictionary)
        st.write("### Role-Based Export Design")
        display_table(role_policy)
        st.write("### Saved Parameter Summary")
        display_table(parameter_summary)
        st.write("### Performance Safeguards")
        display_table(performance_safeguards)

    with tab10:
        st.subheader("Configured Rebuild Types and Regions")
        st.write("### Certified Rebuild Types")
        st.dataframe(rebuild_reference, use_container_width=True)
        st.write("### Configured Regions")
        st.dataframe(region_reference, use_container_width=True)
        st.write("### Observed Rebuild Types in Current Run")
        st.dataframe(pd.DataFrame({"Observed CCR TYPE": sorted(valid["CCR TYPE"].dropna().unique())}), use_container_width=True)
        st.write("### Observed Regions in Current Run")
        st.dataframe(pd.DataFrame({"Observed Region": sorted(valid["Region"].dropna().unique())}), use_container_width=True)

    output = BytesIO()
    export_metadata = pd.concat([
        metadata.copy(),
        pd.DataFrame({"Field": ["Export Mode", "Export Reason", "Scenario Name", "User Role View", "Strict Mode", "Methodology Version", "Outlier Rule Version", "Dealer Rate Version", "Security Control Version", "Export Format Version"], "Value": [export_mode, export_reason_final, scenario_name if scenario_name else "Not provided", user_role_view, "Yes" if strict_mode else "No", METHODOLOGY_VERSION, OUTLIER_RULE_VERSION, DEALER_RATE_VERSION, SECURITY_CONTROL_VERSION, EXPORT_FORMAT_VERSION]}),
    ], ignore_index=True)
    cover_sheet = build_cover_sheet(
        export_metadata,
        run_readiness_summary,
        data_quality_score_summary,
        dealer_rate_coverage_summary,
        export_mode,
        export_reason_final,
        scenario_name_value=scenario_name,
        user_role_view_value=user_role_view,
        strict_mode_value=strict_mode,
    )
    if export_mode == "Power BI Dataset Export":
        # Dedicated clean Power BI writer. Do not write cover page, watermark rows, merged cells,
        # decorative formatting, or normal review-workbook tabs before the dataset tables.
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            write_powerbi_dataset_workbook(writer, analysis, scenario_name, export_reason_final, export_mode)
    else:
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            cover_sheet.to_excel(writer, sheet_name="Cover Page", index=False)
            export_metadata.to_excel(writer, sheet_name="Run Metadata", index=False)
            run_readiness_summary.to_excel(writer, sheet_name="Run Readiness", index=False)
            dealer_rate_coverage_summary.to_excel(writer, sheet_name="Rate Coverage", index=False)
            data_quality_score_summary.to_excel(writer, sheet_name="Data Quality Score", index=False)
            parameter_summary.to_excel(writer, sheet_name="Parameters", index=False)
            performance_safeguards.to_excel(writer, sheet_name="Performance Checks", index=False)
            if export_mode == "Summary Only":
                summary.to_excel(writer, sheet_name="Summary", index=False)
                if show_adjusted_cpth:
                    adjusted_summary.to_excel(writer, sheet_name="Adjusted Summary", index=False)
                else:
                    pd.DataFrame({"Message": ["Adjusted CPT+H comparison hidden because both CMR and CPT+H are required."]}).to_excel(writer, sheet_name="Adjusted CPT-H Hidden", index=False)
                machine_benchmark_ranking.to_excel(writer, sheet_name="Machine Ranking", index=False)
                pd.DataFrame({"Executive Narrative": executive_narrative}).to_excel(writer, sheet_name="Executive Narrative", index=False)
            elif export_mode == "Exceptions Only":
                exception_summary.to_excel(writer, sheet_name="Exceptions", index=False)
                data_quality_summary.to_excel(writer, sheet_name="Data Quality", index=False)
                dealer_rate_exception_rows.to_excel(writer, sheet_name="Rate Exceptions", index=False)
                outlier_perf_all.to_excel(writer, sheet_name="Outlier Performance", index=False)
                cross_rows.to_excel(writer, sheet_name="Cross Type Flags", index=False)
                df[df["Outlier"] == True].to_excel(writer, sheet_name="Outlier Rows", index=False)
            elif export_mode == "Dealer Rate Audit":
                dealer_rate_validation.to_excel(writer, sheet_name="Rate Validation", index=False)
                dealer_rate_coverage_summary.to_excel(writer, sheet_name="Rate Coverage Detail", index=False)
                dealer_rate_exception_rows.to_excel(writer, sheet_name="Rate Exceptions", index=False)
                rate_df.to_excel(writer, sheet_name="Dealer Rates Used", index=False)
                data_quality_summary.to_excel(writer, sheet_name="Data Quality", index=False)
            else:
                summary.to_excel(writer, sheet_name="Summary", index=False)
                if show_adjusted_cpth:
                    adjusted_summary.to_excel(writer, sheet_name="Adjusted Summary", index=False)
                else:
                    pd.DataFrame({"Message": ["Adjusted CPT+H comparison hidden because both CMR and CPT+H are required."]}).to_excel(writer, sheet_name="Adjusted CPT-H Hidden", index=False)
                machine_benchmark_ranking.to_excel(writer, sheet_name="Machine Ranking", index=False)
                pd.DataFrame({"Executive Narrative": executive_narrative}).to_excel(writer, sheet_name="Executive Narrative", index=False)
                exception_summary.to_excel(writer, sheet_name="Exceptions", index=False)
                data_quality_summary.to_excel(writer, sheet_name="Data Quality", index=False)
                dealer_rate_validation.to_excel(writer, sheet_name="Rate Validation", index=False)
                dealer_rate_exception_rows.to_excel(writer, sheet_name="Rate Exceptions", index=False)
                rate_df.to_excel(writer, sheet_name="Dealer Rates Used", index=False)
                outlier_perf_all.to_excel(writer, sheet_name="Outlier Performance", index=False)
                cross_rows.to_excel(writer, sheet_name="Cross Type Flags", index=False)
                df[df["Outlier"] == True].to_excel(writer, sheet_name="Outlier Rows", index=False)
                pd.DataFrame({"Year": list(cpi_table.keys()), "CPI": list(cpi_table.values())}).sort_values("Year").to_excel(writer, sheet_name="BLS CPI", index=False)
                fx_lookup.to_excel(writer, sheet_name="FX Rates", index=False)
                rebuild_reference.to_excel(writer, sheet_name="Rebuild Type Reference", index=False)
                region_reference.to_excel(writer, sheet_name="Region Reference", index=False)
                known_limitations.to_excel(writer, sheet_name="Known Limitations", index=False)
                data_dictionary.to_excel(writer, sheet_name="Data Dictionary", index=False)
                role_policy.to_excel(writer, sheet_name="Role Design", index=False)
                for machine in valid["SALES MODEL"].dropna().unique():
                    valid[valid["SALES MODEL"] == machine].to_excel(writer, sheet_name=safe_sheet_name(machine), index=False)
            apply_excel_brand_formatting(writer.book)
            apply_confidential_watermark(writer.book, scenario_name)
    if render_export_acknowledgement("full_export_ack"):
        safe_scenario = re.sub(r"[^A-Za-z0-9_-]+", "_", scenario_name.strip()) if scenario_name else "Scenario"
        download_label = "Download Power BI Dataset Workbook" if export_mode == "Power BI Dataset Export" else "Download Workbook"
        export_prefix = "Rebuild_Analysis_PowerBI_Dataset" if export_mode == "Power BI Dataset Export" else "Rebuild_Analysis"
        st.download_button(download_label, data=output.getvalue(), file_name=f"{export_prefix}_{safe_scenario}_{export_mode.replace(' ', '_')}_{datetime.now().strftime('%Y-%m-%d')}.xlsx")
    else:
        st.info("Confirm export authorization to enable the full workbook download.")
    st.markdown("""<div class="cat-footer"><strong>Rebuild Analytics Platform</strong> | Internal analytics tool | Cost, inflation, dealer, region, outlier, cross-type exception, governance, performance-safeguard, and Power BI dataset reporting</div>""", unsafe_allow_html=True)
else:
    st.info("Upload file and run analysis")

