
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


APP_VERSION = "V15.3"
APP_LAST_UPDATED = "2026-06-26"

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
        :root {
            --cat-yellow:#FFC500;
            --cat-black:#000000;
            --cat-charcoal:#1F1F1F;
            --cat-gray:#7A7A7A;
            --cat-light:#F4F4F4;
            --cat-border:#D9D9D9;
        }

        .stApp {
            background: var(--cat-light);
            color: var(--cat-charcoal);
        }

        .block-container {
            padding-top: 1rem;
            padding-left: 2rem;
            padding-right: 2rem;
            max-width: 1550px;
        }

        .cat-header {
            background: linear-gradient(90deg,#000000 0%,#1F1F1F 72%,#FFC500 72%,#FFC500 100%);
            padding: 22px 28px;
            border-radius: 4px;
            margin-bottom: 20px;
            box-shadow: 0 4px 14px rgba(0,0,0,.22);
            border-left: 8px solid #FFC500;
            display: flex;
            justify-content: space-between;
            align-items: center;
            gap: 24px;
        }

        .cat-header-title {
            color: #FFFFFF;
            font-size: 34px;
            font-weight: 900;
            letter-spacing: .6px;
            line-height: 1.05;
            margin: 0;
        }

        .cat-header-subtitle {
            color: #D9D9D9;
            font-size: 14px;
            margin-top: 9px;
            font-weight: 600;
        }

        .cat-version-pill {
            display: inline-block;
            background: #FFC500;
            color: #000000;
            font-weight: 900;
            padding: 5px 12px;
            border-radius: 2px;
            margin-top: 12px;
            font-size: 12px;
        }

        .cat-logo-box {
            background:#FFC500;
            color:#000000;
            font-weight: 1000;
            font-size: 34px;
            letter-spacing: -1px;
            padding: 12px 18px;
            border: 3px solid #000000;
            min-width: 95px;
            text-align:center;
        }

        .cat-logo-img {
            max-height: 72px;
            max-width: 210px;
            background:#FFFFFF;
            padding:6px;
            border-radius:3px;
            border:2px solid #000000;
        }

        .method-card {
            background:#FFFFFF;
            border-left:8px solid #FFC500;
            padding:15px 18px;
            margin:12px 0 18px 0;
            box-shadow:0 2px 8px rgba(0,0,0,.08);
            border-radius:4px;
        }

        .method-title {
            font-size:18px;
            font-weight:900;
            color:#000000;
            margin-bottom:5px;
        }

        .method-body {
            font-size:14px;
            color:#333333;
            line-height:1.45;
        }

        h1, h2, h3 {
            color:#1F1F1F;
            font-weight:850;
        }

        h2 {
            border-left:6px solid #FFC500;
            padding-left:10px;
        }

        [data-testid="stMetric"] {
            background:#FFFFFF;
            border:1px solid #D9D9D9;
            border-top:5px solid #FFC500;
            border-radius:4px;
            padding:15px 17px;
            box-shadow:0 2px 8px rgba(0,0,0,.08);
        }

        [data-testid="stMetricLabel"] {
            color:#4A4A4A;
            font-weight:800;
        }

        [data-testid="stMetricValue"] {
            color:#000000;
            font-weight:950;
        }

        button[data-baseweb="tab"] {
            background:#FFFFFF;
            color:#1F1F1F;
            border-radius:2px 2px 0 0;
            font-weight:800;
            border:1px solid #D9D9D9;
            margin-right:4px;
            padding:10px 14px;
        }

        button[data-baseweb="tab"][aria-selected="true"] {
            background:#000000;
            color:#FFC500;
            border-bottom:4px solid #FFC500;
        }

        .stButton > button {
            background:#FFC500;
            color:#000000;
            border:2px solid #000000;
            border-radius:3px;
            font-weight:900;
        }

        .stButton > button:hover {
            background:#000000;
            color:#FFC500;
            border:2px solid #FFC500;
        }

        .stDownloadButton > button {
            background:#000000;
            color:#FFC500;
            border:2px solid #FFC500;
            border-radius:3px;
            font-weight:900;
        }

        .stDownloadButton > button:hover {
            background:#FFC500;
            color:#000000;
            border:2px solid #000000;
        }

        label {
            font-weight:800 !important;
            color:#1F1F1F !important;
        }

        div[data-testid="stAlert"] {
            border-radius:4px;
            border-left:6px solid #FFC500;
        }

        [data-testid="stDataFrame"] {
            background:#FFFFFF;
            border:1px solid #D9D9D9;
            border-radius:4px;
            box-shadow:0 2px 8px rgba(0,0,0,.06);
        }

        details {
            background:#FFFFFF !important;
            border:1px solid #D9D9D9 !important;
            border-radius:4px !important;
        }

        summary {
            font-weight:900 !important;
            color:#111111 !important;
        }

        [data-testid="stVegaLiteChart"] {
            background:#FFFFFF;
            border:1px solid #D9D9D9;
            border-radius:4px;
            padding:10px;
            box-shadow:0 2px 8px rgba(0,0,0,.06);
        }

        div[data-baseweb="select"] > div {
            border:2px solid #1F1F1F !important;
            border-radius:3px !important;
            background:#FFFFFF !important;
        }

        div[data-baseweb="select"]:focus-within > div {
            border-color:#FFC500 !important;
            box-shadow:0 0 0 2px rgba(255,197,0,.35) !important;
        }

        span[data-baseweb="tag"], div[data-baseweb="tag"] {
            background:#FFC500 !important;
            color:#000000 !important;
            border:1px solid #000000 !important;
            font-weight:800 !important;
        }

        span[data-baseweb="tag"] span, div[data-baseweb="tag"] span {
            color:#000000 !important;
            font-weight:800 !important;
        }

        div[role="option"][aria-selected="true"],
        ul[role="listbox"] li[aria-selected="true"] {
            background:#FFC500 !important;
            color:#000000 !important;
            font-weight:900 !important;
        }

        div[role="option"]:hover,
        ul[role="listbox"] li:hover {
            background:#1F1F1F !important;
            color:#FFC500 !important;
        }

        div[data-testid="stCheckbox"] [data-baseweb="checkbox"] > div:first-child {
            border-color:#000000 !important;
        }

        div[data-testid="stCheckbox"] input:checked + div {
            background-color:#FFC500 !important;
            border-color:#000000 !important;
            color:#000000 !important;
        }

        .cat-footer {
            margin-top:35px;
            padding:15px;
            background:#1F1F1F;
            color:#D9D9D9;
            font-size:12px;
            border-top:5px solid #FFC500;
            border-radius:3px;
        }

        .cat-footer strong {
            color:#FFC500;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )


def render_header():
    logo = find_logo_file()
    if logo:
        b64 = img_to_base64(logo)
        if b64:
            logo_html = f'<img class="cat-logo-img" src="data:image/png;base64,{b64}" />'
        else:
            logo_html = '<div class="cat-logo-box">CAT</div>'
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

VALID_REGIONS = [
    "AFRICA",
    "ANZP/INDONESIA",
    "CHINA",
    "EASTERN US",
    "EUROPE",
    "INDIA",
    "JAPAN/ASIA",
    "LATIN AMERICA",
    "MIDDLE EAST & EURASIA",
    "WESTERN US & CANADA",
]

REBUILD_TYPE_ALIASES = {
    "CERTIFIED MACHINE REBUILD": "CMR",
    "CMR - CERTIFIED MACHINE REBUILD": "CMR",
    "CERTIFIED MACHINE REBUILD UPGRADE": "CMR-U",
    "CERTIFIED POWERTRAIN + HYDRAULICS": "CPT+H",
    "CPT PLUS H": "CPT+H",
    "CERTIFIED POWERTRAIN": "CPT-O",
    "CPT 777": "CPT-O",
}

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
"""

if "run_clicked" not in st.session_state:
    st.session_state.run_clicked = False
if "analysis" not in st.session_state:
    st.session_state.analysis = None


# =====================================================
# CHART HELPERS
# =====================================================
CAT_DOMAIN = list(CERTIFIED_REBUILD_TYPES.keys()) + ["CPT+H Adjusted"]
CAT_RANGE = [
    "#000000", "#FFC500", "#7A7A7A", "#4D4D4D", "#FFCD00",
    "#2B2B2B", "#A6A6A6", "#6B5B00", "#C49700", "#595959",
    "#B38F00", "#333333", "#D9A300", "#808080", "#F2B705",
    "#1F1F1F", "#C0C0C0", "#8A6F00", "#E0B000", "#666666",
    "#FFC500",
]
CAT_COLOR_SCALE = alt.Scale(domain=CAT_DOMAIN, range=CAT_RANGE)


def cat_scatter_chart(data, x, y, color="CCR TYPE", tooltip=None, height=420):
    if data.empty:
        st.info("No chart data available.")
        return

    tooltip = tooltip or [x, y, color]
    chart = (
        alt.Chart(data)
        .mark_circle(size=78, opacity=0.82, stroke="#000000", strokeWidth=0.35)
        .encode(
            x=alt.X(
                f"{x}:Q",
                title=x,
                axis=alt.Axis(labelColor="#1F1F1F", titleColor="#1F1F1F", gridColor="#E6E6E6"),
            ),
            y=alt.Y(
                f"{y}:Q",
                title=y,
                axis=alt.Axis(format="$,.0f", labelColor="#1F1F1F", titleColor="#1F1F1F", gridColor="#E6E6E6"),
            ),
            color=alt.Color(f"{color}:N", scale=CAT_COLOR_SCALE, legend=alt.Legend(title=color, orient="right")),
            tooltip=tooltip,
        )
        .properties(height=height)
        .configure_view(stroke="#D9D9D9")
        .configure_axis(labelFontSize=12, titleFontSize=13)
        .configure_legend(labelFontSize=12, titleFontSize=13)
    )
    st.altair_chart(chart, use_container_width=True)


def cat_bar_chart(data, x, y, color=None, tooltip=None, height=360):
    if data.empty:
        st.info("No chart data available.")
        return

    tooltip = tooltip or [x, y]
    if color is None:
        color_encoding = alt.value("#FFC500")
    else:
        color_encoding = alt.Color(f"{color}:N", scale=CAT_COLOR_SCALE, legend=alt.Legend(title=color))

    chart = (
        alt.Chart(data)
        .mark_bar(cornerRadiusTopLeft=2, cornerRadiusTopRight=2, stroke="#000000", strokeWidth=0.45)
        .encode(
            x=alt.X(
                f"{x}:N",
                title=x,
                sort="-y",
                axis=alt.Axis(labelAngle=-35, labelColor="#1F1F1F", titleColor="#1F1F1F"),
            ),
            y=alt.Y(
                f"{y}:Q",
                title=y,
                axis=alt.Axis(format="$,.0f", labelColor="#1F1F1F", titleColor="#1F1F1F", gridColor="#E6E6E6"),
            ),
            color=color_encoding,
            tooltip=tooltip,
        )
        .properties(height=height)
        .configure_view(stroke="#D9D9D9")
        .configure_axis(labelFontSize=12, titleFontSize=13)
        .configure_legend(labelFontSize=12, titleFontSize=13)
    )
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

    aliases = {
        "WESTERN US AND CANADA": "WESTERN US & CANADA",
        "WESTERN US & CANADA": "WESTERN US & CANADA",
        "EASTERN US": "EASTERN US",
        "LATAM": "LATIN AMERICA",
        "MIDDLE EAST AND EURASIA": "MIDDLE EAST & EURASIA",
        "MIDDLE EAST & EURASIA": "MIDDLE EAST & EURASIA",
        "JAPAN ASIA": "JAPAN/ASIA",
        "JAPAN/ASIA": "JAPAN/ASIA",
        "ANZP INDONESIA": "ANZP/INDONESIA",
        "ANZP/INDONESIA": "ANZP/INDONESIA",
    }

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
# FX / CPI / RATE HELPERS
# =====================================================
COMMON_CURRENCY_FALLBACK_TO_USD = {
    "USD": 1.00,
    "CAD": 0.74,
    "EUR": 1.09,
    "GBP": 1.27,
    "AUD": 0.66,
    "NZD": 0.61,
    "MXN": 0.058,
    "BRL": 0.19,
    "CLP": 0.0011,
    "COP": 0.00025,
    "PEN": 0.27,
    "JPY": 0.0065,
    "CNY": 0.14,
    "INR": 0.012,
    "ZAR": 0.055,
}


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
    fallback = {
        2010: 218.056,
        2011: 224.939,
        2012: 229.594,
        2013: 232.957,
        2014: 236.736,
        2015: 237.017,
        2016: 240.007,
        2017: 245.120,
        2018: 251.107,
        2019: 255.657,
        2020: 258.811,
        2021: 270.970,
        2022: 292.655,
        2023: 304.702,
        2024: 313.689,
        2025: 322.000,
        2026: 330.080,
        2027: 339.982,
        2028: 350.181,
        2029: 360.686,
        2030: 371.506,
    }
    return {y: fallback[y] for y in range(int(start_year), int(end_year) + 1) if y in fallback}


def fetch_bls_cpi_annual(start_year, end_year):
    url = "https://api.bls.gov/publicAPI/v2/timeseries/data/"
    payload = {"seriesid": ["CUUR0000SA0"], "startyear": str(int(start_year)), "endyear": str(int(end_year))}
    request = urllib.request.Request(
        url,
        data=json.dumps(payload).encode("utf-8"),
        headers={"Content-Type": "application/json"},
        method="POST",
    )

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
    years = list(range(start, end + 1))
    return pd.DataFrame(
        {
            "Dealer Code": ["DEFAULT"] * len(years),
            "Service Year": years,
            "Rate": [115 + (year - 2016) * 3 for year in years],
        }
    )


def build_rate_table_from_workbook(rate_file):
    rates = pd.read_excel(rate_file, sheet_name=None)
    rows = []

    for sheet_name, sheet in rates.items():
        if str(sheet_name).strip().lower() == "summary":
            continue

        temp = sheet.copy()
        temp.columns = [str(col).strip() for col in temp.columns]
        rename_map = {}

        for col in temp.columns:
            col_upper = col.upper()
            if col_upper == "YEAR":
                rename_map[col] = "Service Year"
            elif "BASE RATE" in col_upper or col_upper == "RATE":
                rename_map[col] = "Rate"

        temp = temp.rename(columns=rename_map)
        dealer_code = str(sheet_name).split("_")[0].strip()

        if "Dealer Code" not in temp.columns:
            temp["Dealer Code"] = dealer_code

        if "Service Year" in temp.columns and "Rate" in temp.columns:
            rows.append(temp[["Dealer Code", "Service Year", "Rate"]])

    if not rows:
        return build_default_rate_table()

    rate_df = pd.concat(rows, ignore_index=True)
    rate_df["Service Year"] = pd.to_numeric(rate_df["Service Year"], errors="coerce")
    rate_df["Rate"] = pd.to_numeric(rate_df["Rate"], errors="coerce")
    rate_df = rate_df.dropna(subset=["Service Year", "Rate"])
    rate_df["Service Year"] = rate_df["Service Year"].astype(int)
    rate_df["Dealer Code"] = rate_df["Dealer Code"].astype(str).str.strip()

    return rate_df


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
        rows.append(
            {
                "Sheet": sheet_name,
                "Rows": len(sheet),
                "Detected CCR TYPE": detected_type,
                "Missing Required Columns": ", ".join(missing) if missing else "None",
            }
        )

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
        for rebuild_type in [rebuild_type for rebuild_type in observed_types if rebuild_type not in CERTIFIED_REBUILD_TYPES]:
            issues.append({"Issue Type": "Unknown Rebuild Type", "Sheet": "All", "Details": rebuild_type})

    issue_df = pd.DataFrame(issues) if issues else pd.DataFrame(columns=["Issue Type", "Sheet", "Details"])
    return profile, issue_df


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

with st.expander("Methodology Lock", expanded=False):
    st.markdown(METHOD_LOCK_TEXT)

st.subheader("Currency Conversion")
auto_currency = st.checkbox("Automatically convert source currency to USD", True)
default_source_currency = st.text_input("Default source currency if no currency column exists", "USD").upper().strip()
dealer_rate_currency_mode = st.selectbox("Dealer labor rate currency", ["USD", "Same as source currency"], index=0)

st.subheader("Dealer Labor Rates")
use_default = st.checkbox("Use Built-in Default Dealer Rates", True)
if use_default:
    st.warning("Default dealer rates are being used. For production analysis, upload the official dealer-by-year rate workbook.")

rate_file = None
if not use_default:
    rate_file = st.file_uploader("Upload Custom Dealer-by-Year Rate Workbook", type=["xlsx"])

apply_inflation = st.checkbox("Apply BLS CPI-U Inflation", True)
base_year = st.number_input("Base Year for Inflation", 2010, 2030, 2026)
machine_input = st.text_input("Machines (optional, comma-separated; leave blank for all)")
start_year = st.number_input("Start Year", 2010, 2030, 2016)
end_year = st.number_input("End Year", 2010, 2030, 2026)
rebuild_filter = st.multiselect("Filter Rebuild Types", options=list(CERTIFIED_REBUILD_TYPES.keys()), default=["CMR", "CPT+H", "CPT-O"])
region_filter = st.multiselect("Filter Regions", options=VALID_REGIONS + ["UNKNOWN", "OTHER"], default=VALID_REGIONS + ["UNKNOWN", "OTHER"])

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

    df["Dealer Code"] = df["DEALER"].astype(str).str.extract(r"([A-Z]\d{3})")
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

    rate_df = build_default_rate_table(2010, 2030) if (use_default or rate_file is None) else build_rate_table_from_workbook(rate_file)
    rate_df["Service Year"] = pd.to_numeric(rate_df["Service Year"], errors="coerce").astype("Int64")
    df = df.merge(rate_df, how="left", on=["Dealer Code", "Service Year"])

    yearly_avg = rate_df.groupby("Service Year")["Rate"].mean()
    overall_avg = rate_df["Rate"].mean()
    df["Rate Source"] = "Dealer-Year Rate"

    missing_dealer_year = df["Rate"].isna()
    df.loc[missing_dealer_year, "Rate"] = df.loc[missing_dealer_year, "Service Year"].map(yearly_avg)
    df.loc[missing_dealer_year, "Rate Source"] = "Global Yearly Fallback Rate"

    missing_overall = df["Rate"].isna()
    df.loc[missing_overall, "Rate"] = overall_avg
    df.loc[missing_overall, "Rate Source"] = "Overall Average Fallback Rate"

    df["Base Rate Year Used"] = df["Service Year"]
    df["Avg Base Rate Used"] = df["Rate"]

    df["PARTS DN"] = pd.to_numeric(df["PARTS DN"], errors="coerce")
    df["REBUILD WORK HRS"] = pd.to_numeric(df["REBUILD WORK HRS"], errors="coerce")
    df["SMU AT REBUILD"] = pd.to_numeric(df["SMU AT REBUILD"], errors="coerce")
    missing_parts_count = int(df["PARTS DN"].isna().sum())
    missing_labor_hours_count = int(df["REBUILD WORK HRS"].isna().sum())
    df = df[(df["PARTS DN"] > 0) & (df["REBUILD WORK HRS"].notna())].copy()

    df["PARTS DN USD"] = df["PARTS DN"] * df["FX to USD"]

    if dealer_rate_currency_mode == "Same as source currency":
        df["Rate FX to USD"] = df["FX to USD"]
    else:
        df["Rate FX to USD"] = 1.0

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
        if len(cmr_valid) >= 3:
            benchmark = cmr_valid[cost_col].median()
            df.loc[group.index, "CMR Benchmark Cost"] = benchmark
            mask = group["CCR TYPE"].eq("CPT+H") & group["Outlier"].eq(False) & (group[cost_col] > benchmark * 1.10)
            df.loc[group.index[mask], "Cross-Type Exception Flag"] = "CPT+H Cost Above Typical CMR"

    valid = df[df["Outlier"] == False].copy()

    summary = valid.groupby("CCR TYPE").agg(
        Avg_Cost=(cost_col, "mean"),
        Avg_SMU=("SMU AT REBUILD", "mean"),
        Count=(cost_col, "count"),
    ).reset_index()
    summary = add_vs_cmr(summary)
    summary["Sample Confidence"] = summary["Count"].apply(sample_confidence)

    adjusted_valid = valid[~((valid["CCR TYPE"] == "CPT+H") & (valid["Cross-Type Exception Flag"] == "CPT+H Cost Above Typical CMR"))].copy()
    adjusted_summary = adjusted_valid.groupby("CCR TYPE").agg(
        Adjusted_Avg_Cost=(cost_col, "mean"),
        Count=(cost_col, "count"),
    ).reset_index()
    adjusted_summary["CCR TYPE"] = adjusted_summary["CCR TYPE"].replace({"CPT+H": "CPT+H Adjusted"})
    adjusted_summary["Sample Confidence"] = adjusted_summary["Count"].apply(sample_confidence)

    rebuild_reference = pd.DataFrame([{"CCR TYPE": key, "Description": value} for key, value in CERTIFIED_REBUILD_TYPES.items()])
    region_reference = pd.DataFrame({"Configured Region": VALID_REGIONS})

    metadata = pd.DataFrame(
        {
            "Field": [
                "App Version",
                "Run Timestamp",
                "Rows Uploaded",
                "Rows After Filters",
                "Valid Rows",
                "Start Year",
                "End Year",
                "Base Year",
                "Machine Filter",
                "Rebuild Type Filter",
                "Region Filter",
                "Default Source Currency",
                "Dealer Rate Currency Mode",
                "Currency Column Detected",
                "BLS CPI Source",
                "FX Source",
                "Analysis Cost Used",
            ],
            "Value": [
                APP_VERSION,
                datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                rows_uploaded,
                len(df),
                len(valid),
                start_year,
                end_year,
                base_year,
                machine_input if machine_input else "All",
                ", ".join(rebuild_filter),
                ", ".join(region_filter),
                default_source_currency,
                dealer_rate_currency_mode,
                currency_col if currency_col else "None",
                cpi_source,
                "Frankfurter annual average; embedded fallback if unavailable",
                cost_col,
            ],
        }
    )

    data_quality_summary = pd.DataFrame(
        {
            "Metric": [
                "Missing Service Date",
                "Missing Dealer Code",
                "Missing Parts DN",
                "Missing Labor Hours",
                "SMU 0 or 1",
                "Unknown/Other Regions",
                "Rows Using Fallback FX",
                "Rows Using Global Yearly Fallback Rate",
                "Rows Using Overall Average Fallback Rate",
            ],
            "Value": [
                missing_service_date_count,
                missing_dealer_code_count,
                missing_parts_count,
                missing_labor_hours_count,
                int(df["Data Quality Exception Flag"].eq("SMU 0 or 1").sum()),
                unknown_region_count,
                fallback_fx_count,
                int((df["Rate Source"] == "Global Yearly Fallback Rate").sum()),
                int((df["Rate Source"] == "Overall Average Fallback Rate").sum()),
            ],
        }
    )

    return {
        "df": df,
        "valid": valid,
        "adjusted_valid": adjusted_valid,
        "summary": summary,
        "adjusted_summary": adjusted_summary,
        "cost_col": cost_col,
        "cpi_table": cpi_table,
        "cpi_source": cpi_source,
        "base_cpi": base_cpi,
        "fx_lookup": fx_lookup,
        "currency_col": currency_col,
        "rebuild_reference": rebuild_reference,
        "region_reference": region_reference,
        "metadata": metadata,
        "data_quality_summary": data_quality_summary,
        "outlier_count": int(df["Outlier"].sum()),
        "cross_count": int((valid["Cross-Type Exception Flag"] != "").sum()),
        "dq_count": int(df["Data Quality Exception Flag"].eq("SMU 0 or 1").sum()),
        "insufficient_count": int(df["Insufficient Sample Group"].sum()),
        "global_year_fallback_count": int((df["Rate Source"] == "Global Yearly Fallback Rate").sum()),
        "overall_fallback_count": int((df["Rate Source"] == "Overall Average Fallback Rate").sum()),
    }


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

    tabs = st.tabs(
        [
            "Executive Dashboard",
            "Machine Detail",
            "Dealer Performance",
            "Region Performance",
            "Exceptions & Data Quality",
            "Executive Insights",
            "Methodology",
            "Reference",
        ]
    )
    tab1, tab2, tab3, tab4, tab5, tab6, tab7, tab8 = tabs

    with tab1:
        st.subheader("Run Summary")
        display_table(metadata)

        st.subheader("Executive Summary")
        col1, col2, col3, col4 = st.columns(4)
        col1.metric("Valid Rows", f"{len(valid):,}")
        col2.metric("Cost Outliers", f"{outlier_count:,}")
        col3.metric("Avg USD Analysis Cost", money(valid[cost_col].mean()))
        col4.metric("Cross-Type Flags", f"{cross_count:,}")

        st.write("### Standard Rebuild Type Average Cost")
        display_table(summary, currency_cols=["Avg_Cost"], percent_cols=["Vs CMR %"], smu_cols=["Avg_SMU"], number_cols=["Count"])

        st.write("### Adjusted Rebuild Type Average Cost")
        display_table(adjusted_summary, currency_cols=["Adjusted_Avg_Cost"], number_cols=["Count"])

        st.write("### SMU vs USD Analysis Cost by Rebuild Type")
        scatter = valid[["SMU AT REBUILD", cost_col, "CCR TYPE"]].dropna().rename(
            columns={"SMU AT REBUILD": "SMU", cost_col: "Cost"}
        )
        cat_scatter_chart(scatter, "SMU", "Cost", "CCR TYPE", tooltip=["CCR TYPE", "SMU", "Cost"])

        st.write("### Currency and Inflation Settings")
        settings = pd.DataFrame(
            {
                "Setting": [
                    "Currency Column Detected",
                    "Default Source Currency",
                    "Dealer Rate Currency Mode",
                    "FX Source",
                    "CPI Source",
                    "BLS Series",
                    "Inflation Applied",
                    "Base Year",
                    "Base Year CPI",
                    "Cost Source",
                    "Labor Cost",
                    "PLUS PARTS DN",
                    "Analysis Cost Used",
                ],
                "Value": [
                    str(currency_col) if currency_col else "None - default source currency used",
                    default_source_currency,
                    dealer_rate_currency_mode,
                    "Frankfurter annual average; embedded fallback if unavailable",
                    cpi_source,
                    "CUUR0000SA0 - CPI-U All Items, U.S. city average, not seasonally adjusted",
                    "Yes" if apply_inflation else "No",
                    str(base_year),
                    f"{base_cpi:,.3f}" if pd.notna(base_cpi) else "N/A",
                    "PARTS DN USD + Labor Cost USD",
                    "Labor Hours × Dealer Service-Year Rate converted to USD where applicable",
                    "Ignored entirely",
                    cost_col,
                ],
            }
        )
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
        mcol1.metric("Standard Avg Cost", money(dfm[cost_col].mean()))
        mcol2.metric("Adjusted Avg Cost", money(machine_adjusted_valid[cost_col].mean()))
        adj_cpth = machine_adjusted_valid[machine_adjusted_valid["CCR TYPE"] == "CPT+H"][cost_col].mean()
        mcol3.metric("Adjusted CPT+H Avg", money(adj_cpth) if pd.notna(adj_cpth) else "N/A")
        mcol4.metric("Machine Outliers", f"{int(dfm_all['Outlier'].sum()):,}")

        machine_type_summary = dfm.groupby("CCR TYPE").agg(
            Avg_Cost=(cost_col, "mean"),
            Avg_SMU=("SMU AT REBUILD", "mean"),
            Count=(cost_col, "count"),
        ).reset_index()
        machine_type_summary = add_vs_cmr(machine_type_summary)
        machine_type_summary["Sample Confidence"] = machine_type_summary["Count"].apply(sample_confidence)

        st.write("### Average by Rebuild Type")
        display_table(machine_type_summary, currency_cols=["Avg_Cost"], percent_cols=["Vs CMR %"], smu_cols=["Avg_SMU"], number_cols=["Count"])

        st.write("### Bar Chart: Average Rebuild Cost by Type")
        cat_bar_chart(
            machine_type_summary.rename(columns={"Avg_Cost": "Average Cost"}),
            "CCR TYPE",
            "Average Cost",
            "CCR TYPE",
            tooltip=["CCR TYPE", "Average Cost"],
        )

        machine_adjusted_summary = machine_adjusted_valid.groupby("CCR TYPE").agg(
            Adjusted_Avg_Cost=(cost_col, "mean"),
            Count=(cost_col, "count"),
        ).reset_index()
        machine_adjusted_summary["CCR TYPE"] = machine_adjusted_summary["CCR TYPE"].replace({"CPT+H": "CPT+H Adjusted"})
        machine_adjusted_summary["Sample Confidence"] = machine_adjusted_summary["Count"].apply(sample_confidence)

        st.write("### Adjusted Average Rebuild Cost by Type")
        display_table(machine_adjusted_summary, currency_cols=["Adjusted_Avg_Cost"], number_cols=["Count"])

        if not machine_adjusted_summary.empty:
            cat_bar_chart(
                machine_adjusted_summary.rename(columns={"Adjusted_Avg_Cost": "Adjusted Average Cost"}),
                "CCR TYPE",
                "Adjusted Average Cost",
                "CCR TYPE",
                tooltip=["CCR TYPE", "Adjusted Average Cost"],
            )
        else:
            st.info("No adjusted rebuild cost data available for this machine.")

        st.write("### Machine SMU vs USD Analysis Cost")
        machine_chart = dfm[["SMU AT REBUILD", cost_col, "CCR TYPE"]].dropna().rename(
            columns={"SMU AT REBUILD": "SMU", cost_col: "Cost"}
        )
        cat_scatter_chart(machine_chart, "SMU", "Cost", "CCR TYPE", tooltip=["CCR TYPE", "SMU", "Cost"])

        st.write("### Cross-Type Exception Review")
        benchmark_value = dfm_all["CMR Benchmark Cost"].dropna().iloc[0] if not dfm_all["CMR Benchmark Cost"].dropna().empty else np.nan
        machine_cross = dfm[dfm["Cross-Type Exception Flag"] != ""]
        st.dataframe(
            pd.DataFrame(
                {
                    "Metric": ["CMR Benchmark Cost", "CPT+H Cross-Type Flags", "Machine Outliers"],
                    "Value": [money(benchmark_value) if not pd.isna(benchmark_value) else "N/A", f"{len(machine_cross):,}", f"{int(dfm_all['Outlier'].sum()):,}"],
                }
            ),
            use_container_width=True,
        )

        if not machine_cross.empty:
            display_table(
                machine_cross[["DEALER", "Region", "CCR TYPE", "SMU AT REBUILD", cost_col, "CMR Benchmark Cost", "Cross-Type Exception Flag"]],
                currency_cols=[cost_col, "CMR Benchmark Cost"],
                smu_cols=["SMU AT REBUILD"],
            )
        else:
            st.write("No CPT+H cross-type exceptions for this machine.")

        st.write("### Outlier Performance")
        outlier_perf = dfm_all.groupby("CCR TYPE").agg(
            Total_Rows=(cost_col, "count"),
            Outlier_Rows=("Outlier", "sum"),
            Avg_Cost_All=(cost_col, "mean"),
        ).reset_index()
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

        dealer_summary = dealer_df.groupby("DEALER").agg(
            Avg_Cost=(cost_col, "mean"),
            Avg_SMU=("SMU AT REBUILD", "mean"),
            Count=(cost_col, "count"),
            Cross_Type_Flags=("Cross-Type Exception Flag", lambda x: (x != "").sum()),
        ).reset_index().sort_values("Avg_Cost", ascending=False)
        dealer_summary = add_vs_section_avg(dealer_summary)

        dealer_outliers = dealer_full_df.groupby("DEALER").agg(
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
        )
        dealer_summary["Performance Score"] = dealer_summary["Performance Score"].clip(lower=0, upper=100)
        dealer_summary["Performance Label"] = np.where(dealer_summary["Performance Score"] >= 85, "Strong", np.where(dealer_summary["Performance Score"] >= 70, "Monitor", "Review Needed"))

        st.write(f"### Dealer Summary - {dealer_machine}")
        display_table(
            dealer_summary,
            currency_cols=["Avg_Cost"],
            percent_cols=["Vs Section Avg %", "Outlier Rate %", "DQ Rate %", "Cross Flag Rate %"],
            smu_cols=["Avg_SMU"],
            number_cols=["Count", "Cross_Type_Flags", "Total_Rows", "Outlier_Rows", "DQ_Rows", "Performance Score"],
        )

        dealer_type = dealer_df.groupby(["DEALER", "CCR TYPE"]).agg(
            Avg_Cost=(cost_col, "mean"),
            Avg_SMU=("SMU AT REBUILD", "mean"),
            Count=(cost_col, "count"),
        ).reset_index().sort_values(["DEALER", "CCR TYPE"])

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

        region_summary = region_df.groupby("Region").agg(
            Avg_Cost=(cost_col, "mean"),
            Avg_SMU=("SMU AT REBUILD", "mean"),
            Count=(cost_col, "count"),
            Cross_Type_Flags=("Cross-Type Exception Flag", lambda x: (x != "").sum()),
        ).reset_index().sort_values("Avg_Cost", ascending=False)
        region_summary = add_vs_section_avg(region_summary)

        st.write(f"### Region Summary - {region_machine}")
        display_table(region_summary, currency_cols=["Avg_Cost"], percent_cols=["Vs Section Avg %"], smu_cols=["Avg_SMU"], number_cols=["Count", "Cross_Type_Flags"])

        region_type = region_df.groupby(["Region", "CCR TYPE"]).agg(
            Avg_Cost=(cost_col, "mean"),
            Avg_SMU=("SMU AT REBUILD", "mean"),
            Count=(cost_col, "count"),
        ).reset_index().sort_values(["Region", "CCR TYPE"])

        st.write("### Region by Rebuild Type")
        display_table(region_type, currency_cols=["Avg_Cost"], smu_cols=["Avg_SMU"], number_cols=["Count"])

        st.write("### Region Average Cost Chart")
        if not region_summary.empty:
            cat_bar_chart(region_summary.rename(columns={"Avg_Cost": "Average Cost"}), "Region", "Average Cost", None, tooltip=["Region", "Average Cost"])
        else:
            st.info("No region summary data available.")

    with tab5:
        st.subheader("Exception Summary")
        exception_summary = pd.DataFrame(
            {
                "Metric": ["Cost Outliers", "CPT+H Cross-Type Flags", "SMU 0 or 1", "Insufficient Sample Rows", "Rows Using Global Yearly Fallback Rate", "Rows Using Overall Average Fallback Rate"],
                "Value": [outlier_count, cross_count, dq_count, insufficient_count, global_year_fallback_count, overall_fallback_count],
            }
        )
        st.dataframe(exception_summary, use_container_width=True)

        st.write("### Data Quality Summary")
        st.dataframe(data_quality_summary, use_container_width=True)

        st.write("### Outlier Performance by Machine + Rebuild Type")
        outlier_perf_all = df.groupby(["SALES MODEL", "CCR TYPE"]).agg(
            Total_Rows=(cost_col, "count"),
            Outlier_Rows=("Outlier", "sum"),
            Avg_Cost_All=(cost_col, "mean"),
        ).reset_index()
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
            display_table(
                outlier_rows,
                currency_cols=["PARTS DN", "PARTS DN USD", "Rate", "Rate USD", "Labor Cost USD", "Adjusted Total Cost USD", "Inflation-Adjusted Adjusted Total Cost USD"],
                smu_cols=["SMU AT REBUILD"],
                year_cols=["Service Year"],
            )
        else:
            st.write("No cost outliers detected.")

    with tab6:
        st.subheader("Executive Insights")
        insights = []

        if not summary.empty:
            top_type = summary.sort_values("Avg_Cost", ascending=False).iloc[0]
            insights.append(f"Highest average rebuild type cost is {top_type['CCR TYPE']} at {money(top_type['Avg_Cost'])}.")

        machine_avg = valid.groupby("SALES MODEL")[cost_col].mean().sort_values(ascending=False)
        if len(machine_avg) > 0:
            insights.append(f"Highest average machine cost is {machine_avg.index[0]} at {money(machine_avg.iloc[0])}.")

        review_dealers = dealer_summary[dealer_summary["Performance Label"] == "Review Needed"] if "Performance Label" in dealer_summary.columns else pd.DataFrame()
        if not review_dealers.empty:
            insights.append(f"{len(review_dealers)} dealer(s) are marked Review Needed based on score, cost position, outlier rate, cross-type flags, and data quality.")

        insights.append(f"Cost outliers excluded from core averages: {outlier_count:,}.")
        insights.append(f"CPT+H cross-type review flags: {cross_count:,}.")
        insights.append(f"Rows using fallback labor rates: {global_year_fallback_count + overall_fallback_count:,}.")

        for insight in insights:
            st.write("•", insight)

    with tab7:
        st.subheader("Methodology")
        st.markdown(METHOD_LOCK_TEXT)
        st.markdown(
            """
            **Visual style:** Charts, checkboxes, filter controls, tabs, and workbook headers use a Caterpillar-inspired black, yellow, and gray palette.  
            **Important:** Official Caterpillar logo usage should follow internal brand/asset approval rules.
            """
        )

    with tab8:
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
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        metadata.to_excel(writer, sheet_name="Run Metadata", index=False)
        summary.to_excel(writer, sheet_name="Summary", index=False)
        adjusted_summary.to_excel(writer, sheet_name="Adjusted Summary", index=False)
        dealer_summary.to_excel(writer, sheet_name="Dealer Summary", index=False)
        dealer_type.to_excel(writer, sheet_name="Dealer By Type", index=False)
        region_summary.to_excel(writer, sheet_name="Region Summary", index=False)
        region_type.to_excel(writer, sheet_name="Region By Type", index=False)
        exception_summary.to_excel(writer, sheet_name="Exceptions", index=False)
        data_quality_summary.to_excel(writer, sheet_name="Data Quality", index=False)
        outlier_perf_all.to_excel(writer, sheet_name="Outlier Performance", index=False)
        cross_rows.to_excel(writer, sheet_name="Cross Type Flags", index=False)
        df[df["Outlier"] == True].to_excel(writer, sheet_name="Outlier Rows", index=False)
        pd.DataFrame({"Year": list(cpi_table.keys()), "CPI": list(cpi_table.values())}).sort_values("Year").to_excel(writer, sheet_name="BLS CPI", index=False)
        fx_lookup.to_excel(writer, sheet_name="FX Rates", index=False)
        rebuild_reference.to_excel(writer, sheet_name="Rebuild Type Reference", index=False)
        region_reference.to_excel(writer, sheet_name="Region Reference", index=False)

        for machine in valid["SALES MODEL"].dropna().unique():
            valid[valid["SALES MODEL"] == machine].to_excel(writer, sheet_name=safe_sheet_name(machine), index=False)

        workbook = writer.book
        header_fill = PatternFill(start_color="1F1F1F", end_color="1F1F1F", fill_type="solid")
        header_font = Font(color="FFC500", bold=True)
        thin = Side(style="thin", color="D9E2F3")
        currency_keywords = ["COST", "RATE", "PARTS DN", "LABOR", "BENCHMARK"]
        percent_keywords = ["%"]
        decimal_keywords = ["CPI", "FACTOR", "FX"]

        for ws in workbook.worksheets:
            ws.freeze_panes = "A2"
            ws.auto_filter.ref = ws.dimensions

            for cell in ws[1]:
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = Alignment(horizontal="center")

            header_map = {cell.column: str(cell.value).upper() if cell.value is not None else "" for cell in ws[1]}
            for col_idx, header in header_map.items():
                col_letter = get_column_letter(col_idx)
                ws.column_dimensions[col_letter].width = min(max(len(header) + 4, 14), 38)

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

    st.download_button("Download Workbook", data=output.getvalue(), file_name="Rebuild_Analysis_Output.xlsx")
    st.markdown(
        """
        <div class="cat-footer">
            <strong>Rebuild Analytics Platform</strong> | Internal analytics tool | Cost, inflation, dealer, region, outlier, and cross-type exception reporting
        </div>
        """,
        unsafe_allow_html=True,
    )
else:
    st.info("Upload file and run analysis")

