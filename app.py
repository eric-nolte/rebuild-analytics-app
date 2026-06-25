import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
from openpyxl.styles import Font, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

st.set_page_config(layout="wide")

st.title("Rebuild Analytics Platform V10")
st.caption("Dealer-year labor rates + labor hours + PARTS DN, component inflation, cost-only outliers, cross-type review, adjusted CPT+H view")

# =====================================================
# SESSION STATE
# =====================================================
if "run_clicked" not in st.session_state:
    st.session_state.run_clicked = False
if "analysis" not in st.session_state:
    st.session_state.analysis = None

# =====================================================
# FORMAT HELPERS
# =====================================================
def money(x):
    try:
        if pd.isna(x):
            return ""
        return f"${x:,.0f}"
    except Exception:
        return x


def pct(x):
    try:
        if pd.isna(x):
            return ""
        return f"{x:.1f}%"
    except Exception:
        return x


def display_table(df, currency_cols=None, percent_cols=None, number_cols=None, year_cols=None, smu_cols=None):
    currency_cols = currency_cols or []
    percent_cols = percent_cols or []
    number_cols = number_cols or []
    year_cols = year_cols or []
    smu_cols = smu_cols or []

    # Auto-detect year and SMU columns by name
    for col in df.columns:
        upper = str(col).upper()
        if "YEAR" in upper and col not in year_cols:
            year_cols.append(col)
        if "SMU" in upper and col not in smu_cols:
            smu_cols.append(col)

    fmt = {}
    for col in currency_cols:
        if col in df.columns:
            fmt[col] = "${:,.0f}"
    for col in percent_cols:
        if col in df.columns:
            fmt[col] = "{:.1f}%"
    for col in number_cols + year_cols + smu_cols:
        if col in df.columns:
            fmt[col] = "{:,.0f}"

    if fmt:
        st.dataframe(df.style.format(fmt), use_container_width=True)
    else:
        st.dataframe(df, use_container_width=True)


def safe_sheet_name(name):
    invalid = ['\\', '/', '*', '[', ']', ':', '?']
    name = str(name)
    for ch in invalid:
        name = name.replace(ch, "-")
    return name[:31]


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
# DEALER RATE BUILDER
# =====================================================
def build_default_rate_table(start=2010, end=2030):
    years = list(range(start, end + 1))
    return pd.DataFrame({
        "Dealer Code": ["DEFAULT"] * len(years),
        "Service Year": years,
        "Rate": [115 + (y - 2016) * 3 for y in years]
    })


def build_rate_table_from_workbook(rate_file):
    rates = pd.read_excel(rate_file, sheet_name=None)
    rows = []

    for sheet_name, sheet in rates.items():
        if str(sheet_name).strip().lower() == "summary":
            continue

        temp = sheet.copy()
        temp.columns = [str(c).strip() for c in temp.columns]

        rename_map = {}
        for c in temp.columns:
            cu = c.upper()
            if cu == "YEAR":
                rename_map[c] = "Service Year"
            elif "BASE RATE" in cu or cu == "RATE":
                rename_map[c] = "Rate"
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
# INPUTS
# =====================================================
rebuild_file = st.file_uploader("Upload Rebuild File", type=["xlsx"])

st.subheader("Dealer Labor Rates")
use_default = st.checkbox("Use Built-in Default Dealer Rates", True)
rate_file = None
if not use_default:
    rate_file = st.file_uploader("Upload Custom Dealer-by-Year Rate Workbook", type=["xlsx"])

st.info(
    "Cost basis: Adjusted Total Cost = PARTS DN + Labor Cost. "
    "Labor Cost = Labor Hours × dealer service-year labor rate. "
    "Inflation, when enabled, is applied by component: parts and base rate."
)

apply_inflation = st.checkbox("Apply Inflation", True)
base_year = st.number_input("Base Year for Inflation", 2010, 2030, 2026)
inflation_rate_pct = st.number_input("Annual Inflation Assumption (%)", 0.0, 20.0, 3.0, step=0.25)

machine_input = st.text_input("Machines (optional, comma-separated; leave blank for all)")
start_year = st.number_input("Start Year", 2010, 2030, 2016)
end_year = st.number_input("End Year", 2010, 2030, 2026)

rebuild_filter = st.multiselect(
    "Filter Rebuild Types",
    options=["CMR", "CPT+H", "CPT-O"],
    default=["CMR", "CPT+H", "CPT-O"]
)

if st.button("Run Analysis"):
    st.session_state.run_clicked = True
    st.session_state.analysis = None

# =====================================================
# ANALYSIS FUNCTION
# =====================================================
def run_analysis(rebuild_file, rate_file):
    rebuild = pd.read_excel(rebuild_file, sheet_name=None)

    # ---------- Combine and assign CCR TYPE by sheet name ----------
    frames = []
    for sheet_name, sheet in rebuild.items():
        sheet = sheet.copy()
        sheet_upper = str(sheet_name).upper()
        if "CMR" in sheet_upper:
            sheet["CCR TYPE"] = "CMR"
        elif "CPT+H" in sheet_upper or "CPTH" in sheet_upper:
            sheet["CCR TYPE"] = "CPT+H"
        else:
            sheet["CCR TYPE"] = "CPT-O"
        frames.append(sheet)

    df = pd.concat(frames, ignore_index=True)

    # ---------- Required columns ----------
    required_cols = [
        "DELIVERED DATE", "ENROLL DATE", "IN SERVICE DATE", "SALES MODEL",
        "DEALER", "PARTS DN", "SMU AT REBUILD", "REBUILD WORK HRS"
    ]
    missing = [c for c in required_cols if c not in df.columns]
    if missing:
        raise ValueError(f"Missing required columns: {missing}")

    # ---------- Filters ----------
    df = df[df["CCR TYPE"].isin(rebuild_filter)].copy()
    df["Service Date"] = df["DELIVERED DATE"].fillna(df["ENROLL DATE"]).fillna(df["IN SERVICE DATE"])
    df["Service Year"] = pd.to_datetime(df["Service Date"], errors="coerce").dt.year
    df = df[(df["Service Year"] >= start_year) & (df["Service Year"] <= end_year)].copy()

    if machine_input:
        machine_list = [m.strip() for m in machine_input.split(",")]
        df = df[df["SALES MODEL"].isin(machine_list)].copy()

    if df.empty:
        raise ValueError("No rows remain after filters. Adjust machine, year, or rebuild type filters.")

    # ---------- Dealer and Region ----------
    df["Dealer Code"] = df["DEALER"].astype(str).str.extract(r"([A-Z]\d{3})")
    if "REGION" in df.columns:
        df["Region"] = df["REGION"].fillna("Unknown")
    else:
        df["Region"] = "Unknown"

    # ---------- Rate logic ----------
    if use_default or rate_file is None:
        rate_df = build_default_rate_table(2010, 2030)
    else:
        rate_df = build_rate_table_from_workbook(rate_file)

    rate_df["Service Year"] = pd.to_numeric(rate_df["Service Year"], errors="coerce").astype("Int64")
    df["Service Year"] = pd.to_numeric(df["Service Year"], errors="coerce").astype("Int64")

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

    # ---------- Numeric cost fields ----------
    df["PARTS DN"] = pd.to_numeric(df["PARTS DN"], errors="coerce")
    df["REBUILD WORK HRS"] = pd.to_numeric(df["REBUILD WORK HRS"], errors="coerce")
    df["SMU AT REBUILD"] = pd.to_numeric(df["SMU AT REBUILD"], errors="coerce")

    df = df[(df["PARTS DN"] > 0) & (df["REBUILD WORK HRS"].notna())].copy()

    # ---------- Cost logic ----------
    annual_rate = inflation_rate_pct / 100
    df["Labor Cost"] = df["REBUILD WORK HRS"] * df["Rate"]
    df["Adjusted Total Cost"] = df["PARTS DN"] + df["Labor Cost"]

    if apply_inflation:
        df["Inflation Factor"] = df["Service Year"].apply(
            lambda y: (1 + annual_rate) ** (base_year - int(y)) if pd.notnull(y) else 1
        )
    else:
        df["Inflation Factor"] = 1.0

    # Component inflation
    df["Inflation-Adjusted Parts DN"] = df["PARTS DN"] * df["Inflation Factor"]
    df["Inflation-Adjusted Base Rate"] = df["Rate"] * df["Inflation Factor"]
    df["Inflation-Adjusted Labor Cost"] = df["REBUILD WORK HRS"] * df["Inflation-Adjusted Base Rate"]
    df["Inflation-Adjusted Adjusted Total Cost"] = (
        df["Inflation-Adjusted Parts DN"] + df["Inflation-Adjusted Labor Cost"]
    )

    cost_col = "Inflation-Adjusted Adjusted Total Cost" if apply_inflation else "Adjusted Total Cost"
    df = df[df[cost_col] > 0].copy()

    if df.empty:
        raise ValueError("No valid cost rows remain after processing.")

    # ---------- Data quality flags ----------
    df["Data Quality Exception Flag"] = ""
    df.loc[df["SMU AT REBUILD"].isin([0, 1]), "Data Quality Exception Flag"] = "SMU 0 or 1"

    # ---------- Outliers by Machine + CCR TYPE ----------
    df["Outlier"] = False
    df["Outlier Reason"] = ""
    df["Insufficient Sample Group"] = False

    for (machine, ccr), g in df.groupby(["SALES MODEL", "CCR TYPE"], dropna=False):
        idx = g.index
        n = len(g)
        if n < 5:
            df.loc[idx, "Insufficient Sample Group"] = True
            df.loc[idx, "Outlier Reason"] = "Insufficient sample; no statistical outlier removal"
            continue

        log_vals = np.log(g[cost_col])
        q1, q3 = np.percentile(log_vals, [25, 75])
        iqr = q3 - q1
        mult = 2.0 if n < 8 else 1.5
        low, high = q1 - mult * iqr, q3 + mult * iqr
        mask = (log_vals < low) | (log_vals > high)
        df.loc[idx, "Outlier"] = mask
        df.loc[idx[mask], "Outlier Reason"] = "Cost outlier: log(cost) IQR rule"

    valid = df[df["Outlier"] == False].copy()
    if valid.empty:
        raise ValueError("All rows were excluded as cost outliers. Review data or filters.")

    # ---------- Cross-Type Cost Exception Rule ----------
    df["Cross-Type Exception Flag"] = ""
    df["CMR Benchmark Cost"] = np.nan

    for machine, group in df.groupby("SALES MODEL", dropna=False):
        cmr_valid = group[(group["CCR TYPE"] == "CMR") & (group["Outlier"] == False)]
        if len(cmr_valid) >= 3:
            benchmark = cmr_valid[cost_col].median()
            df.loc[group.index, "CMR Benchmark Cost"] = benchmark
            mask = (
                (group["CCR TYPE"] == "CPT+H") &
                (group["Outlier"] == False) &
                (group[cost_col] > benchmark * 1.10)
            )
            df.loc[group.index[mask], "Cross-Type Exception Flag"] = "CPT+H Cost Above Typical CMR"

    valid = df[df["Outlier"] == False].copy()

    # ---------- Standard averages ----------
    summary = valid.groupby("CCR TYPE").agg(
        Avg_Cost=(cost_col, "mean"),
        Avg_SMU=("SMU AT REBUILD", "mean"),
        Count=(cost_col, "count")
    ).reset_index()
    summary = add_vs_cmr(summary)

    # ---------- Adjusted averages ----------
    adjusted_valid = valid.copy()
    adjusted_valid = adjusted_valid[~(
        (adjusted_valid["CCR TYPE"] == "CPT+H") &
        (adjusted_valid["Cross-Type Exception Flag"] == "CPT+H Cost Above Typical CMR")
    )].copy()

    adjusted_summary = adjusted_valid.groupby("CCR TYPE").agg(
        Adjusted_Avg_Cost=(cost_col, "mean"),
        Count=(cost_col, "count")
    ).reset_index()
    adjusted_summary["CCR TYPE"] = adjusted_summary["CCR TYPE"].replace({"CPT+H": "CPT+H Adjusted"})

    return {
        "df": df,
        "valid": valid,
        "adjusted_valid": adjusted_valid,
        "summary": summary,
        "adjusted_summary": adjusted_summary,
        "cost_col": cost_col,
        "outlier_count": int(df["Outlier"].sum()),
        "cross_count": int((valid["Cross-Type Exception Flag"] != "").sum()),
        "dq_count": int(df["Data Quality Exception Flag"].eq("SMU 0 or 1").sum()),
        "insufficient_count": int(df["Insufficient Sample Group"].sum()),
        "global_year_fallback_count": int((df["Rate Source"] == "Global Yearly Fallback Rate").sum()),
        "overall_fallback_count": int((df["Rate Source"] == "Overall Average Fallback Rate").sum())
    }

# =====================================================
# RUN / LOAD ANALYSIS
# =====================================================
if st.session_state.run_clicked and rebuild_file:
    if st.session_state.analysis is None:
        try:
            st.session_state.analysis = run_analysis(rebuild_file, rate_file)
        except Exception as e:
            st.error(str(e))
            st.stop()

    analysis = st.session_state.analysis
    df = analysis["df"]
    valid = analysis["valid"]
    adjusted_valid = analysis["adjusted_valid"]
    summary = analysis["summary"]
    adjusted_summary = analysis["adjusted_summary"]
    cost_col = analysis["cost_col"]
    outlier_count = analysis["outlier_count"]
    cross_count = analysis["cross_count"]
    dq_count = analysis["dq_count"]
    insufficient_count = analysis["insufficient_count"]
    global_year_fallback_count = analysis["global_year_fallback_count"]
    overall_fallback_count = analysis["overall_fallback_count"]

    tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs([
        "Dashboard", "Machine Detail", "Dealer", "Region", "Exceptions", "Methodology"
    ])

    # =====================================================
    # DASHBOARD
    # =====================================================
    with tab1:
        st.subheader("Executive Summary")
        col1, col2, col3, col4 = st.columns(4)
        col1.metric("Valid Rows", f"{len(valid):,}")
        col2.metric("Cost Outliers", f"{outlier_count:,}")
        col3.metric("Avg Analysis Cost", money(valid[cost_col].mean()))
        col4.metric("Cross-Type Flags", f"{cross_count:,}")

        st.write("### Standard Rebuild Type Average Cost")
        display_table(summary, currency_cols=["Avg_Cost"], percent_cols=["Vs CMR %"], number_cols=["Avg_SMU", "Count"], smu_cols=["Avg_SMU"])

        st.write("### Adjusted Rebuild Type Average Cost")
        display_table(adjusted_summary, currency_cols=["Adjusted_Avg_Cost"], number_cols=["Count"])

        st.write("### SMU vs Analysis Cost by Rebuild Type")
        scatter = valid[["SMU AT REBUILD", cost_col, "CCR TYPE"]].dropna().rename(
            columns={"SMU AT REBUILD": "SMU", cost_col: "Cost"}
        )
        st.scatter_chart(scatter, x="SMU", y="Cost", color="CCR TYPE")

        st.write("### Cost and Inflation Settings")
        inflation_settings = pd.DataFrame({
            "Setting": [
                "Cost Source", "Labor Cost", "PLUS PARTS DN", "Inflation Applied", "Base Year",
                "Annual Inflation Assumption", "Analysis Cost Used"
            ],
            "Value": [
                "PARTS DN + Labor Cost",
                "Labor Hours × Dealer Service-Year Rate",
                "Ignored entirely",
                "Yes" if apply_inflation else "No",
                str(base_year),
                pct(inflation_rate_pct),
                cost_col
            ]
        })
        st.dataframe(inflation_settings, use_container_width=True)

    # =====================================================
    # MACHINE DETAIL
    # =====================================================
    with tab2:
        st.subheader("Machine Detail")
        machine_list = sorted(valid["SALES MODEL"].dropna().unique())
        selected_machine = st.selectbox("Select Machine", machine_list, key="machine_detail_selector")
        dfm = valid[valid["SALES MODEL"] == selected_machine].copy()
        dfm_all = df[df["SALES MODEL"] == selected_machine].copy()

        st.write(f"### Machine: {selected_machine}")

        machine_type_summary = dfm.groupby("CCR TYPE").agg(
            Avg_Cost=(cost_col, "mean"),
            Avg_SMU=("SMU AT REBUILD", "mean"),
            Count=(cost_col, "count")
        ).reset_index()
        machine_type_summary = add_vs_cmr(machine_type_summary)

        st.write("### Average by Rebuild Type")
        display_table(machine_type_summary, currency_cols=["Avg_Cost"], percent_cols=["Vs CMR %"], number_cols=["Avg_SMU", "Count"], smu_cols=["Avg_SMU"])

        st.write("### Bar Chart: Average Rebuild Cost by Type")
        bar_df = machine_type_summary[["CCR TYPE", "Avg_Cost"]].rename(columns={"Avg_Cost": "Average Cost"}).set_index("CCR TYPE")
        st.bar_chart(bar_df)

        machine_adjusted_valid = dfm[~((dfm["CCR TYPE"] == "CPT+H") & (dfm["Cross-Type Exception Flag"] == "CPT+H Cost Above Typical CMR"))].copy()
        machine_adjusted_summary = machine_adjusted_valid.groupby("CCR TYPE").agg(
            Adjusted_Avg_Cost=(cost_col, "mean"),
            Count=(cost_col, "count")
        ).reset_index()
        machine_adjusted_summary["CCR TYPE"] = machine_adjusted_summary["CCR TYPE"].replace({"CPT+H": "CPT+H Adjusted"})

        st.write("### Adjusted Average Rebuild Cost by Type")
        display_table(machine_adjusted_summary, currency_cols=["Adjusted_Avg_Cost"], number_cols=["Count"])
        if not machine_adjusted_summary.empty:
            adjusted_bar = machine_adjusted_summary[["CCR TYPE", "Adjusted_Avg_Cost"]].rename(columns={"Adjusted_Avg_Cost": "Adjusted Average Cost"}).set_index("CCR TYPE")
            st.bar_chart(adjusted_bar)

        st.write("### Machine SMU vs Analysis Cost")
        machine_chart = dfm[["SMU AT REBUILD", cost_col, "CCR TYPE"]].dropna().rename(
            columns={"SMU AT REBUILD": "SMU", cost_col: "Cost"}
        )
        st.scatter_chart(machine_chart, x="SMU", y="Cost", color="CCR TYPE")

        st.write("### Cross-Type Exception Review")
        benchmark_value = dfm_all["CMR Benchmark Cost"].dropna().iloc[0] if not dfm_all["CMR Benchmark Cost"].dropna().empty else np.nan
        machine_cross = dfm[dfm["Cross-Type Exception Flag"] != ""]
        cross_review_summary = pd.DataFrame({
            "Metric": ["CMR Benchmark Cost", "CPT+H Cross-Type Flags", "Machine Outliers"],
            "Value": [money(benchmark_value) if not pd.isna(benchmark_value) else "N/A", f"{len(machine_cross):,}", f"{int(dfm_all['Outlier'].sum()):,}"]
        })
        st.dataframe(cross_review_summary, use_container_width=True)
        if not machine_cross.empty:
            display_table(
                machine_cross[["DEALER", "Region", "CCR TYPE", "SMU AT REBUILD", cost_col, "CMR Benchmark Cost", "Cross-Type Exception Flag"]],
                currency_cols=[cost_col, "CMR Benchmark Cost"],
                smu_cols=["SMU AT REBUILD"]
            )
        else:
            st.write("No CPT+H cross-type exceptions for this machine.")

        st.write("### Outlier Performance")
        outlier_perf = dfm_all.groupby("CCR TYPE").agg(
            Total_Rows=(cost_col, "count"),
            Outlier_Rows=("Outlier", "sum"),
            Avg_Cost_All=(cost_col, "mean")
        ).reset_index()
        valid_counts = dfm.groupby("CCR TYPE").size().reset_index(name="Valid_Rows")
        outlier_perf = outlier_perf.merge(valid_counts, on="CCR TYPE", how="left").fillna({"Valid_Rows": 0})
        outlier_perf["Outlier Rate %"] = (outlier_perf["Outlier_Rows"] / outlier_perf["Total_Rows"]) * 100
        display_table(outlier_perf, currency_cols=["Avg_Cost_All"], percent_cols=["Outlier Rate %"], number_cols=["Total_Rows", "Outlier_Rows", "Valid_Rows"])

        st.write("### Machine Raw Data Preview")
        raw_cols = [
            "DEALER", "Region", "CCR TYPE", "Service Year", "SMU AT REBUILD", "PARTS DN",
            "REBUILD WORK HRS", "Rate", "Labor Cost", "Adjusted Total Cost", "Inflation Factor",
            "Inflation-Adjusted Parts DN", "Inflation-Adjusted Base Rate", "Inflation-Adjusted Labor Cost",
            "Inflation-Adjusted Adjusted Total Cost", "Cross-Type Exception Flag", "Data Quality Exception Flag", "Outlier Reason"
        ]
        raw_cols = [c for c in raw_cols if c in dfm.columns]
        display_table(
            dfm[raw_cols].head(50),
            currency_cols=[
                "PARTS DN", "Rate", "Labor Cost", "Adjusted Total Cost", "Inflation-Adjusted Parts DN",
                "Inflation-Adjusted Base Rate", "Inflation-Adjusted Labor Cost", "Inflation-Adjusted Adjusted Total Cost"
            ],
            year_cols=["Service Year"],
            smu_cols=["SMU AT REBUILD"],
            number_cols=["REBUILD WORK HRS"]
        )

    # =====================================================
    # DEALER SECTION
    # =====================================================
    with tab3:
        st.subheader("Dealer Analysis")
        dealer_machine_options = ["All Machines"] + sorted(valid["SALES MODEL"].dropna().unique().tolist())
        dealer_machine = st.selectbox("Focus Dealer Analysis on Machine", dealer_machine_options, key="dealer_machine_selector")
        dealer_df = valid.copy()
        if dealer_machine != "All Machines":
            dealer_df = dealer_df[dealer_df["SALES MODEL"] == dealer_machine]

        dealer_summary = dealer_df.groupby("DEALER").agg(
            Avg_Cost=(cost_col, "mean"),
            Avg_SMU=("SMU AT REBUILD", "mean"),
            Count=(cost_col, "count"),
            Cross_Type_Flags=("Cross-Type Exception Flag", lambda x: (x != "").sum())
        ).reset_index().sort_values("Avg_Cost", ascending=False)
        dealer_summary = add_vs_section_avg(dealer_summary)
        st.write(f"### Dealer Summary - {dealer_machine}")
        display_table(dealer_summary, currency_cols=["Avg_Cost"], percent_cols=["Vs Section Avg %"], smu_cols=["Avg_SMU"], number_cols=["Count", "Cross_Type_Flags"])

        dealer_type = dealer_df.groupby(["DEALER", "CCR TYPE"]).agg(
            Avg_Cost=(cost_col, "mean"),
            Avg_SMU=("SMU AT REBUILD", "mean"),
            Count=(cost_col, "count")
        ).reset_index().sort_values(["DEALER", "CCR TYPE"])
        st.write("### Dealer by Rebuild Type")
        display_table(dealer_type, currency_cols=["Avg_Cost"], smu_cols=["Avg_SMU"], number_cols=["Count"])

        st.write("### Dealer Average Cost Chart")
        if not dealer_summary.empty:
            st.bar_chart(dealer_summary[["DEALER", "Avg_Cost"]].rename(columns={"Avg_Cost": "Average Cost"}).set_index("DEALER"))

    # =====================================================
    # REGION SECTION
    # =====================================================
    with tab4:
        st.subheader("Region Analysis")
        region_machine_options = ["All Machines"] + sorted(valid["SALES MODEL"].dropna().unique().tolist())
        region_machine = st.selectbox("Focus Region Analysis on Machine", region_machine_options, key="region_machine_selector")
        region_df = valid.copy()
        if region_machine != "All Machines":
            region_df = region_df[region_df["SALES MODEL"] == region_machine]

        region_summary = region_df.groupby("Region").agg(
            Avg_Cost=(cost_col, "mean"),
            Avg_SMU=("SMU AT REBUILD", "mean"),
            Count=(cost_col, "count"),
            Cross_Type_Flags=("Cross-Type Exception Flag", lambda x: (x != "").sum())
        ).reset_index().sort_values("Avg_Cost", ascending=False)
        region_summary = add_vs_section_avg(region_summary)
        st.write(f"### Region Summary - {region_machine}")
        display_table(region_summary, currency_cols=["Avg_Cost"], percent_cols=["Vs Section Avg %"], smu_cols=["Avg_SMU"], number_cols=["Count", "Cross_Type_Flags"])

        region_type = region_df.groupby(["Region", "CCR TYPE"]).agg(
            Avg_Cost=(cost_col, "mean"),
            Avg_SMU=("SMU AT REBUILD", "mean"),
            Count=(cost_col, "count")
        ).reset_index().sort_values(["Region", "CCR TYPE"])
        st.write("### Region by Rebuild Type")
        display_table(region_type, currency_cols=["Avg_Cost"], smu_cols=["Avg_SMU"], number_cols=["Count"])

        st.write("### Region Average Cost Chart")
        if not region_summary.empty:
            st.bar_chart(region_summary[["Region", "Avg_Cost"]].rename(columns={"Avg_Cost": "Average Cost"}).set_index("Region"))

    # =====================================================
    # EXCEPTIONS
    # =====================================================
    with tab5:
        st.subheader("Exception Summary")
        exception_summary = pd.DataFrame({
            "Metric": [
                "Cost Outliers", "CPT+H Cross-Type Flags", "SMU 0 or 1", "Insufficient Sample Rows",
                "Rows Using Global Yearly Fallback Rate", "Rows Using Overall Average Fallback Rate"
            ],
            "Value": [outlier_count, cross_count, dq_count, insufficient_count, global_year_fallback_count, overall_fallback_count]
        })
        st.dataframe(exception_summary, use_container_width=True)

        st.write("### Outlier Performance by Machine + Rebuild Type")
        outlier_perf_all = df.groupby(["SALES MODEL", "CCR TYPE"]).agg(
            Total_Rows=(cost_col, "count"),
            Outlier_Rows=("Outlier", "sum"),
            Avg_Cost_All=(cost_col, "mean")
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
            display_table(outlier_rows, currency_cols=["PARTS DN", "Rate", "Labor Cost", "Adjusted Total Cost", "Inflation-Adjusted Adjusted Total Cost"], smu_cols=["SMU AT REBUILD"], year_cols=["Service Year"])
        else:
            st.write("No cost outliers detected.")

    # =====================================================
    # METHODOLOGY
    # =====================================================
    with tab6:
        st.subheader("Methodology")
        st.markdown(
            """
            **Cost basis:** Adjusted Total Cost = PARTS DN + Labor Cost. PLUS PARTS DN is ignored entirely.  
            **Labor:** Labor Cost = REBUILD WORK HRS × dealer service-year labor rate. Missing dealer-year rates fall back to yearly average, then overall average.  
            **Inflation:** Inflation is applied by component: PARTS DN × factor and Base Rate × factor. Inflation-adjusted labor is rebuilt as labor hours × inflation-adjusted base rate.  
            **Outliers:** Cost-only statistical screening using log(cost) + IQR by Machine + CCR TYPE.  
            **Sample-size rules:** <5 records = no removal; 5–7 records = 2.0×IQR; 8+ records = 1.5×IQR.  
            **SMU:** SMU 0 or 1 is a data-quality flag only and is not removed as a statistical outlier.  
            **Cross-type exception:** CPT+H is flagged when above machine-level valid CMR median × 1.10, when at least 3 valid CMR rows exist.  
            **Adjusted CPT+H:** Adjusted averages exclude cross-type flagged CPT+H rows only; CMR and CPT-O stay unchanged.
            """
        )

    # =====================================================
    # EXPORT
    # =====================================================
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        summary.to_excel(writer, sheet_name="Summary", index=False)
        adjusted_summary.to_excel(writer, sheet_name="Adjusted Summary", index=False)
        dealer_summary.to_excel(writer, sheet_name="Dealer Summary", index=False)
        dealer_type.to_excel(writer, sheet_name="Dealer By Type", index=False)
        region_summary.to_excel(writer, sheet_name="Region Summary", index=False)
        region_type.to_excel(writer, sheet_name="Region By Type", index=False)
        exception_summary.to_excel(writer, sheet_name="Exceptions", index=False)
        outlier_perf_all.to_excel(writer, sheet_name="Outlier Performance", index=False)
        cross_rows.to_excel(writer, sheet_name="Cross Type Flags", index=False)
        df[df["Outlier"] == True].to_excel(writer, sheet_name="Outlier Rows", index=False)
        for m in valid["SALES MODEL"].dropna().unique():
            valid[valid["SALES MODEL"] == m].to_excel(writer, sheet_name=safe_sheet_name(m), index=False)

        # Excel number formatting for years, SMU, currency, and percentages
        workbook = writer.book
        currency_keywords = ["COST", "RATE", "PARTS DN", "LABOR"]
        percent_keywords = ["%"]
        for ws in workbook.worksheets:
            header_map = {cell.column: str(cell.value).upper() if cell.value is not None else "" for cell in ws[1]}
            for col_idx, header in header_map.items():
                col_letter = get_column_letter(col_idx)
                for cell in ws[col_letter][1:]:
                    if "YEAR" in header or "SMU" in header or header in ["COUNT", "TOTAL_ROWS", "OUTLIER_ROWS", "VALID_ROWS", "VALUE"]:
                        cell.number_format = '#,##0'
                    elif any(k in header for k in percent_keywords):
                        cell.number_format = '0.0%'
                    elif any(k in header for k in currency_keywords):
                        cell.number_format = '$#,##0'

    st.download_button(
        "Download Workbook",
        data=output.getvalue(),
        file_name="Rebuild_Analysis_Output.xlsx"
    )

else:
    st.info("Upload file and run analysis")

