import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO

st.set_page_config(layout="wide")

st.title("Rebuild Analytics Platform V7")
st.caption("Executive rebuild analytics with machine, dealer, and region breakdowns")

# =============================
# SESSION STATE
# =============================
if "run_clicked" not in st.session_state:
    st.session_state.run_clicked = False

# =============================
# FORMAT HELPERS
# =============================
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


def display_table(df, currency_cols=None, percent_cols=None):
    currency_cols = currency_cols or []
    percent_cols = percent_cols or []
    fmt = {}
    for col in currency_cols:
        if col in df.columns:
            fmt[col] = "${:,.0f}"
    for col in percent_cols:
        if col in df.columns:
            fmt[col] = "{:.1f}%"
    if fmt:
        st.dataframe(df.style.format(fmt), use_container_width=True)
    else:
        st.dataframe(df, use_container_width=True)


def add_vs_machine_avg(table, avg_col="Avg_Cost"):
    if table.empty or avg_col not in table.columns:
        return table
    overall_avg = table[avg_col].mean()
    if overall_avg and not pd.isna(overall_avg):
        table["Vs Section Avg %"] = ((table[avg_col] - overall_avg) / overall_avg) * 100
    else:
        table["Vs Section Avg %"] = np.nan
    return table


def add_vs_cmr(table, type_col="CCR TYPE", avg_col="Avg_Cost"):
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


# =============================
# INPUTS
# =============================
rebuild_file = st.file_uploader("Upload Rebuild File", type=["xlsx"])

st.subheader("Dealer Rates")
use_default = st.checkbox("Use Built-in Dealer Rates", True)
rate_file = None
if not use_default:
    rate_file = st.file_uploader("Upload Custom Dealer Rates", type=["xlsx"])

machine_input = st.text_input("Machines (optional, comma-separated)")
start_year = st.number_input("Start Year", 2010, 2030, 2016)
end_year = st.number_input("End Year", 2010, 2030, 2026)
base_year = st.number_input("Base Year", 2010, 2030, 2026)

rebuild_filter = st.multiselect(
    "Filter Rebuild Types",
    options=["CMR", "CPT+H", "CPT-O"],
    default=["CMR", "CPT+H", "CPT-O"]
)

if st.button("Run Analysis"):
    st.session_state.run_clicked = True


# =============================
# INSIGHTS
# =============================
def generate_insights(df, summary):
    insights = []

    try:
        if "CMR" in summary["CCR TYPE"].values and "CPT+H" in summary["CCR TYPE"].values:
            cmr = summary.loc[summary["CCR TYPE"] == "CMR", "Avg_Cost"].values[0]
            cpth = summary.loc[summary["CCR TYPE"] == "CPT+H", "Avg_Cost"].values[0]
            diff = ((cpth - cmr) / cmr) * 100 if cmr else 0
            direction = "higher" if diff >= 0 else "lower"
            insights.append(
                f"CPT+H average cost is {abs(diff):.1f}% {direction} than CMR."
            )
    except Exception:
        pass

    try:
        high_var = df.groupby("CCR TYPE")["Adj Cost"].std().sort_values(ascending=False)
        if len(high_var) > 0:
            insights.append(f"Highest cost variability is currently observed in {high_var.index[0]}.")
    except Exception:
        pass

    try:
        highest_machine = df.groupby("SALES MODEL")["Adj Cost"].mean().sort_values(ascending=False)
        if len(highest_machine) > 0:
            insights.append(
                f"Highest average machine cost is {highest_machine.index[0]} at {money(highest_machine.iloc[0])}."
            )
    except Exception:
        pass

    return insights


# =============================
# MAIN PROCESSING
# =============================
if st.session_state.run_clicked and rebuild_file:

    rebuild = pd.read_excel(rebuild_file, sheet_name=None)

    # ---------- Combine ----------
    frames = []
    for name, sheet in rebuild.items():
        sheet = sheet.copy()

        if "CMR" in name:
            sheet["CCR TYPE"] = "CMR"
        elif "CPT+H" in name:
            sheet["CCR TYPE"] = "CPT+H"
        else:
            sheet["CCR TYPE"] = "CPT-O"

        frames.append(sheet)

    df = pd.concat(frames, ignore_index=True)

    if "CCR TYPE" not in df.columns:
        df["CCR TYPE"] = "UNKNOWN"

    df = df[df["CCR TYPE"].isin(rebuild_filter)]

    # ---------- Required field handling ----------
    required_cols = [
        "DELIVERED DATE", "ENROLL DATE", "IN SERVICE DATE", "SALES MODEL",
        "DEALER", "REBUILD WORK HRS", "PARTS DN", "SMU AT REBUILD"
    ]
    missing = [c for c in required_cols if c not in df.columns]
    if missing:
        st.error(f"Missing required columns: {missing}")
        st.stop()

    # ---------- Dates ----------
    df["Service Date"] = df["DELIVERED DATE"].fillna(df["ENROLL DATE"]).fillna(df["IN SERVICE DATE"])
    df["Service Year"] = pd.to_datetime(df["Service Date"], errors="coerce").dt.year

    df = df[(df["Service Year"] >= start_year) & (df["Service Year"] <= end_year)]

    # ---------- Machine Filter ----------
    if machine_input:
        machines = [m.strip() for m in machine_input.split(",")]
        df = df[df["SALES MODEL"].isin(machines)]

    if df.empty:
        st.warning("No rows remain after filters. Adjust machine, year, or rebuild type filters.")
        st.stop()

    # ---------- Dealer and Region ----------
    df["Dealer Code"] = df["DEALER"].astype(str).str.extract(r"([A-Z]\d{3})")
    if "REGION" in df.columns:
        df["Region"] = df["REGION"].fillna("Unknown")
    else:
        df["Region"] = df["Dealer Code"].str[0].fillna("Unknown")

    # ---------- Rates ----------
    if use_default or not rate_file:
        years = list(range(2010, 2031))
        rate_df = pd.DataFrame({
            "Service Year": years,
            "Rate": [115 + (y - 2016) * 3 for y in years]
        })
        df = df.merge(rate_df, on="Service Year", how="left")
    else:
        rates = pd.read_excel(rate_file, sheet_name=None)
        rows = []
        for k, v in rates.items():
            if k != "Summary":
                temp = v.copy()
                temp = temp.rename(columns={"Year": "Service Year", "Base Rate (USD)": "Rate"})
                if "Service Year" in temp.columns and "Rate" in temp.columns:
                    rows.append(temp[["Service Year", "Rate"]])
        rate_df = pd.concat(rows) if rows else pd.DataFrame(columns=["Service Year", "Rate"])
        df = df.merge(rate_df, on="Service Year", how="left")

    df["Base Rate"] = df["Rate"].fillna(df["Rate"].mean())

    # ---------- Cost ----------
    df["Labor Cost"] = df["REBUILD WORK HRS"] * df["Base Rate"]
    df["Infl"] = df["Service Year"].apply(
        lambda y: 1 + (base_year - y) * 0.03 if pd.notnull(y) else 1
    )
    df["Adj Cost"] = df["PARTS DN"] * df["Infl"] + df["Labor Cost"] * df["Infl"]
    df = df[df["Adj Cost"] > 0]

    if df.empty:
        st.warning("No valid cost rows remain after processing.")
        st.stop()

    # ---------- Outliers ----------
    df["Outlier"] = False

    for (m, t), g in df.groupby(["SALES MODEL", "CCR TYPE"]):
        if len(g) >= 5:
            log_vals = np.log(g["Adj Cost"])
            q1, q3 = np.percentile(log_vals, [25, 75])
            iqr = q3 - q1
            mult = 2 if len(g) < 8 else 1.5
            low, high = q1 - mult * iqr, q3 + mult * iqr
            mask = (log_vals < low) | (log_vals > high)
            df.loc[g.index, "Outlier"] = mask

    valid = df[df["Outlier"] == False].copy()

    if valid.empty:
        st.warning("All rows were excluded as outliers. Review data or filters.")
        st.stop()

    # ---------- Summary ----------
    summary = valid.groupby("CCR TYPE").agg(
        Avg_Cost=("Adj Cost", "mean"),
        Avg_SMU=("SMU AT REBUILD", "mean"),
        Count=("Adj Cost", "count")
    ).reset_index()
    summary = add_vs_cmr(summary)

    insights = generate_insights(valid, summary)

    # =============================
    # TABS
    # =============================
    tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs([
        "Dashboard", "Machine Detail", "Dealer", "Region", "Exceptions", "Insights"
    ])

    # =============================
    # DASHBOARD
    # =============================
    with tab1:
        st.subheader("Executive Summary")

        col1, col2, col3, col4 = st.columns(4)
        col1.metric("Total Rows", f"{len(valid):,}")
        col2.metric("Outliers", f"{int(df['Outlier'].sum()):,}")
        col3.metric("Avg Cost", money(valid["Adj Cost"].mean()))

        if "CMR" in summary["CCR TYPE"].values and "CPT+H" in summary["CCR TYPE"].values:
            cmr = summary.loc[summary["CCR TYPE"] == "CMR", "Avg_Cost"].values[0]
            cpth = summary.loc[summary["CCR TYPE"] == "CPT+H", "Avg_Cost"].values[0]
            delta = ((cpth - cmr) / cmr) * 100 if cmr else np.nan
            col4.metric("CPT+H vs CMR", pct(delta))
        else:
            col4.metric("CPT+H vs CMR", "N/A")

        st.write("### Rebuild Type Summary")
        display_table(summary, currency_cols=["Avg_Cost"], percent_cols=["Vs CMR %"])

        st.write("### SMU vs Cost by Rebuild Type")
        chart_data = valid[["SMU AT REBUILD", "Adj Cost", "CCR TYPE"]].dropna().rename(
            columns={"SMU AT REBUILD": "SMU", "Adj Cost": "Cost"}
        )
        st.scatter_chart(chart_data, x="SMU", y="Cost", color="CCR TYPE")

    # =============================
    # MACHINE DETAIL
    # =============================
    with tab2:
        st.subheader("Machine Detail")
        machine_list = sorted(valid["SALES MODEL"].dropna().unique())
        selected_machine = st.selectbox("Select Machine", machine_list, key="machine_detail_selector")
        dfm = valid[valid["SALES MODEL"] == selected_machine].copy()

        st.write(f"### Machine: {selected_machine}")

        machine_type_summary = dfm.groupby("CCR TYPE").agg(
            Avg_Cost=("Adj Cost", "mean"),
            Avg_SMU=("SMU AT REBUILD", "mean"),
            Count=("Adj Cost", "count")
        ).reset_index()
        machine_type_summary = add_vs_cmr(machine_type_summary)

        st.write("### Average by Rebuild Type")
        display_table(machine_type_summary, currency_cols=["Avg_Cost"], percent_cols=["Vs CMR %"])

        st.write("### Machine SMU vs Cost")
        machine_chart = dfm[["SMU AT REBUILD", "Adj Cost", "CCR TYPE"]].dropna().rename(
            columns={"SMU AT REBUILD": "SMU", "Adj Cost": "Cost"}
        )
        st.scatter_chart(machine_chart, x="SMU", y="Cost", color="CCR TYPE")

        st.write("### Machine Raw Data Preview")
        display_table(
            dfm.head(50),
            currency_cols=["PARTS DN", "Base Rate", "Labor Cost", "Adj Cost"],
            percent_cols=[]
        )

    # =============================
    # DEALER SECTION
    # =============================
    with tab3:
        st.subheader("Dealer Analysis")
        dealer_machine_options = ["All Machines"] + sorted(valid["SALES MODEL"].dropna().unique().tolist())
        dealer_machine = st.selectbox("Focus Dealer Analysis on Machine", dealer_machine_options, key="dealer_machine_selector")

        dealer_df = valid.copy()
        if dealer_machine != "All Machines":
            dealer_df = dealer_df[dealer_df["SALES MODEL"] == dealer_machine]

        st.write(f"### Dealer Summary - {dealer_machine}")
        dealer_summary = dealer_df.groupby("DEALER").agg(
            Avg_Cost=("Adj Cost", "mean"),
            Avg_SMU=("SMU AT REBUILD", "mean"),
            Count=("Adj Cost", "count")
        ).reset_index().sort_values("Avg_Cost", ascending=False)
        dealer_summary = add_vs_machine_avg(dealer_summary)
        display_table(dealer_summary, currency_cols=["Avg_Cost"], percent_cols=["Vs Section Avg %"])

        st.write("### Dealer by Rebuild Type")
        dealer_type = dealer_df.groupby(["DEALER", "CCR TYPE"]).agg(
            Avg_Cost=("Adj Cost", "mean"),
            Avg_SMU=("SMU AT REBUILD", "mean"),
            Count=("Adj Cost", "count")
        ).reset_index().sort_values(["DEALER", "CCR TYPE"])
        display_table(dealer_type, currency_cols=["Avg_Cost"], percent_cols=[])

        st.write("### Dealer Average Cost Chart")
        dealer_chart = dealer_summary[["DEALER", "Avg_Cost"]].rename(columns={"Avg_Cost": "Average Cost"}).set_index("DEALER")
        st.bar_chart(dealer_chart)

    # =============================
    # REGION SECTION
    # =============================
    with tab4:
        st.subheader("Region Analysis")
        region_machine_options = ["All Machines"] + sorted(valid["SALES MODEL"].dropna().unique().tolist())
        region_machine = st.selectbox("Focus Region Analysis on Machine", region_machine_options, key="region_machine_selector")

        region_df = valid.copy()
        if region_machine != "All Machines":
            region_df = region_df[region_df["SALES MODEL"] == region_machine]

        st.write(f"### Region Summary - {region_machine}")
        region_summary = region_df.groupby("Region").agg(
            Avg_Cost=("Adj Cost", "mean"),
            Avg_SMU=("SMU AT REBUILD", "mean"),
            Count=("Adj Cost", "count")
        ).reset_index().sort_values("Avg_Cost", ascending=False)
        region_summary = add_vs_machine_avg(region_summary)
        display_table(region_summary, currency_cols=["Avg_Cost"], percent_cols=["Vs Section Avg %"])

        st.write("### Region by Rebuild Type")
        region_type = region_df.groupby(["Region", "CCR TYPE"]).agg(
            Avg_Cost=("Adj Cost", "mean"),
            Avg_SMU=("SMU AT REBUILD", "mean"),
            Count=("Adj Cost", "count")
        ).reset_index().sort_values(["Region", "CCR TYPE"])
        display_table(region_type, currency_cols=["Avg_Cost"], percent_cols=[])

        st.write("### Region Average Cost Chart")
        region_chart = region_summary[["Region", "Avg_Cost"]].rename(columns={"Avg_Cost": "Average Cost"}).set_index("Region")
        st.bar_chart(region_chart)

    # =============================
    # EXCEPTIONS
    # =============================
    with tab5:
        st.subheader("Exception Summary")
        exception_summary = pd.DataFrame({
            "Metric": ["Cost Outliers", "SMU 0 or 1"],
            "Value": [int(df["Outlier"].sum()), int(df["SMU AT REBUILD"].isin([0, 1]).sum())]
        })
        st.dataframe(exception_summary, use_container_width=True)

        st.write("### Outlier Rows")
        outlier_rows = df[df["Outlier"] == True]
        if not outlier_rows.empty:
            display_table(outlier_rows, currency_cols=["PARTS DN", "Base Rate", "Labor Cost", "Adj Cost"])
        else:
            st.write("No cost outliers detected.")

    # =============================
    # INSIGHTS
    # =============================
    with tab6:
        st.subheader("Auto Insights")
        if insights:
            for insight in insights:
                st.write("•", insight)
        else:
            st.write("No insights generated.")

    # =============================
    # EXPORT
    # =============================
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        summary.to_excel(writer, sheet_name="Summary", index=False)
        dealer_summary.to_excel(writer, sheet_name="Dealer Summary", index=False)
        dealer_type.to_excel(writer, sheet_name="Dealer By Type", index=False)
        region_summary.to_excel(writer, sheet_name="Region Summary", index=False)
        region_type.to_excel(writer, sheet_name="Region By Type", index=False)
        exception_summary.to_excel(writer, sheet_name="Exceptions", index=False)
        for m in valid["SALES MODEL"].dropna().unique():
            valid[valid["SALES MODEL"] == m].to_excel(writer, sheet_name=str(m)[:31], index=False)

    st.download_button(
        "Download Workbook",
        data=output.getvalue(),
        file_name="Rebuild_Analysis_Output.xlsx"
    )

else:
    st.info("Upload file and run analysis")

