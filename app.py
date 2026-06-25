import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO

st.set_page_config(layout="wide")

st.title("Rebuild Analytics Platform V6")

# ---------- Inputs ----------
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
    options=['CMR', 'CPT+H', 'CPT-O'],
    default=['CMR', 'CPT+H', 'CPT-O']
)

run = st.button("Run Analysis")


# ---------- Insights ----------
def generate_insights(df, summary):
    insights = []

    try:
        if 'CMR' in summary['CCR TYPE'].values and 'CPT+H' in summary['CCR TYPE'].values:
            cmr = summary.loc[summary['CCR TYPE'] == 'CMR', 'Avg_Cost'].values[0]
            cpth = summary.loc[summary['CCR TYPE'] == 'CPT+H', 'Avg_Cost'].values[0]
            diff = ((cpth - cmr) / cmr) * 100 if cmr else 0

            insights.append(
                f"CPT+H average cost is {diff:.1f}% higher than CMR, indicating potential cost inefficiency"
            )
    except:
        pass

    try:
        high_var = df.groupby('CCR TYPE')['Adj Cost'].std().sort_values(ascending=False)
        if len(high_var) > 0:
            insights.append(f"Highest cost variability observed in {high_var.index[0]}")
    except:
        pass

    return insights


if run and rebuild_file:

    rebuild = pd.read_excel(rebuild_file, sheet_name=None)

    # ---------- Combine ----------
    frames = []
    for name, sheet in rebuild.items():
        sheet = sheet.copy()

        if 'CMR' in name:
            sheet['CCR TYPE'] = 'CMR'
        elif 'CPT+H' in name:
            sheet['CCR TYPE'] = 'CPT+H'
        else:
            sheet['CCR TYPE'] = 'CPT-O'

        frames.append(sheet)

    df = pd.concat(frames, ignore_index=True)

    if 'CCR TYPE' not in df.columns:
        df['CCR TYPE'] = 'UNKNOWN'

    # ---------- Filter by Rebuild Type ----------
    df = df[df['CCR TYPE'].isin(rebuild_filter)]

    # ---------- Dates ----------
    df['Service Date'] = df['DELIVERED DATE'].fillna(df['ENROLL DATE']).fillna(df['IN SERVICE DATE'])
    df['Service Year'] = pd.to_datetime(df['Service Date'], errors='coerce').dt.year

    df = df[(df['Service Year'] >= start_year) & (df['Service Year'] <= end_year)]

    # ---------- Machine Filter ----------
    if machine_input:
        machines = [m.strip() for m in machine_input.split(',')]
        df = df[df['SALES MODEL'].isin(machines)]

    # ---------- Rates ----------
    if use_default or not rate_file:
        years = list(range(2010, 2031))
        rate_df = pd.DataFrame({
            'Service Year': years,
            'Rate': [115 + (y - 2016) * 3 for y in years]
        })

        df = df.merge(rate_df, on='Service Year', how='left')

    else:
        rates = pd.read_excel(rate_file, sheet_name=None)
        rows = []
        for k, v in rates.items():
            if k != "Summary":
                temp = v.copy()
                temp = temp.rename(columns={'Year': 'Service Year', 'Base Rate (USD)': 'Rate'})
                rows.append(temp[['Service Year', 'Rate']])
        rate_df = pd.concat(rows)

        df = df.merge(rate_df, on='Service Year', how='left')

    df['Base Rate'] = df['Rate'].fillna(df['Rate'].mean())

    # ---------- Cost ----------
    df['Labor Cost'] = df['REBUILD WORK HRS'] * df['Base Rate']

    df['Infl'] = df['Service Year'].apply(
        lambda y: 1 + (base_year - y) * 0.03 if pd.notnull(y) else 1
    )

    df['Adj Cost'] = df['PARTS DN'] * df['Infl'] + df['Labor Cost'] * df['Infl']

    df = df[df['Adj Cost'] > 0]

    # ---------- Outlier Detection ----------
    df['Outlier'] = False

    for (m, t), g in df.groupby(['SALES MODEL', 'CCR TYPE']):
        idx = g.index

        if len(g) >= 5:
            log_vals = np.log(g['Adj Cost'])
            q1, q3 = np.percentile(log_vals, [25, 75])
            iqr = q3 - q1
            mult = 2 if len(g) < 8 else 1.5
            low, high = q1 - mult * iqr, q3 + mult * iqr

            mask = (log_vals < low) | (log_vals > high)
            df.loc[idx, 'Outlier'] = mask

    valid = df[df['Outlier'] == False].copy()

    # ---------- Summary ----------
    summary = valid.groupby('CCR TYPE', dropna=False).agg(
        Avg_Cost=('Adj Cost', 'mean'),
        Count=('Adj Cost', 'count')
    ).reset_index()

    insights = generate_insights(valid, summary)

    # ---------- Tabs ----------
    tab1, tab2, tab3, tab4 = st.tabs(
        ["Dashboard", "Machine Detail", "Exceptions", "Insights"]
    )

    # ---------- Dashboard ----------
    with tab1:

        col1, col2, col3, col4 = st.columns(4)

        col1.metric("Total Rows", len(valid))
        col2.metric("Outliers", int(df['Outlier'].sum()))
        col3.metric("Avg Cost", f"${int(valid['Adj Cost'].mean()):,}" if len(valid) else "$0")

        if 'CMR' in summary['CCR TYPE'].values and 'CPT+H' in summary['CCR TYPE'].values:
            cmr = summary.loc[summary['CCR TYPE']=='CMR','Avg_Cost'].values[0]
            cpth = summary.loc[summary['CCR TYPE']=='CPT+H','Avg_Cost'].values[0]
            delta = ((cpth - cmr)/cmr)*100 if cmr else 0
            col4.metric("CPT+H vs CMR (%)", f"{delta:.1f}%")

        st.dataframe(summary)

        # ✅ Combined Scatter
        st.subheader("SMU vs Cost (All Rebuild Types)")

        chart_data = valid[['SMU AT REBUILD', 'Adj Cost', 'CCR TYPE']].dropna()
        chart_data = chart_data.rename(columns={
            'SMU AT REBUILD': 'SMU',
            'Adj Cost': 'Cost'
        })

        st.scatter_chart(
            chart_data,
            x='SMU',
            y='Cost',
            color='CCR TYPE'
        )

    # ---------- Machine Detail ----------
    with tab2:

        machines = valid['SALES MODEL'].dropna().unique()

        if len(machines) > 0:
            selected = st.selectbox("Select Machine", machines)

            dfm = valid[valid['SALES MODEL'] == selected]

            st.dataframe(dfm.head(50))

            st.subheader("Machine Chart")

            chart_df = dfm[['SMU AT REBUILD', 'Adj Cost', 'CCR TYPE']].rename(
                columns={'SMU AT REBUILD': 'SMU', 'Adj Cost': 'Cost'}
            )

            st.scatter_chart(chart_df, x='SMU', y='Cost', color='CCR TYPE')

        else:
            st.warning("No machine data")

    # ---------- Exceptions ----------
    with tab3:
        st.write("Outliers:", int(df['Outlier'].sum()))
        st.write("SMU 0/1:", int(df['SMU AT REBUILD'].isin([0, 1]).sum()))

    # ---------- Insights ----------
    with tab4:
        st.subheader("Auto Insights")
        for i in insights:
            st.write("•", i)

    # ---------- Export ----------
    output = BytesIO()

    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        summary.to_excel(writer, sheet_name='Summary', index=False)

        for m in valid['SALES MODEL'].dropna().unique():
            valid[valid['SALES MODEL'] == m].to_excel(
                writer,
                sheet_name=str(m)[:31],
                index=False
            )

    st.download_button(
        "Download Workbook",
        data=output.getvalue(),
        file_name="Rebuild_Analysis_Output.xlsx"
    )

else:
    st.info("Upload file and run analysis")
