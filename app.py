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

run = st.button("Run Analysis")

# ---------- Insights Function ----------
def generate_insights(df, summary):
    insights = []
    
    try:
        if ('CMR' in summary['CCR TYPE'].values) and ('CPT+H' in summary['CCR TYPE'].values):
            cmr = summary.loc[summary['CCR TYPE'] == 'CMR', 'Avg_Cost'].values[0]
            cpth = summary.loc[summary['CCR TYPE'] == 'CPT+H', 'Avg_Cost'].values[0]
            diff = ((cpth - cmr) / cmr) * 100 if cmr else 0
            insights.append(f"CPT+H is {diff:.1f}% vs CMR baseline")
    except:
        pass

    try:
        high_var = df.groupby('CCR TYPE')['Adj Cost'].std().sort_values(ascending=False)
        if len(high_var) > 0:
            insights.append(f"Highest cost variability in {high_var.index[0]}")
    except:
        pass

    return insights


# ---------- Run ----------
if run and rebuild_file:

    rebuild = pd.read_excel(rebuild_file, sheet_name=None)

    # ---- Rates ----
    if use_default or not rate_file:
        years = list(range(2010, 2031))
        rate_df = pd.DataFrame({
            'Dealer Code': ['DEFAULT'] * len(years),
            'Service Year': years,
            'Rate': [115 + (y - 2016) * 3 for y in years]
        })
    else:
        rates = pd.read_excel(rate_file, sheet_name=None)
        rows = []
        for k, v in rates.items():
            if k != "Summary":
                temp = v.copy()
                temp['Dealer Code'] = k.split('_')[0]
                temp = temp.rename(columns={
                    'Year': 'Service Year',
                    'Base Rate (USD)': 'Rate'
                })
                rows.append(temp[['Dealer Code', 'Service Year', 'Rate']])
        rate_df = pd.concat(rows)

    # ---- Combine ----
    frames = []
    for name, df in rebuild.items():
        if 'CMR' in name:
            df['CCR TYPE'] = 'CMR'
        elif 'CPT+H' in name:
            df['CCR TYPE'] = 'CPT+H'
        else:
            df['CCR TYPE'] = 'CPT-O'
        frames.append(df)

    df = pd.concat(frames, ignore_index=True)

    # ---- Core Logic ----
    df['Service Date'] = df['DELIVERED DATE'].fillna(
        df['ENROLL DATE']).fillna(df['IN SERVICE DATE'])
    df['Service Year'] = pd.to_datetime(
        df['Service Date'], errors='coerce').dt.year

    df = df[(df['Service Year'] >= start_year) &
            (df['Service Year'] <= end_year)]

    if machine_input:
        machines = [m.strip() for m in machine_input.split(',')]
        df = df[df['SALES MODEL'].isin(machines)]

    df['Dealer Code'] = df['DEALER'].str.extract(r'([A-Z]\d{3})')
    df = df.merge(rate_df, how='left',
                  on=['Dealer Code', 'Service Year'])

    df['Base Rate'] = df['Rate'].fillna(rate_df['Rate'].mean())

    # ---- Cost ----
    df['Labor Cost'] = df['REBUILD WORK HRS'] * df['Base Rate']
    df['Infl'] = df['Service Year'].apply(
        lambda y: 1 + (base_year - y) * 0.03)
    df['Adj Cost'] = df['PARTS DN'] * df['Infl'] + \
        df['Labor Cost'] * df['Infl']

    # ---- Outliers ----
    df['log'] = np.log(df['Adj Cost'])

    def detect(g):
        if len(g) < 5:
            g['Outlier'] = False
        else:
            q1, q3 = g['log'].quantile([.25, .75])
            iqr = q3 - q1
            m = 2 if len(g) < 8 else 1.5
            low, high = q1 - m * iqr, q3 + m * iqr
            g['Outlier'] = (g['log'] < low) | (g['log'] > high)
        return g

    df = df.groupby(['SALES MODEL', 'CCR TYPE'],
                    group_keys=False).apply(detect)

    valid = df[~df['Outlier']]

    summary = valid.groupby('CCR TYPE').agg(
        Avg_Cost=('Adj Cost', 'mean'),
        Count=('Adj Cost', 'count')
    ).reset_index()

    insights = generate_insights(valid, summary)

    # ---------- Tabs ----------
    tab1, tab2, tab3, tab4 = st.tabs(
        ["Dashboard", "Machine Detail", "Exceptions", "Insights"])

    # ---------- Dashboard ----------
    with tab1:
        col1, col2, col3 = st.columns(3)

        col1.metric("Total Rows", len(valid))
        col2.metric("Outliers", int(df['Outlier'].sum()))
        col3.metric("Avg Cost", f"${int(valid['Adj Cost'].mean()):,}")

        st.dataframe(summary)

        st.subheader("SMU vs Cost")

        for t in valid['CCR TYPE'].dropna().unique():
            subset = valid[valid['CCR TYPE'] == t]
            st.write(f"### {t}")

            chart_df = subset[['SMU AT REBUILD', 'Adj Cost']] \
                .rename(columns={
                    'SMU AT REBUILD': 'SMU',
                    'Adj Cost': 'Cost'
                })

            st.scatter_chart(chart_df)

    # ---------- Machine ----------
    with tab2:
        machines = valid['SALES MODEL'].unique()
        selected = st.selectbox("Select Machine", machines)

        dfm = valid[valid['SALES MODEL'] == selected]
        st.dataframe(dfm.head(50))

        st.subheader("Machine Chart")

        for t in dfm['CCR TYPE'].unique():
            subset = dfm[dfm['CCR TYPE'] == t]
            st.write(f"### {t}")

            chart_df = subset[['SMU AT REBUILD', 'Adj Cost']] \
                .rename(columns={
                    'SMU AT REBUILD': 'SMU',
                    'Adj Cost': 'Cost'
                })

            st.scatter_chart(chart_df)

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

        for m in valid['SALES MODEL'].unique():
            valid[valid['SALES MODEL'] == m].to_excel(
                writer, sheet_name=str(m)[:31], index=False)

    st.download_button(
        "Download Workbook",
        data=output.getvalue(),
        file_name="Rebuild_Analysis_Output.xlsx"
    )

else:
    st.info("Upload file and run analysis")
