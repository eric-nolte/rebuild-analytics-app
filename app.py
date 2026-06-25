import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO

st.set_page_config(layout="wide")

st.title("Rebuild Analytics Platform V6+")

# ---------- SESSION STATE FIX ----------
if "run_clicked" not in st.session_state:
    st.session_state.run_clicked = False

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

if st.button("Run Analysis"):
    st.session_state.run_clicked = True

# ---------- Run ----------
if st.session_state.run_clicked and rebuild_file:

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

    df = df[df['CCR TYPE'].isin(rebuild_filter)]

    # ---------- Dates ----------
    df['Service Date'] = df['DELIVERED DATE'].fillna(df['ENROLL DATE']).fillna(df['IN SERVICE DATE'])
    df['Service Year'] = pd.to_datetime(df['Service Date'], errors='coerce').dt.year

    df = df[(df['Service Year'] >= start_year) & (df['Service Year'] <= end_year)]

    # ---------- Machine Filter ----------
    if machine_input:
        machines = [m.strip() for m in machine_input.split(',')]
        df = df[df['SALES MODEL'].isin(machines)]

    # ---------- Dealer + Region Extraction ----------
    df['Dealer Code'] = df['DEALER'].astype(str).str.extract(r'([A-Z]\d{3})')

    # Simple region logic (customize later if needed)
    df['Region'] = df['Dealer Code'].str[0]

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

    # ---------- Outliers ----------
    df['Outlier'] = False

    for (m, t), g in df.groupby(['SALES MODEL', 'CCR TYPE']):
        if len(g) >= 5:
            log_vals = np.log(g['Adj Cost'])
            q1, q3 = np.percentile(log_vals, [25, 75])
            iqr = q3 - q1
            mult = 2 if len(g) < 8 else 1.5
            low, high = q1 - mult * iqr, q3 + mult * iqr

            mask = (log_vals < low) | (log_vals > high)
            df.loc[g.index, 'Outlier'] = mask

    valid = df[df['Outlier'] == False].copy()

    # ---------- Summary ----------
    summary = valid.groupby('CCR TYPE').agg(
        Avg_Cost=('Adj Cost', 'mean'),
        Count=('Adj Cost', 'count')
    ).reset_index()

    # ---------- Tabs ----------
    tab1, tab2, tab3, tab4 = st.tabs(
        ["Dashboard", "Machine Detail", "Dealer/Region", "Exceptions"]
    )

    # ---------- Dashboard ----------
    with tab1:
        st.subheader("Executive Summary")
        st.dataframe(summary)

        chart_data = valid[['SMU AT REBUILD', 'Adj Cost', 'CCR TYPE']].rename(
            columns={'SMU AT REBUILD': 'SMU', 'Adj Cost': 'Cost'}
        )

        st.scatter_chart(chart_data, x="SMU", y="Cost", color="CCR TYPE")

    # ---------- Machine ----------
    with tab2:

        machines = valid['SALES MODEL'].dropna().unique()

        selected = st.selectbox("Select Machine", machines)

        dfm = valid[valid['SALES MODEL'] == selected]

        st.subheader(f"Machine: {selected}")

        # ✅ AVERAGE BY TYPE (WHAT YOU WANTED)
        type_summary = dfm.groupby('CCR TYPE').agg(
            Avg_Cost=('Adj Cost', 'mean'),
            Count=('Adj Cost', 'count')
        ).reset_index()

        st.write("### Average Cost by Rebuild Type")
        st.dataframe(type_summary)

        st.write("### Raw Data")
        st.dataframe(dfm.head(50))

    # ---------- Dealer / Region ----------
    with tab3:

        st.subheader("Dealer Performance")

        dealer_summary = valid.groupby('Dealer Code').agg(
            Avg_Cost=('Adj Cost', 'mean'),
            Count=('Adj Cost', 'count')
        ).reset_index()

        st.dataframe(dealer_summary)

        st.subheader("Dealer + Rebuild Type")

        dealer_type = valid.groupby(['Dealer Code', 'CCR TYPE']).agg(
            Avg_Cost=('Adj Cost', 'mean')
        ).reset_index()

        st.dataframe(dealer_type)

        st.subheader("Region Summary")

        region_summary = valid.groupby('Region').agg(
            Avg_Cost=('Adj Cost', 'mean'),
            Count=('Adj Cost', 'count')
        ).reset_index()

        st.dataframe(region_summary)

    # ---------- Exceptions ----------
    with tab4:
        st.write("Outliers:", int(df['Outlier'].sum()))
        st.write("SMU 0/1:", int(df['SMU AT REBUILD'].isin([0, 1]).sum()))

    # ---------- Export ----------
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        summary.to_excel(writer, sheet_name='Summary', index=False)

        for m in valid['SALES MODEL'].dropna().unique():
            valid[valid['SALES MODEL'] == m].to_excel(
                writer, sheet_name=str(m)[:31], index=False
            )

    st.download_button(
        "Download Workbook",
        data=output.getvalue(),
        file_name="Rebuild_Analysis_Output.xlsx"
    )

else:
    st.info("Upload file and run analysis")
``
