# Rebuild Analytics Platform

**Current app version:** V18.5 Final QA & Documentation Polish  
**Primary downstream output:** Power BI Dataset Export  
**Methodology version:** 2026.07-PowerBI-CrossType-Outlier-v18

## Purpose

The Rebuild Analytics Platform is a Streamlit application for analyzing certified rebuild cost performance by machine, rebuild type, dealer, region, and service year. The app normalizes rebuild costs to USD, calculates labor using dealer-year labor rates, applies inflation adjustment, detects cost outliers, treats qualifying CPT+H cross-type exceptions as outliers, and produces a clean Power BI-ready export.

The app is designed to support both analysis review and Power BI handoff. The Power BI Dataset Export is the primary output and should be used as the source file for Power BI Desktop or a Power BI Service refresh workflow.

## Required repository files

The deployed repository should include:

```text
app.py
requirements.txt
expanded_dealer_base_rate_by_year_2016_2026.xlsx
```

Optional files:

```text
cat_logo.png
golden_test_rebuild_workbook.xlsx
```

The expanded dealer-rate workbook is required for the preferred built-in dealer-rate workflow. If that workbook is missing, the app can fall back to emergency generic rates, but output values may change.

## App workflow

1. Upload the rebuild source workbook.
2. Confirm dealer labor-rate source.
3. Optionally upload custom machine grouping.
4. Configure machine, year, rebuild type, region, currency, inflation, and governance settings.
5. Run the analysis.
6. Review Executive Dashboard, Machine Detail, Dealer Performance, Region Performance, Exceptions & Data Quality, Power BI Readiness, and Final Handoff.
7. Download the Power BI Dataset Export.
8. Save or rename the exported file as:

```text
Rebuild_Analytics_PowerBI_Dataset.xlsx
```

9. Replace the existing file in the approved SharePoint or OneDrive Power BI source folder.
10. Refresh the Power BI report.

## Power BI source-file rule

For the Level 2 Power BI pipeline, keep the final source file name and folder path stable:

```text
Rebuild_Analytics_PowerBI_Dataset.xlsx
```

Do not use changing source names such as:

```text
Rebuild_Analytics_PowerBI_Dataset_v2.xlsx
Rebuild_Analytics_PowerBI_Dataset_Final.xlsx
Rebuild_Analytics_PowerBI_Dataset_2026-07-20.xlsx
```

Archive dated scenario copies somewhere else. The Power BI source workbook should be replaced in place.

## Methodology summary

### Cost basis

```text
Adjusted Total Cost USD = PARTS DN USD + Labor Cost USD
Labor Cost USD = REBUILD WORK HRS x dealer service-year labor rate
```

PLUS PARTS DN is ignored.

### Currency

Source currency is converted to USD before inflation and analysis.

### Inflation

The app applies BLS CPI-U All Items, U.S. city average, not seasonally adjusted. Inflation is applied by component when enabled.

### Statistical cost outliers

Cost outliers are detected by Machine + CCR TYPE using log(cost) IQR logic:

```text
Groups with fewer than 5 records: no statistical outlier removal
Groups with 5 to 7 records: 2.0 x IQR
Groups with 8 or more records: 1.5 x IQR
```

SMU values of 0 or 1 are data-quality flags only and are not statistical outlier criteria.

### Cross-type CPT+H outliers

CPT+H rows above the machine-level valid CMR median cost x 1.10 are treated as outliers when at least 3 valid CMR rows exist for the machine.

### CPT+H reporting

The app reports one CPT+H value after excluding statistical cost outliers and cross-type CPT+H outliers. Excluded rows remain available in audit tables.

## Power BI export structure

The Power BI Dataset Export includes locked-schema fact, dimension, audit, relationship, documentation, and handoff tables. Every exported worksheet has headers on row 1.

Key tables include:

```text
Fact_Rebuild_Rows
Fact_Valid_Rebuild_Rows
Fact_Global_RebuildType_AvgCost
Fact_Machine_RebuildType_AvgCost
Fact_MachineRegion_RebuildType_AvgCost
Fact_MachineRegionYear_RebuildType_AvgCost
Fact_Dealer_Performance
Fact_DealerYear_Performance
Fact_Region_Performance
Fact_Exception_Rows
Fact_Outlier_Rows
Fact_CrossType_Outliers
Dim_Machine
Dim_Machine_Group
Dim_Dealer
Dim_Rebuild_Type
Dim_Region
Dim_Service_Year
PowerBI_Relationship_Checks
PowerBI_Schema_Status
PowerBI_Export_Health
PowerBI_Visual_Coverage_Matrix
PowerBI_Table_Dictionary
PowerBI_Modeling_Notes
Handoff_Checklist
Sample_Workflow
Version_History
Methodology_Snapshot
```

Some logical table names are shortened at the Excel worksheet level due to Excel's 31-character sheet-name limit. Use `PowerBI_Sheet_Name_Map` to reconcile worksheet names with logical names.

## Recommended Power BI modeling rules

Use a clean star-schema style model:

```text
Dimension tables filter fact tables
Cardinality: One-to-many
Cross filter direction: Single
Relationship active: Yes
```

Avoid:

```text
Many-to-many relationships
Bidirectional relationships
Fact-to-fact relationships
Power BI auto-generated ambiguous relationships
```

For visuals that need to respond to Service Year, prefer year-aware tables such as:

```text
Fact_MachineYear_RebuildType_AvgCost
Fact_RegionYear_RebuildType_AvgCost
Fact_DealerYear_Performance
Fact_MachineRegionYear_RebuildType_AvgCost
```

For fully dynamic slicer-responsive visuals, use:

```text
Fact_Rebuild_Rows
```

## Golden test workbook

The repository can include:

```text
golden_test_rebuild_workbook.xlsx
```

Use this workbook to quickly verify that the app launches, reads the expected sheet structure, runs analysis, detects cross-type CPT+H outlier behavior, exports Power BI tables, and preserves row 1 headers.

Suggested golden test process:

1. Upload `golden_test_rebuild_workbook.xlsx`.
2. Use built-in dealer rates.
3. Run analysis for all machines and all years.
4. Confirm the app produces Power BI Readiness and Final Handoff outputs.
5. Export the Power BI Dataset Export.
6. Confirm `PowerBI_Schema_Status` and `PowerBI_Export_Health` are available.
7. Open the exported workbook and confirm every sheet has row 1 headers.
8. Refresh the Power BI report after replacing the fixed source file.

## Version history

| Version | Theme | Key change |
|---|---|---|
| V18.0 | Power BI core export and cross-type outlier methodology | Power BI became the primary export focus; cross-type CPT+H exceptions became outliers; one CPT+H value reported. |
| V18.1 | Power BI model handoff | Added relationship checks, sheet-name map, machine group dimension, pipeline guide, build checklist, expanded DAX starter, and layout guidance. |
| V18.2 | Visual coverage matrix | Added visual coverage matrix mapping app outputs to Power BI visuals. |
| V18.3 | Power BI reliability and year-aware export | Added locked schemas, export health, and year-aware aggregate tables. |
| V18.4 | Final handoff and Power BI stabilization | Added schema status, methodology snapshot, modeling notes, handoff checklist, sample workflow, and version history. |
| V18.5 | Final QA and documentation polish | Added README support, golden test workbook guidance, and stronger Excel export header enforcement. |

## Deployment notes

For local testing:

```bash
pip install -r requirements.txt
streamlit run app.py
```

For GitHub deployment, upload or replace:

```text
app.py
requirements.txt
expanded_dealer_base_rate_by_year_2016_2026.xlsx
```

If using the golden test workbook, also upload:

```text
golden_test_rebuild_workbook.xlsx
```

## Handoff checklist

Before handoff, confirm:

```text
[ ] app.py is current
[ ] requirements.txt is current
[ ] expanded dealer-rate workbook is present
[ ] golden test workbook runs successfully
[ ] Power BI Dataset Export produces row 1 headers
[ ] PowerBI_Schema_Status is available
[ ] PowerBI_Export_Health is available
[ ] Power BI source file is named Rebuild_Analytics_PowerBI_Dataset.xlsx
[ ] Power BI refresh has been tested
```

