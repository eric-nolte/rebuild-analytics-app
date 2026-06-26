
"""
V15.7 PATCH: Phase 1 Security + V15.5 Excel Row Highlighting
------------------------------------------------------------
Purpose:
- Merge Caterpillar: Confidential Yellow security controls with the V15.5 Excel row-highlighting patch.
- Paste this patch into app.py after imports and before workbook export sections.
- Replace your existing apply_excel_brand_formatting(workbook) function with the version below.

Phase 1 security controls included:
1. Confidential Yellow banner/label support.
2. .xlsx-only upload validation helper.
3. File-size limit helper.
4. Clear-session helper.
5. Excel formula-injection sanitization before DataFrame.to_excel exports.
6. Export authorization acknowledgement helper.
7. Excel export formatting with:
   - Red highlight for outlier rows.
   - Orange highlight for cross-type exception rows.
   - Yellow highlight for insufficient-sample rows.

Recommended constants:
APP_VERSION = "V15.7"
CONFIDENTIALITY_LABEL = "Caterpillar: Confidential Yellow"
MAX_UPLOAD_MB = 50
"""

import pandas as pd
import streamlit as st
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter


CONFIDENTIALITY_LABEL = "Caterpillar: Confidential Yellow"
MAX_UPLOAD_MB = 50


# =====================================================
# PHASE 1 SECURITY HELPERS
# =====================================================
def render_confidential_yellow_banner():
    """Display a persistent Confidential Yellow warning banner in the app."""
    st.warning(
        f"{CONFIDENTIALITY_LABEL} — Authorized use only. "
        "Do not upload, analyze, or export data unless approved to handle this information."
    )


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
    """
    Enforce Phase 1 upload checks.
    Use immediately after st.file_uploader returns a file.
    """
    if uploaded_file is None:
        return True

    if not uploaded_file.name.lower().endswith(".xlsx"):
        st.error("Only .xlsx files are allowed for security and compatibility reasons.")
        st.stop()

    uploaded_size_mb = file_size_mb(uploaded_file)
    if uploaded_size_mb is not None and uploaded_size_mb > max_upload_mb:
        st.error(
            f"Uploaded file is {uploaded_size_mb:.1f} MB. "
            f"Maximum allowed size is {max_upload_mb} MB."
        )
        st.stop()

    return True


def reset_app_state():
    """Clear analysis state from Streamlit session."""
    st.session_state.run_clicked = False
    st.session_state.analysis = None


def render_clear_session_button():
    """Render a button that clears analysis/session state."""
    if st.button("Clear uploaded data / reset analysis"):
        reset_app_state()
        st.success("Session analysis state cleared. Refresh the page if you also want to clear the file uploader selection.")
        st.stop()


def sanitize_excel_value(value):
    """
    Prevent Excel formula injection in exported workbooks.
    If a text value begins with =, +, -, or @, prefix with apostrophe.
    """
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


# Monkeypatch pandas DataFrame.to_excel once so existing export code is automatically protected.
# If app.py already has this monkeypatch, do not duplicate it.
if not hasattr(pd.DataFrame, "_cat_original_to_excel"):
    pd.DataFrame._cat_original_to_excel = pd.DataFrame.to_excel

    def _secure_to_excel(self, *args, **kwargs):
        return pd.DataFrame._cat_original_to_excel(sanitize_excel_df(self), *args, **kwargs)

    pd.DataFrame.to_excel = _secure_to_excel


def render_export_acknowledgement(key="export_ack"):
    """
    Returns True only when the user acknowledges export authorization.
    Use before displaying any download button.
    """
    return st.checkbox(
        f"I understand this export is {CONFIDENTIALITY_LABEL} and I am authorized to download it.",
        key=key,
    )


# =====================================================
# V15.5 + SECURITY EXCEL FORMATTING / HIGHLIGHTING
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

    Priority order:
    1. Outlier red
    2. Cross-type orange
    3. Insufficient-sample yellow
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

        header_map = {
            str(cell.value).upper().strip() if cell.value is not None else "": cell.column
            for cell in ws[1]
        }

        outlier_col = header_map.get("OUTLIER")
        cross_type_col = header_map.get("CROSS-TYPE EXCEPTION FLAG")
        insufficient_col = header_map.get("INSUFFICIENT SAMPLE GROUP")

        for col_idx in range(1, ws.max_column + 1):
            header_cell = ws.cell(row=1, column=col_idx)
            header = str(header_cell.value).upper() if header_cell.value is not None else ""
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

        # Ensure header style wins over row highlight operations.
        for cell in ws[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center")


# =====================================================
# INTEGRATION NOTES
# =====================================================
"""
Recommended integration into app.py:

1. After render_header(), add:

    render_confidential_yellow_banner()

2. After file uploader:

    rebuild_file = st.file_uploader("Upload Rebuild File", type=["xlsx"])
    render_clear_session_button()
    validate_uploaded_file_security(rebuild_file)

3. Before any download button:

    if render_export_acknowledgement("full_export_ack"):
        st.download_button(...)
    else:
        st.info("Confirm export authorization to enable the workbook download.")

4. Keep all existing dataframe.to_excel(...) calls. The monkeypatch above sanitizes exported cells.

5. Keep existing apply_excel_brand_formatting(writer.book) calls.
   This merged function applies brand formatting and row highlights.
"""

