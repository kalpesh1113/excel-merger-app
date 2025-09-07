import streamlit as st
import pandas as pd
import io
import xlrd
import openpyxl
from openpyxl.styles import Border, Side


# ---- Helper function to add borders to all cells ----
def add_borders_to_excel(file_path):
    wb = openpyxl.load_workbook(file_path)
    ws = wb.active

    thin = Side(border_style="thin", color="000000")
    border = Border(top=thin, left=thin, right=thin, bottom=thin)

    for row in ws.iter_rows():
        for cell in row:
            cell.border = border

    wb.save(file_path)


# ---- Your existing merge logic ----
if merged_df is not None:
    # Normalize column names
    merged_df.columns = merged_df.columns.str.strip().str.upper()

    # BU → TIME SLOT mapping
    if "BU" in merged_df.columns and "TIME SLOT" in merged_df.columns:
        merged_df["TIME SLOT"] = "APP-BAL-" + merged_df["BU"].astype(str)
    else:
        st.warning("⚠️ 'BU' or 'TIME SLOT' column not found. Columns found: " + str(list(merged_df.columns)))

    output_file = "Merged_Output.xlsx"
    merged_df.to_excel(output_file, index=False)

    # ✅ Add all borders
    add_borders_to_excel(output_file)

    # Download button
    with open(output_file, "rb") as f:
        st.download_button(
            label="📥 Download Merged Excel",
            data=f,
            file_name=output_file,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    st.success("✅ Merged file created successfully with borders!")
