import streamlit as st
import pandas as pd
import io
import xlrd
import openpyxl

st.set_page_config(page_title="Excel Merger Tool", layout="centered")

st.title("📊 Excel Merger Tool")

st.write("""
यह tool multiple Excel files को merge करता है:
1. पहले 4 rows हटाए जाते हैं (header merge वाले).
2. Row no. 5 को header माना जाता है.
3. BU column के हिसाब से "TIME SOLAT" column auto भरता है → `APP-BAL-<BU>`.
4. सभी files merge होकर एक ही Excel बनती है.
""")

uploaded_files = st.file_uploader(
    "Select Excel Files (.xls / .xlsx)", type=["xls", "xlsx"], accept_multiple_files=True
)


# ---- Helper function to convert .xls → .xlsx ----
def convert_xls_to_xlsx(xls_bytes):
    book_xls = xlrd.open_workbook(file_contents=xls_bytes)
    book_xlsx = openpyxl.Workbook()
    sheet_xlsx = book_xlsx.active

    sheet_xls = book_xls.sheet_by_index(0)
    for row in range(sheet_xls.nrows):
        for col in range(sheet_xls.ncols):
            sheet_xlsx.cell(row=row+1, column=col+1).value = sheet_xls.cell_value(row, col)

    # Save to bytes buffer
    xlsx_bytes = io.BytesIO()
    book_xlsx.save(xlsx_bytes)
    xlsx_bytes.seek(0)
    return xlsx_bytes


if st.button("🚀 Merge Files"):
    if not uploaded_files:
        st.warning("Please upload Excel files first.")
    else:
        merged_df = None
        try:
            for i, file in enumerate(uploaded_files):
                file_bytes = file.read()
                
                # Convert old .xls to .xlsx
                if file.name.endswith(".xls"):
                    st.info(f"Converting {file.name} to .xlsx format...")
                    file_bytes = convert_xls_to_xlsx(file_bytes).read()
                
                # Read into pandas
                df = pd.read_excel(io.BytesIO(file_bytes), skiprows=4, dtype=str)

                if i == 0:
                    merged_df = df.copy()
                else:
                    df = df.iloc[1:].reset_index(drop=True)
                    merged_df = pd.concat([merged_df, df], ignore_index=True)

            if merged_df is not None:
                if "BU" in merged_df.columns and "TIME SOLAT" in merged_df.columns:
                    merged_df["TIME SOLAT"] = "APP-BAL-" + merged_df["BU"].astype(str)

                output_file = "Merged_Output.xlsx"
                merged_df.to_excel(output_file, index=False)

                with open(output_file, "rb") as f:
                    st.download_button(
                        label="📥 Download Merged Excel",
                        data=f,
                        file_name=output_file,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )

                st.success("✅ Merged file created successfully!")
        except Exception as e:
            st.error(f"Error: {e}")
