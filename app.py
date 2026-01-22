import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="Excel Sheet Merger", layout="centered")

st.title("📊 Excel Multi-Sheet Row Merger by S M Baqir")
st.write(
    "Upload an Excel file, select a **starting row**, and this app will "
    "merge data from all sheets into a single sheet."
)

# Upload Excel file
uploaded_file = st.file_uploader(
    "Upload Excel Workbook", type=["xlsx"]
)

# Row input
start_row = st.number_input(
    "Pick data from row number (same for all sheets)",
    min_value=1,
    value=21,
    step=1
)

if uploaded_file is not None:
    try:
        excel_file = pd.ExcelFile(uploaded_file)

        st.success(f"Found {len(excel_file.sheet_names)} sheets:")
        st.write(excel_file.sheet_names)

        merged_data = []

        for sheet in excel_file.sheet_names:
            df = pd.read_excel(
                excel_file,
                sheet_name=sheet,
                header=None
            )

            # Convert Excel row number to zero-based index
            extracted_df = df.iloc[start_row - 1:].copy()
            extracted_df["Source_Sheet"] = sheet

            merged_data.append(extracted_df)

        final_df = pd.concat(merged_data, ignore_index=True)

        st.subheader("🔍 Preview of Merged Data")
        st.dataframe(final_df.head(20))

        # Save to new Excel file
        output = BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            final_df.to_excel(
                writer,
                sheet_name="Merged_Data",
                index=False,
                header=False
            )

        st.download_button(
            label="⬇️ Download Merged Excel File",
            data=output.getvalue(),
            file_name="Merged_Output.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"Error processing file: {e}")
