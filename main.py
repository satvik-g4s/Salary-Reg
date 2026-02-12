import streamlit as st
import pandas as pd
from io import BytesIO
import gc

st.set_page_config(layout="wide")
st.title("Salary Register Processor")

uploaded_files = st.file_uploader(
    "Upload Salary .xlsb Files",
    type=["xlsb"],
    accept_multiple_files=True
)

if uploaded_files:

    output = BytesIO()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:

        progress = st.progress(0)

        for i, file in enumerate(uploaded_files):

            # 🔹 Read file
            df1 = pd.read_excel(
                file,
                sheet_name="REG",
                engine="pyxlsb"
            )

            # 🔹 Convert Excel serial date
            df1["PAYDATE"] = (
                pd.to_datetime("1899-12-30")
                + pd.to_timedelta(df1["PAYDATE"], unit="D")
            )

            month_year = df1["PAYDATE"].iloc[0].strftime("%b-%Y")

            # 🔹 Filter rows
            df1 = df1[
                df1["NATIONAL FESTIVAL HOLIDAY"] +
                df1["ARREAR NFH"] != 0
            ]

            # 🔹 Pivot
            keys = ["HUB NAME", "BRANCH NAME", "CLIENT CODE", "CLIENT NAME"]

            pivot_df = (
                df1.groupby(keys, as_index=False)
                   .agg({
                       "NATIONAL FESTIVAL HOLIDAY": "sum",
                       "ARREAR NFH": "sum"
                   })
            )

            pivot_df = pivot_df.sort_values(
                by=["HUB NAME", "BRANCH NAME", "CLIENT NAME"]
            )

            # 🔹 Write to output Excel (new sheets per month)
            df1.to_excel(writer, sheet_name=f"{month_year} Reg", index=False)
            pivot_df.to_excel(writer, sheet_name=f"{month_year} Pivot", index=False)

            # 🔹 Delete to free memory
            del df1
            del pivot_df
            gc.collect()

            progress.progress((i + 1) / len(uploaded_files))

    st.success("Processing Complete ✅")

    st.download_button(
        label="Download Final Excel File",
        data=output.getvalue(),
        file_name="Salary_Reg_Final.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
