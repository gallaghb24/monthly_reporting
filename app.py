import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="Monthly Artwork Analysis", layout="wide")
st.title("🎨 Monthly Artwork Versions Analysis")

st.markdown("Upload your Excel file below (with headers in row 2) — we’ll analyse amends, right-first-time rate, and show a category breakdown 👇")

uploaded_file = st.file_uploader("Upload your Excel file", type=["xlsx"])

if uploaded_file:
    try:
        xls = pd.ExcelFile(uploaded_file)
        df = pd.read_excel(xls, sheet_name="general_report", header=1)

        # Column name matching
        col_map = {col.strip().lower(): col for col in df.columns}
        if "client versions" not in col_map:
            st.error("❌ Couldn’t find a column called 'Client Versions'. Please check your file.")
        else:
            client_col = col_map["client versions"]

            # Step-by-step transformation
            df = df.sort_values(by=client_col, ascending=False)
            df = df.drop_duplicates(subset=["POS Code"], keep="first")
            df = df[df[client_col] != 0]
            df["Amends"] = df[client_col] - 1
            df["Right First Time"] = df[client_col].apply(lambda x: 1 if x == 1 else 0)
            df.loc[df["Project Description"].str.contains("ROI", na=False), "Category"] = "ROI"
            df.loc[df["Category"].isin(["Members", "Starbuys"]), "Category"] = "Main Event"
            df.loc[df["Category"].isin(["Loyalty / CRM", "Mobile"]), "Category"] = "Other"

            # Summary output table
            summary = pd.DataFrame()
            categories = sorted(df["Category"].dropna().unique())
            summary.loc["New Artwork Lines", categories] = df.groupby("Category").size()
            summary.loc["Amends", categories] = df.groupby("Category")["Amends"].sum()
            summary.loc["Right First Time", categories] = df.groupby("Category")["Right First Time"].sum()
            summary.loc["Average Round of Amends", categories] = df.groupby("Category")["Amends"].mean().round(2)

            # Display
            summary_display = summary.reset_index().rename(columns={"index": ""})
            st.subheader("📊 Category Breakdown Table")
            st.dataframe(summary_display, use_container_width=True)

            # Optional export
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                summary_display.to_excel(writer, index=False, sheet_name="Summary")

            st.download_button(
                label="📥 Download Summary Table",
                data=output.getvalue(),
                file_name="content_creation_summary.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:
        st.error("❌ There was an issue processing your file.")
        st.exception(e)
else:
    st.info("👆 Please upload a .xlsx file to get started.")
