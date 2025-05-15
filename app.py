import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="Monthly Artwork Analysis", layout="wide")
st.title("🎨 Monthly Artwork Versions Analysis")

st.markdown("Upload your Excel file below (with headers in row 2) — we’ll analyse amends, right-first-time rate, and more 👇")

uploaded_file = st.file_uploader("Upload your Excel file", type=["xlsx"])

if uploaded_file:
    try:
        xls = pd.ExcelFile(uploaded_file)
        df = pd.read_excel(xls, sheet_name="general_report", header=1)

        # Column name safety
        col_map = {col.strip().lower(): col for col in df.columns}
        if "client versions" not in col_map:
            st.error("❌ Couldn’t find a column called 'Client Versions' in row 2. Please check your file.")
        else:
            client_col = col_map["client versions"]

            # Data transformations
            df = df.sort_values(by=client_col, ascending=False)
            df = df.drop_duplicates(subset=["POS Code"], keep="first")
            df = df[df[client_col] != 0]
            df["Amends"] = df[client_col] - 1
            df["Right First Time"] = df[client_col].apply(lambda x: 1 if x == 1 else 0)
            df.loc[df["Project Description"].str.contains("ROI", na=False), "Category"] = "ROI"
            df.loc[df["Category"].isin(["Members", "Starbuys"]), "Category"] = "Main Event"
            df.loc[df["Category"].isin(["Loyalty / CRM", "Mobile"]), "Category"] = "Other"

            # Stats
            num_new_artworks = len(df)
            total_amends = df["Amends"].sum()
            num_rft = df["Right First Time"].sum()
            rft_percentage = round((num_rft / num_new_artworks) * 100, 2)
            avg_amends = round(df["Amends"].mean(), 2)

            # Display
            st.subheader("📊 Key Stats")
            col1, col2, col3 = st.columns(3)
            col1.metric("New Artworks", num_new_artworks)
            col2.metric("Total Amends", total_amends)
            col3.metric("Right First Time", f"{rft_percentage}% ({num_rft})")

            st.metric("Average Amend Rate", avg_amends)

            with st.expander("🔍 View Processed Data"):
                st.dataframe(df, use_container_width=True)

            # Prepare Excel download
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df.to_excel(writer, index=False)
            st.download_button(
                label="📥 Download Cleaned Data",
                data=output.getvalue(),
                file_name="processed_artwork_data.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:
        st.error("❌ There was an issue processing your file.")
        st.exception(e)
else:
    st.info("👆 Please upload a .xlsx file to get started.")
