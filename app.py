import streamlit as st
import pandas as pd

st.set_page_config(page_title="Monthly Artwork Analysis", layout="wide")
st.title("🎨 Monthly Artwork Versions Analysis")

st.markdown("Upload your Excel file below (with headers in row 2) — we'll handle the rest 👇")

uploaded_file = st.file_uploader("Upload your Excel file", type=["xlsx"])

if uploaded_file:
    try:
        xls = pd.ExcelFile(uploaded_file)
        # Read from second row (index 1), which has the headers
        df = pd.read_excel(xls, sheet_name="general_report", header=1)

        st.subheader("📋 Column Preview")
        st.write("Here are the detected column names:")
        st.write(list(df.columns))

        # Match actual column names safely
        col_map = {col.strip().lower(): col for col in df.columns}
        if "client versions" not in col_map:
            st.error("❌ Couldn’t find a column called 'Client Versions' in row 2. Please check your file.")
        else:
            client_col = col_map["client versions"]

            df = df.sort_values(by=client_col, ascending=False)
            df = df.drop_duplicates(subset=["POS Code"], keep="first")
            df = df[df[client_col] != 0]
            df["Amends"] = df[client_col] - 1
            df["Right First Time"] = df[client_col].apply(lambda x: 1 if x == 1 else 0)
            df.loc[df["Project Description"].str.contains("ROI", na=False), "Category"] = "ROI"
            df.loc[df["Category"].isin(["Members", "Starbuys"]), "Category"] = "Main Event"
            df.loc[df["Category"].isin(["Loyalty / CRM", "Mobile"]), "Category"] = "Other"

            # Metrics
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

            st.download_button("📥 Download Cleaned Data", df.to_excel(index=False), file_name="processed_artwork_data.xlsx")

    except Exception as e:
        st.error("❌ There was an issue processing your file.")
        st.exception(e)
else:
    st.info("👆 Please upload a .xlsx file to get started.")
