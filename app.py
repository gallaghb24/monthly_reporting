import streamlit as st
import pandas as pd

st.set_page_config(page_title="Monthly Artwork Analysis", layout="wide")
st.title("🎨 Monthly Artwork Versions Analysis")

st.markdown("Upload your Excel file below and we'll do the rest — amends, right-first-time stats, and all that good stuff 👇")

uploaded_file = st.file_uploader("Upload your Excel file", type=["xlsx"])

if uploaded_file:
    try:
        xls = pd.ExcelFile(uploaded_file)
        sheet_names = xls.sheet_names
        st.success(f"✅ File uploaded. Found {len(sheet_names)} sheet(s): {', '.join(sheet_names)}")

        df = pd.read_excel(xls, sheet_name="general_report")

        # Step-by-step transformation
        df = df.sort_values(by="Client Versions", ascending=False)
        df = df.drop_duplicates(subset=["POS Code"], keep="first")
        df = df[df["Client Versions"] != 0]
        df["Amends"] = df["Client Versions"] - 1
        df["Right First Time"] = df["Client Versions"].apply(lambda x: 1 if x == 1 else 0)
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

        st.download_button("📥 Download Cleaned Data", df.to_excel(index=False), file_name="processed_artwork_data.xlsx")

    except Exception as e:
        st.error("❌ There was an issue processing your file.")
        st.exception(e)
else:
    st.info("👆 Please upload a .xlsx file to get started.")
