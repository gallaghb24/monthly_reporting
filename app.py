import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="Monthly Reporting Analysis", layout="wide")
st.title("Monthly Reporting Analysis")

# Create tabs
tab1, tab2 = st.tabs(["Content Production Analysis", "Stock Order Analysis"])

with tab1:
    st.header("Content Production Analysis")
    st.markdown("Upload your Monthly Versions Client Excel Export below  — we’ll analyse amends, right-first-time rate, and show a category breakdowns ready to copy and paste into Keynote")

    uploaded_file = st.file_uploader("Upload your Excel file", type=["xlsx"], key="content_upload")

    if uploaded_file:
        try:
            xls = pd.ExcelFile(uploaded_file)
            df = pd.read_excel(xls, sheet_name="general_report", header=1)

            col_map = {col.strip().lower(): col for col in df.columns}
            if "client versions" not in col_map:
                st.error("Couldn’t find a column called 'Client Versions'. Please check your file.")
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

                num_new_artworks = len(df)
                total_amends = df["Amends"].sum()
                num_rft = df["Right First Time"].sum()
                rft_percentage = round((num_rft / num_new_artworks) * 100, 1)
                avg_amends = round(df["Amends"].mean(), 2)
                over_v3 = df[df[client_col] > 3].shape[0]
                over_v3_pct = round((over_v3 / num_new_artworks) * 100, 1)

                st.subheader("Overall Stats")
                col1, col2, col3 = st.columns(3)
                col1.metric("New Artworks", num_new_artworks)
                col2.metric("Total Amends", total_amends)
                col3.metric("Right First Time", f"{num_rft} ({rft_percentage}%)")

                col4, col5, col6 = st.columns(3)
                col4.metric("Average Amend Rate", avg_amends)
                col5.metric("Artworks Beyond V3", f"{over_v3} ({over_v3_pct}%)")

                summary = pd.DataFrame()
                categories = sorted(df["Category"].dropna().unique())
                summary.loc["New Artwork Lines", categories] = df.groupby("Category").size()
                summary.loc["Amends", categories] = df.groupby("Category")["Amends"].sum()
                summary.loc["Right First Time", categories] = df.groupby("Category")["Right First Time"].sum()
                summary.loc["Average Round of Amends", categories] = df.groupby("Category")["Amends"].mean().round(2)

                summary_display = summary.reset_index().rename(columns={"index": ""})
                table_1 = summary_display[summary_display[""] != "Average Round of Amends"]
                table_2 = summary_display[summary_display[""] == "Average Round of Amends"]

                st.subheader("Volume by Area")
                st.dataframe(table_1, use_container_width=True)

                st.subheader("Average Rounds of Amends by Area")
                st.dataframe(table_2, use_container_width=True)

                version_counts = df[client_col].value_counts().sort_index()
                version_table = pd.DataFrame([version_counts])
                version_table.index = ["Amends"]
                version_table.columns = [f"V{int(col)}" for col in version_table.columns]
                version_display = version_table.reset_index().rename(columns={"index": ""})

                st.subheader("Volume by Version Number")
                st.dataframe(version_display, use_container_width=True)

                output = BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    summary_display.to_excel(writer, index=False, sheet_name="Summary")
                    version_display.to_excel(writer, index=False, sheet_name="Version Breakdown")

                st.download_button(
                    label="Download Summary Table",
                    data=output.getvalue(),
                    file_name="content_creation_summary.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

        except Exception as e:
            st.error("There was an issue processing your file.")
            st.exception(e)
    else:
        st.info("Please upload a .xlsx file to get started.")

with tab2:
    st.header("Stock Order Analysis")
    st.markdown("Upload stock data below to generate summary statistics.")
    stock_file = st.file_uploader("Upload your Stock Excel file", type=["xlsx"], key="stock_upload")
    if stock_file:
        st.success("Stock data uploaded! (Summary analysis functionality coming soon.)")
