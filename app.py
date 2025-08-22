import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
from io import BytesIO
import base64

# Page configuration
st.set_page_config(page_title="Monthly Reporting Analysis", layout="wide")

# Custom CSS for styling
st.markdown("""
<style>
    .metric-card {
        background-color: #f0f2f6;
        padding: 1rem;
        border-radius: 0.5rem;
        border-left: 4px solid #ff1493;
    }
    .stMetric > label {
        color: #262730 !important;
        font-size: 1rem !important;
        font-weight: 600 !important;
    }
    .stMetric > div {
        color: #ff1493 !important;
        font-size: 2rem !important;
        font-weight: 700 !important;
    }
    .main-title {
        color: #ff1493;
        font-size: 3rem;
        font-weight: 700;
        margin-bottom: 0.5rem;
    }
    .section-header {
        color: #ff1493;
        font-size: 2rem;
        font-weight: 600;
        margin: 2rem 0 1rem 0;
    }
</style>
""", unsafe_allow_html=True)

st.markdown('<h1 class="main-title">Monthly Reporting Analysis</h1>', unsafe_allow_html=True)

# Create tabs
tab1, tab2 = st.tabs(["📊 Content Production Analysis", "📦 Stock Order Analysis"])

def create_volume_by_area_chart(df):
    """Create grouped bar chart for volume by area"""
    # Prepare data
    summary_data = []
    categories = sorted(df["Category"].dropna().unique())
    
    for category in categories:
        cat_data = df[df["Category"] == category]
        summary_data.append({
            'Area': category,
            'New Artwork Lines': len(cat_data),
            'Amends': cat_data["Amends"].sum(),
            'Right First Time': cat_data["Right First Time"].sum()
        })
    
    summary_df = pd.DataFrame(summary_data)
    
    # Create grouped bar chart
    fig = go.Figure()
    
    # Define colors matching your slides
    colors = {
        'New Artwork Lines': '#ff1493',  # Bright pink
        'Amends': '#ff69b4',  # Medium pink  
        'Right First Time': '#8b4513'   # Brown
    }
    
    for metric in ['New Artwork Lines', 'Amends', 'Right First Time']:
        fig.add_trace(go.Bar(
            name=metric,
            x=summary_df['Area'],
            y=summary_df[metric],
            marker_color=colors[metric],
            text=summary_df[metric],
            textposition='outside'
        ))
    
    fig.update_layout(
        title={
            'text': 'Volume by Area',
            'font': {'size': 24, 'color': '#262730'},
            'x': 0.5
        },
        xaxis_title='',
        yaxis_title='',
        barmode='group',
        height=500,
        showlegend=True,
        legend=dict(
            orientation="h",
            yanchor="bottom",
            y=1.02,
            xanchor="right",
            x=1
        ),
        plot_bgcolor='white',
        paper_bgcolor='white'
    )
    
    fig.update_xaxes(showgrid=False)
    fig.update_yaxes(showgrid=True, gridcolor='lightgray')
    
    return fig

def create_version_chart(df, client_col):
    """Create bar chart for volume by version number"""
    version_counts = df[client_col].value_counts().sort_index()
    
    fig = go.Figure(data=[
        go.Bar(
            x=[f'V{int(v)}' for v in version_counts.index],
            y=version_counts.values,
            marker_color='#ff1493',
            text=version_counts.values,
            textposition='outside'
        )
    ])
    
    fig.update_layout(
        title={
            'text': 'Volume by Version Number',
            'font': {'size': 24, 'color': '#262730'},
            'x': 0.5
        },
        xaxis_title='',
        yaxis_title='',
        height=400,
        plot_bgcolor='white',
        paper_bgcolor='white'
    )
    
    fig.update_xaxes(showgrid=False)
    fig.update_yaxes(showgrid=True, gridcolor='lightgray')
    
    return fig

def create_amends_table(df):
    """Create styled table for average rounds of amends"""
    categories = sorted(df["Category"].dropna().unique())
    amends_data = []
    
    for category in categories:
        cat_data = df[df["Category"] == category]
        avg_amends = round(cat_data["Amends"].mean(), 2)
        amends_data.append(avg_amends)
    
    # Create DataFrame for table
    table_df = pd.DataFrame([amends_data], columns=categories)
    table_df.index = ['Average Round of Amends']
    
    return table_df

def download_charts_as_images(volume_fig, version_fig, summary_stats):
    """Generate downloadable images of the charts"""
    # Convert plots to images
    volume_img = volume_fig.to_image(format="png", width=1200, height=600, scale=2)
    version_img = version_fig.to_image(format="png", width=1200, height=600, scale=2)
    
    return volume_img, version_img

with tab1:
    st.markdown('<h2 class="section-header">Content Production Analysis</h2>', unsafe_allow_html=True)
    st.markdown("Upload your Monthly Versions Client Excel Export below — we'll analyze amends, right-first-time rate, and create presentation-ready charts")

    uploaded_file = st.file_uploader("Upload your Excel file", type=["xlsx"], key="content_upload")

    if uploaded_file:
        try:
            # Load and process data
            xls = pd.ExcelFile(uploaded_file)
            df = pd.read_excel(xls, sheet_name="general_report", header=1)

            # Column mapping
            col_map = {col.strip().lower(): col for col in df.columns}
            if "client versions" not in col_map:
                st.error("Couldn't find a column called 'Client Versions'. Please check your file.")
            else:
                client_col = col_map["client versions"]

                # Data processing
                df = df.sort_values(by=client_col, ascending=False)
                df = df.drop_duplicates(subset=["POS Code"], keep="first")
                df = df[df[client_col] != 0]
                df["Amends"] = df[client_col] - 1
                df["Right First Time"] = df[client_col].apply(lambda x: 1 if x == 1 else 0)
                
                # Category assignment
                df.loc[df["Project Description"].str.contains("ROI", na=False), "Category"] = "ROI"
                df.loc[df["Category"].isin(["Members", "Starbuys"]), "Category"] = "Main Event"
                df.loc[df["Category"].isin(["Loyalty / CRM", "Mobile"]), "Category"] = "Other"

                # Calculate key statistics
                num_new_artworks = len(df)
                total_amends = df["Amends"].sum()
                num_rft = df["Right First Time"].sum()
                rft_percentage = round((num_rft / num_new_artworks) * 100, 1)
                avg_amends = round(df["Amends"].mean(), 2)
                over_v3 = df[df[client_col] > 3].shape[0]
                over_v3_pct = round((over_v3 / num_new_artworks) * 100, 1)

                # Key Statistics Section
                st.markdown('<h3 class="section-header">Key Statistics</h3>', unsafe_allow_html=True)
                
                col1, col2 = st.columns(2)
                
                with col1:
                    st.markdown("""
                    <div style='background-color: #f8f9fa; padding: 2rem; border-radius: 10px; border-left: 6px solid #ff1493;'>
                        <h2 style='color: #ff1493; margin-bottom: 1rem;'>Key Stats Summary</h2>
                        <p style='font-size: 1.2rem; margin: 0.5rem 0;'><strong>{} new artworks created</strong></p>
                        <p style='font-size: 1.2rem; margin: 0.5rem 0;'><strong>{} rounds of amends</strong></p>
                        <p style='font-size: 1.2rem; margin: 0.5rem 0;'><strong>{} right first time ({}%)</strong></p>
                        <p style='font-size: 1.2rem; margin: 0.5rem 0;'><strong>{} artworks went beyond version 3 ({}%)</strong></p>
                    </div>
                    """.format(num_new_artworks, total_amends, num_rft, rft_percentage, over_v3, over_v3_pct), 
                    unsafe_allow_html=True)
                
                with col2:
                    # Additional metrics in cards
                    st.metric("Average Amends per Artwork", f"{avg_amends}")
                    st.metric("Efficiency Score", f"{rft_percentage}%")

                # Charts Section
                st.markdown('<h3 class="section-header">Presentation Charts</h3>', unsafe_allow_html=True)
                
                # Volume by Area Chart
                volume_fig = create_volume_by_area_chart(df)
                st.plotly_chart(volume_fig, use_container_width=True)
                
                # Average Amends Table
                st.markdown("### Content Amends")
                amends_table = create_amends_table(df)
                
                # Style the table to match your presentation
                st.markdown("""
                <style>
                .amends-table {
                    background: linear-gradient(90deg, #ff1493 0%, #ff1493 100%);
                    color: white;
                    font-weight: bold;
                    text-align: center;
                }
                </style>
                """, unsafe_allow_html=True)
                
                st.dataframe(amends_table.style.set_table_styles([
                    {'selector': 'thead th', 'props': [('background-color', '#ff1493'), ('color', 'white'), ('font-weight', 'bold')]},
                    {'selector': 'td', 'props': [('text-align', 'center')]}
                ]), use_container_width=True)
                
                # Version Number Chart
                version_fig = create_version_chart(df, client_col)
                st.plotly_chart(version_fig, use_container_width=True)

                # Download Options
                st.markdown('<h3 class="section-header">Download Options</h3>', unsafe_allow_html=True)
                
                col1, col2 = st.columns(2)
                
                with col1:
                    # Excel download
                    output = BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        # Summary table
                        summary_data = []
                        categories = sorted(df["Category"].dropna().unique())
                        for category in categories:
                            cat_data = df[df["Category"] == category]
                            summary_data.append({
                                'Category': category,
                                'New Artwork Lines': len(cat_data),
                                'Amends': cat_data["Amends"].sum(),
                                'Right First Time': cat_data["Right First Time"].sum(),
                                'Average Round of Amends': round(cat_data["Amends"].mean(), 2)
                            })
                        
                        summary_df = pd.DataFrame(summary_data)
                        summary_df.to_excel(writer, index=False, sheet_name="Summary")
                        
                        # Version breakdown
                        version_counts = df[client_col].value_counts().sort_index()
                        version_df = pd.DataFrame({
                            'Version': [f'V{int(v)}' for v in version_counts.index],
                            'Count': version_counts.values
                        })
                        version_df.to_excel(writer, index=False, sheet_name="Version_Breakdown")

                    st.download_button(
                        label="📊 Download Excel Summary",
                        data=output.getvalue(),
                        file_name=f"content_analysis_{pd.Timestamp.now().strftime('%Y%m%d')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                
                with col2:
                    # Chart downloads
                    if st.button("📈 Download Charts as Images"):
                        volume_img = volume_fig.to_image(format="png", width=1400, height=700, scale=2)
                        version_img = version_fig.to_image(format="png", width=1400, height=700, scale=2)
                        
                        st.success("Charts generated! Right-click on the charts above and 'Save image as...' to download them for your presentation.")

        except Exception as e:
            st.error("There was an issue processing your file.")
            st.exception(e)
            
    else:
        st.info("Please upload a .xlsx file to get started.")

with tab2:
    st.markdown('<h2 class="section-header">Stock Order Analysis</h2>', unsafe_allow_html=True)
    st.markdown("Upload Order Line Level Data Export for the month below to generate summary stats")
    
    stock_file = st.file_uploader("Upload your Excel file", type=["xlsx"], key="stock_upload")
    
    if stock_file:
        try:
            df_stock = pd.read_excel(stock_file, header=1)

            # Categorise Order Type
            df_stock["Order Type"] = df_stock["Ordered By"].str.lower().str.contains("store")
            df_stock["Order Type"] = df_stock["Order Type"].map({True: "Store Order", False: "Helpdesk Order"})

            # Count unique order numbers per order type
            unique_orders = df_stock.drop_duplicates(subset=["Order Number"])
            order_type_counts = unique_orders["Order Type"].value_counts().reset_index()
            order_type_counts.columns = ["Order Type", "Unique Order Count"]

            st.markdown("### Order Type Breakdown")
            
            # Create pie chart for order types
            fig_pie = px.pie(order_type_counts, values='Unique Order Count', names='Order Type',
                           color_discrete_sequence=['#ff1493', '#ff69b4'])
            fig_pie.update_layout(height=400)
            st.plotly_chart(fig_pie, use_container_width=True)

            # Top 10 locations by order line volume
            top_locations = (
                df_stock.groupby(["Location Code", "Location Name"])
                .size()
                .reset_index(name="Order Line Count")
                .sort_values(by="Order Line Count", ascending=False)
                .head(10)
            )

            st.markdown("### Top 10 Locations by Order Line Volume")
            
            # Create horizontal bar chart for locations
            fig_locations = px.bar(top_locations, 
                                 x='Order Line Count', 
                                 y='Location Name',
                                 orientation='h',
                                 color_discrete_sequence=['#ff1493'])
            fig_locations.update_layout(height=500, yaxis={'categoryorder':'total ascending'})
            st.plotly_chart(fig_locations, use_container_width=True)

        except Exception as e:
            st.error("There was an issue processing the stock order file.")
            st.exception(e)
