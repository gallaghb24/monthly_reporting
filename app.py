import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
from io import BytesIO
import base64

# Page configuration
st.set_page_config(page_title="Monthly Reporting Analysis", layout="wide")

# Custom styling
st.markdown("""
<style>
    .main-header {
        color: #ff1493;
        font-size: 3rem;
        font-weight: 700;
        text-align: center;
        margin-bottom: 2rem;
    }
    .section-header {
        color: #ff1493;
        font-size: 2rem;
        font-weight: 600;
        margin: 2rem 0 1rem 0;
    }
    .key-stats-box {
        background: linear-gradient(135deg, #ff1493 0%, #ff69b4 100%);
        color: white;
        padding: 2rem;
        border-radius: 15px;
        margin: 1rem 0;
        box-shadow: 0 8px 25px rgba(255, 20, 147, 0.3);
    }
    .metric-card {
        background: white;
        border: 2px solid #ff1493;
        border-radius: 10px;
        padding: 1.5rem;
        text-align: center;
        box-shadow: 0 4px 15px rgba(0,0,0,0.1);
    }
    .metric-value {
        font-size: 2.5rem;
        font-weight: 700;
        color: #ff1493;
        margin: 0;
    }
    .metric-label {
        color: #666;
        font-size: 0.9rem;
        margin-top: 0.5rem;
    }
    .chart-container {
        background: white;
        padding: 2rem;
        border-radius: 15px;
        box-shadow: 0 4px 15px rgba(0,0,0,0.1);
        margin: 2rem 0;
    }
</style>
""", unsafe_allow_html=True)

def create_styled_bar_chart(data, title, x_col, y_cols, colors):
    """Create a publication-ready matplotlib chart"""
    fig, ax = plt.subplots(figsize=(14, 8))
    
    # Set the style
    plt.style.use('default')
    
    # Create grouped bar chart
    x = range(len(data))
    width = 0.25
    
    for i, (col, color) in enumerate(zip(y_cols, colors)):
        offset = (i - len(y_cols)/2 + 0.5) * width
        bars = ax.bar([pos + offset for pos in x], data[col], width, 
                     label=col, color=color, alpha=0.8)
        
        # Add value labels on bars
        for bar in bars:
            height = bar.get_height()
            if height > 0:
                ax.text(bar.get_x() + bar.get_width()/2., height + max(data[y_cols].values.flatten()) * 0.01,
                       f'{int(height)}', ha='center', va='bottom', fontweight='bold', fontsize=10)
    
    # Customize the chart
    ax.set_xlabel('')
    ax.set_ylabel('')
    ax.set_title(title, fontsize=20, fontweight='bold', color='#333', pad=20)
    ax.set_xticks(x)
    ax.set_xticklabels(data[x_col], rotation=45, ha='right')
    ax.legend(loc='upper right', frameon=True, fancybox=True, shadow=True)
    
    # Style the chart
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    ax.grid(True, alpha=0.3, linestyle='-', linewidth=0.5)
    ax.set_facecolor('#fafafa')
    
    plt.tight_layout()
    return fig

def create_version_chart(version_data):
    """Create version distribution chart"""
    fig, ax = plt.subplots(figsize=(14, 6))
    
    bars = ax.bar(version_data.index, version_data.values, 
                  color='#ff1493', alpha=0.8, edgecolor='white', linewidth=2)
    
    # Add value labels
    for bar in bars:
        height = bar.get_height()
        ax.text(bar.get_x() + bar.get_width()/2., height + max(version_data.values) * 0.01,
               f'{int(height)}', ha='center', va='bottom', fontweight='bold', fontsize=12)
    
    ax.set_title('Volume by Version Number', fontsize=20, fontweight='bold', color='#333', pad=20)
    ax.set_xlabel('')
    ax.set_ylabel('')
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    ax.grid(True, alpha=0.3, linestyle='-', linewidth=0.5)
    ax.set_facecolor('#fafafa')
    
    plt.tight_layout()
    return fig

# Main app
st.markdown('<h1 class="main-header">Monthly Reporting Analysis</h1>', unsafe_allow_html=True)

tab1, tab2 = st.tabs(["📊 Content Production Analysis", "📦 Stock Order Analysis"])

with tab1:
    st.markdown("Upload your Monthly Versions Client Excel Export below")
    
    uploaded_file = st.file_uploader("Upload your Excel file", type=["xlsx"], key="content_upload")
    
    if uploaded_file:
        try:
            # Your existing data processing logic here...
            xls = pd.ExcelFile(uploaded_file)
            df = pd.read_excel(xls, sheet_name="general_report", header=1)

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

                # Calculate statistics
                num_new_artworks = len(df)
                total_amends = df["Amends"].sum()
                num_rft = df["Right First Time"].sum()
                rft_percentage = round((num_rft / num_new_artworks) * 100, 1)
                avg_amends = round(df["Amends"].mean(), 2)
                over_v3 = df[df[client_col] > 3].shape[0]
                over_v3_pct = round((over_v3 / num_new_artworks) * 100, 1)

                # Key Statistics Display
                st.markdown("""
                <div class="key-stats-box">
                    <h2 style="margin-bottom: 1rem;">📊 Key Stats Summary</h2>
                    <div style="display: grid; grid-template-columns: repeat(2, 1fr); gap: 1rem;">
                        <div><strong>{} new artworks created</strong></div>
                        <div><strong>{} rounds of amends</strong></div>
                        <div><strong>{} right first time ({}%)</strong></div>
                        <div><strong>{} artworks beyond V3 ({}%)</strong></div>
                    </div>
                </div>
                """.format(num_new_artworks, total_amends, num_rft, rft_percentage, over_v3, over_v3_pct), 
                unsafe_allow_html=True)

                # Metrics cards
                col1, col2, col3, col4 = st.columns(4)
                
                with col1:
                    st.markdown(f"""
                    <div class="metric-card">
                        <div class="metric-value">{avg_amends}</div>
                        <div class="metric-label">Average Amends</div>
                    </div>
                    """, unsafe_allow_html=True)
                
                with col2:
                    st.markdown(f"""
                    <div class="metric-card">
                        <div class="metric-value">{rft_percentage}%</div>
                        <div class="metric-label">Right First Time</div>
                    </div>
                    """, unsafe_allow_html=True)
                
                with col3:
                    st.markdown(f"""
                    <div class="metric-card">
                        <div class="metric-value">{num_new_artworks}</div>
                        <div class="metric-label">Total Artworks</div>
                    </div>
                    """, unsafe_allow_html=True)
                
                with col4:
                    st.markdown(f"""
                    <div class="metric-card">
                        <div class="metric-value">{over_v3_pct}%</div>
                        <div class="metric-label">Beyond V3</div>
                    </div>
                    """, unsafe_allow_html=True)

                st.markdown('<div style="margin: 3rem 0;"></div>', unsafe_allow_html=True)

                # Prepare data for charts
                categories = sorted(df["Category"].dropna().unique())
                chart_data = []
                
                for category in categories:
                    cat_data = df[df["Category"] == category]
                    chart_data.append({
                        'Category': category,
                        'New Artwork Lines': len(cat_data),
                        'Amends': cat_data["Amends"].sum(),
                        'Right First Time': cat_data["Right First Time"].sum()
                    })
                
                chart_df = pd.DataFrame(chart_data)

                # Volume by Area Chart
                st.markdown('<div class="chart-container">', unsafe_allow_html=True)
                colors = ['#ff1493', '#ff69b4', '#8b4513']
                fig1 = create_styled_bar_chart(chart_df, 'Volume by Area', 'Category', 
                                             ['New Artwork Lines', 'Amends', 'Right First Time'], colors)
                st.pyplot(fig1, use_container_width=True)
                st.markdown('</div>', unsafe_allow_html=True)

                # Amends table
                st.markdown('<h3 class="section-header">Content Amends</h3>', unsafe_allow_html=True)
                amends_data = []
                for category in categories:
                    cat_data = df[df["Category"] == category]
                    amends_data.append(round(cat_data["Amends"].mean(), 2))
                
                amends_df = pd.DataFrame([amends_data], columns=categories, index=['Average Round of Amends'])
                
                # Style the dataframe
                styled_df = amends_df.style.set_properties(**{
                    'background-color': '#ff1493',
                    'color': 'white',
                    'font-weight': 'bold',
                    'text-align': 'center'
                }).set_table_styles([
                    {'selector': 'th', 'props': [('background-color', '#ff1493'), ('color', 'white'), ('font-weight', 'bold')]}
                ])
                
                st.dataframe(styled_df, use_container_width=True)

                # Version chart
                st.markdown('<div class="chart-container">', unsafe_allow_html=True)
                version_counts = df[client_col].value_counts().sort_index()
                version_counts.index = [f'V{int(v)}' for v in version_counts.index]
                fig2 = create_version_chart(version_counts)
                st.pyplot(fig2, use_container_width=True)
                st.markdown('</div>', unsafe_allow_html=True)

                # Download button
                st.markdown('<h3 class="section-header">Download Options</h3>', unsafe_allow_html=True)
                
                # Prepare Excel download
                output = BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    chart_df.to_excel(writer, index=False, sheet_name="Summary")
                    pd.DataFrame({'Version': version_counts.index, 'Count': version_counts.values}).to_excel(
                        writer, index=False, sheet_name="Version_Breakdown")

                st.download_button(
                    label="📊 Download Summary Tables",
                    data=output.getvalue(),
                    file_name="content_analysis_summary.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

        except Exception as e:
            st.error("There was an issue processing your file.")
            st.exception(e)
    else:
        # Show demo data
        st.info("Upload a file to see your data, or view the demo below:")
        
        # Demo with sample data
        st.markdown("""
        <div class="key-stats-box">
            <h2 style="margin-bottom: 1rem;">📊 Demo Key Stats</h2>
            <div style="display: grid; grid-template-columns: repeat(2, 1fr); gap: 1rem;">
                <div><strong>914 new artworks created</strong></div>
                <div><strong>1,056 rounds of amends</strong></div>
                <div><strong>309 right first time (33.8%)</strong></div>
                <div><strong>119 artworks beyond V3 (13.0%)</strong></div>
            </div>
        </div>
        """, unsafe_allow_html=True)

with tab2:
    st.markdown('<h2 class="section-header">Stock Order Analysis</h2>', unsafe_allow_html=True)
    st.markdown("Upload Order Line Level Data Export for the month below")
    
    stock_file = st.file_uploader("Upload your Excel file", type=["xlsx"], key="stock_upload")
    
    if stock_file:
        # Your existing stock analysis logic here
        pass
    else:
        st.info("Please upload a .xlsx file for stock analysis.")
