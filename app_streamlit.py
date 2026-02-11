"""Streamlit application to visualize reconciliation results"""
import streamlit as st
import polars as pl
import plotly.express as px
import plotly.graph_objects as go
from pathlib import Path
from datetime import datetime

# Page configuration
st.set_page_config(
    page_title="Ekofisk Reconciliation",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

@st.cache_data
def load_reconciliation_file(file_path):
    """Load Range Reconciliation Excel file"""
    try:
        return pl.read_excel(file_path)
    except Exception as e:
        st.error(f"Error loading file: {e}")
        return None

@st.cache_data
def find_latest_reconciliation_file():
    """Find the latest Range Reconciliation file in the repo"""
    range_files = list(Path(".").glob("Range_Reconciliation_*.xlsx"))
    if range_files:
        latest_file = max(range_files, key=lambda x: x.stat().st_mtime)
        return latest_file, load_reconciliation_file(latest_file)
    return None, None

def main():
    # Navigation sidebar
    with st.sidebar:
        st.header("📊 Navigation")
        st.markdown("---")
        
        # Main level: Ekofisk
        st.subheader("Ekofisk")
        
        # Domain selection
        domain = st.radio(
            "Domain",
            ["Product", "Vendor", "Customer"],
            index=0,
            key="domain_selector"
        )
        
        st.markdown("---")
        
        # Display current selection
        st.caption(f"📍 **Ekofisk > {domain}**")
    
    # Main content area
    if domain == "Product":
        show_product_reconciliation()
    elif domain == "Vendor":
        st.title("📊 Ekofisk Vendor Reconciliation")
        st.info("🚧 Vendor reconciliation coming soon...")
    elif domain == "Customer":
        st.title("📊 Ekofisk Customer Reconciliation")
        st.info("🚧 Customer reconciliation coming soon...")

def show_product_reconciliation():
    """Display Product reconciliation view"""
    st.title("📊 Ekofisk Product Reconciliation - JEEVES vs CT vs STIBO")
    st.markdown("---")
    
    # Load data automatically from repo
    repo_file, repo_df = find_latest_reconciliation_file()
    
    if repo_df is None or repo_file is None:
        st.warning("⚠️ No reconciliation data available")
        st.info("Please ensure that `Range_Reconciliation_*.xlsx` file exists in the repository.")
        return
    
    range_df = repo_df
    
    # Display data source info in sidebar
    with st.sidebar:
        st.markdown("---")
        st.markdown("### 📊 Data Source")
        st.caption(f"From repo: {repo_file.name}")
    
    # Convert to pandas for Streamlit (easier for display)
    range_pd = range_df.to_pandas()
    
    # Check that expected columns exist
    required_cols = ["ProductCode", "CT", "JEEVES", "STIBO", "Absent_from"]
    missing_cols = [col for col in required_cols if col not in range_pd.columns]
    
    if missing_cols:
        st.error(f"Missing columns: {missing_cols}")
        st.info(f"Available columns: {list(range_pd.columns)}")
        return
    
    # Tabs for different views
    tab_range, tab_overview = st.tabs(["✅ Range Reconciliation", "📈 Overview"])
    
    with tab_range:
        st.header("Range Reconciliation")
        st.markdown("List of all products with their presence in CT, JEEVES and STIBO")
        
        # Statistics
        col1, col2, col3, col4 = st.columns(4)
        total_products = len(range_pd)
        ct_count = len(range_pd[range_pd["CT"] == "X"])
        jeves_count = len(range_pd[range_pd["JEEVES"] == "X"])
        stibo_count = len(range_pd[range_pd["STIBO"] == "X"])
        
        with col1:
            st.metric("Total products", f"{total_products:,}")
        with col2:
            st.metric("In CT", f"{ct_count:,}")
        with col3:
            st.metric("In JEEVES", f"{jeves_count:,}")
        with col4:
            st.metric("In STIBO", f"{stibo_count:,}")
        
        st.markdown("---")
        
        # Filters
        col1, col2, col3 = st.columns(3)
        with col1:
            filter_ct = st.selectbox("CT", ["All", "X Present", "Absent"], index=0)
        with col2:
            filter_jeves = st.selectbox("JEEVES", ["All", "X Present", "Absent"], index=0)
        with col3:
            filter_stibo = st.selectbox("STIBO", ["All", "X Present", "Absent"], index=0)
        
        # Search
        search_term = st.text_input("🔍 Search product (code)", "")
        
        # Apply filters
        filtered_range = range_pd.copy()
        
        if filter_ct != "All":
            filtered_range = filtered_range[filtered_range["CT"] == ("X" if filter_ct == "X Present" else "")]
        if filter_jeves != "All":
            filtered_range = filtered_range[filtered_range["JEEVES"] == ("X" if filter_jeves == "X Present" else "")]
        if filter_stibo != "All":
            filtered_range = filtered_range[filtered_range["STIBO"] == ("X" if filter_stibo == "X Present" else "")]
        
        if search_term:
            filtered_range = filtered_range[
                filtered_range["ProductCode"].astype(str).str.contains(search_term, case=False, na=False)
            ]
        
        st.info(f"📊 {len(filtered_range)} products displayed out of {total_products} total")
        
        # Visualizations
        col_left, col_right = st.columns(2)
        
        with col_left:
            # Distribution chart
            status_counts = {
                "In all 3": len(filtered_range[
                    (filtered_range["CT"] == "X") & 
                    (filtered_range["JEEVES"] == "X") & 
                    (filtered_range["STIBO"] == "X")
                ]),
                "In 2": len(filtered_range[
                    ((filtered_range["CT"] == "X").astype(int) + 
                     (filtered_range["JEEVES"] == "X").astype(int) + 
                     (filtered_range["STIBO"] == "X").astype(int)) == 2
                ]),
                "In 1": len(filtered_range[
                    ((filtered_range["CT"] == "X").astype(int) + 
                     (filtered_range["JEEVES"] == "X").astype(int) + 
                     (filtered_range["STIBO"] == "X").astype(int)) == 1
                ]),
                "In none": len(filtered_range[
                    (filtered_range["CT"] == "") & 
                    (filtered_range["JEEVES"] == "") & 
                    (filtered_range["STIBO"] == "")
                ])
            }
            
            fig_pie = px.pie(
                values=list(status_counts.values()),
                names=list(status_counts.keys()),
                title="Distribution by number of sources"
            )
            st.plotly_chart(fig_pie, use_container_width=True, key="pie_range")
        
        with col_right:
            # Bar chart by source
            source_counts = {
                "CT": ct_count,
                "JEEVES": jeves_count,
                "STIBO": stibo_count
            }
            fig_bar = px.bar(
                x=list(source_counts.keys()),
                y=list(source_counts.values()),
                title="Number of products by source",
                labels={"x": "Source", "y": "Number of products"}
            )
            st.plotly_chart(fig_bar, use_container_width=True, key="bar_range")
        
        # Data table
        st.subheader("Detailed Data")
        
        # Sort by Absent_from descending: non-empty values first (products missing from sources)
        filtered_range_sorted = filtered_range.copy()
        # Convert to string and handle NaN/None as empty string for consistent sorting
        filtered_range_sorted['Absent_from'] = filtered_range_sorted['Absent_from'].astype(str).replace('nan', '').replace('None', '')
        # Sort descending: non-empty values will be at the top
        filtered_range_sorted = filtered_range_sorted.sort_values(
            by='Absent_from',
            ascending=False,
            na_position='last'
        )
        
        # Display 705 rows
        num_rows = 705
        
        st.dataframe(
            filtered_range_sorted[["ProductCode", "CT", "JEEVES", "STIBO", "Absent_from"]].head(num_rows),
            use_container_width=True,
            height=400
        )
        
        # Download
        csv_range = filtered_range.to_csv(index=False)
        st.download_button(
            label="📥 Download Range Reconciliation (CSV)",
            data=csv_range,
            file_name=f"range_reconciliation_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
            mime="text/csv"
        )
    
    with tab_overview:
        st.header("Overview")
        
        # Main metrics
        col1, col2, col3, col4, col5 = st.columns(5)
        
        total_products = len(range_pd)
        ct_count = len(range_pd[range_pd["CT"] == "X"])
        jeves_count = len(range_pd[range_pd["JEEVES"] == "X"])
        stibo_count = len(range_pd[range_pd["STIBO"] == "X"])
        all_three = len(range_pd[
            (range_pd["CT"] == "X") & 
            (range_pd["JEEVES"] == "X") & 
            (range_pd["STIBO"] == "X")
        ])
        
        with col1:
            st.metric("Total unique products", f"{total_products:,}")
        with col2:
            st.metric("In all 3 sources", f"{all_three:,}", 
                     delta=f"{all_three/total_products*100:.1f}%" if total_products > 0 else "0%")
        # Fill NaN values with empty string for consistent comparison
        range_pd_filled = range_pd.fillna({'CT': '', 'JEEVES': '', 'STIBO': ''})
        
        with col3:
            ct_only = range_pd_filled[(range_pd_filled['CT'] == 'X') & (range_pd_filled['JEEVES'] != 'X') & (range_pd_filled['STIBO'] != 'X')]
            st.metric("In CT only", f"{len(ct_only):,}")
        with col4:
            jeves_only = range_pd_filled[(range_pd_filled['CT'] != 'X') & (range_pd_filled['JEEVES'] == 'X') & (range_pd_filled['STIBO'] != 'X')]
            st.metric("In JEEVES only", f"{len(jeves_only):,}")
        with col5:
            stibo_only = range_pd_filled[(range_pd_filled['CT'] != 'X') & (range_pd_filled['JEEVES'] != 'X') & (range_pd_filled['STIBO'] == 'X')]
            st.metric("In STIBO only", f"{len(stibo_only):,}")
        
        st.markdown("---")
        
        # Pie chart for presence
        col_left, col_right = st.columns(2)
        
        with col_left:
            # Count products by presence pattern
            presence_patterns = {
                "All 3": all_three,
                "CT + JEEVES": len(range_pd[(range_pd["CT"] == "X") & (range_pd["JEEVES"] == "X") & (range_pd["STIBO"] == "")]),
                "CT + STIBO": len(range_pd[(range_pd["CT"] == "X") & (range_pd["JEEVES"] == "") & (range_pd["STIBO"] == "X")]),
                "JEEVES + STIBO": len(range_pd[(range_pd["CT"] == "") & (range_pd["JEEVES"] == "X") & (range_pd["STIBO"] == "X")]),
                "CT only": len(range_pd[(range_pd["CT"] == "X") & (range_pd["JEEVES"] == "") & (range_pd["STIBO"] == "")]),
                "JEEVES only": len(range_pd[(range_pd["CT"] == "") & (range_pd["JEEVES"] == "X") & (range_pd["STIBO"] == "")]),
                "STIBO only": len(range_pd[(range_pd["CT"] == "") & (range_pd["JEEVES"] == "") & (range_pd["STIBO"] == "X")]),
                "None": len(range_pd[(range_pd["CT"] == "") & (range_pd["JEEVES"] == "") & (range_pd["STIBO"] == "")])
            }
            
            fig_pie = px.pie(
                values=list(presence_patterns.values()),
                names=list(presence_patterns.keys()),
                title="Product distribution by source combination"
            )
            st.plotly_chart(fig_pie, use_container_width=True, key="pie_overview")
        
        with col_right:
            # Bar chart
            source_counts = {
                "CT": ct_count,
                "JEEVES": jeves_count,
                "STIBO": stibo_count
            }
            fig_bar = px.bar(
                x=list(source_counts.keys()),
                y=list(source_counts.values()),
                title="Number of products by source",
                labels={"x": "Source", "y": "Number of products"}
            )
            st.plotly_chart(fig_bar, use_container_width=True, key="bar_overview")
    

if __name__ == "__main__":
    main()
