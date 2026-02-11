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
    # Sidebar for options
    with st.sidebar:
        st.header("⚙️ Options")
        st.markdown("---")
        
        # Option 1: Upload Range Reconciliation file directly
        st.subheader("📤 Upload Range Reconciliation File")
        uploaded_recon_file = st.file_uploader(
            "Range Reconciliation File (.xlsx)", 
            type=["xlsx"], 
            key="recon_file",
            help="Upload a previously generated Range_Reconciliation file"
        )
        
        if uploaded_recon_file:
            if 'uploaded_df' not in st.session_state or st.session_state.get('uploaded_file_name') != uploaded_recon_file.name:
                # Save uploaded file temporarily
                temp_file = tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx')
                temp_file.write(uploaded_recon_file.getvalue())
                temp_file.close()
                
                # Load the file
                df = load_reconciliation_file(temp_file.name)
                if df is not None:
                    st.session_state['uploaded_df'] = df
                    st.session_state['uploaded_file_name'] = uploaded_recon_file.name
                    st.success(f"File loaded: {uploaded_recon_file.name}")
                
                # Cleanup temp file
                os.unlink(temp_file.name)
        
        st.markdown("---")
        
        # Option 2: Upload source files and generate reconciliation
        st.subheader("📤 Upload Source Files")
        jeves_file = st.file_uploader("JEEVES File (.xlsx)", type=["xlsx"], key="jeves")
        ct_file = st.file_uploader("CT File (.xlsb)", type=["xlsb"], key="ct")
        stibo_file = st.file_uploader("STIBO File (.xlsx)", type=["xlsx"], key="stibo")
        
        if jeves_file and ct_file and stibo_file:
            if st.button("🔄 Generate Reconciliation", use_container_width=True):
                with st.spinner("Running reconciliation..."):
                    try:
                        reconciliation_df = run_reconciliation_from_upload(jeves_file, ct_file, stibo_file)
                        st.session_state['generated_df'] = reconciliation_df
                        st.session_state['generated_timestamp'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                        st.success("Reconciliation generated successfully!")
                        st.cache_data.clear()
                        st.rerun()
                    except Exception as e:
                        st.error(f"Error: {str(e)}")
                        import traceback
                        st.code(traceback.format_exc())
        
        st.markdown("---")
        
        # Option 3: Load from repo/local file (if exists)
        st.subheader("📁 Repository File")
        repo_file, _ = find_latest_reconciliation_file()
        if repo_file:
            st.caption(f"Available: {repo_file.name}")
            if st.button("📂 Load from Repository", use_container_width=True):
                df = load_reconciliation_file(repo_file)
                if df is not None:
                    st.session_state['local_df'] = df
                    st.session_state['local_file_name'] = repo_file.name
                    st.success("File loaded!")
                    st.cache_data.clear()
                    st.rerun()
        else:
            st.caption("No file in repository")
    
    # Determine which data source to use
    range_df = None
    source_info = None
    
    # Priority: 1) Uploaded file, 2) Generated from upload, 3) Local file, 4) Repo file
    if 'uploaded_df' in st.session_state:
        range_df = st.session_state['uploaded_df']
        source_info = f"Uploaded: {st.session_state.get('uploaded_file_name', 'Unknown')}"
    elif 'generated_df' in st.session_state:
        range_df = st.session_state['generated_df']
        source_info = f"Generated: {st.session_state.get('generated_timestamp', 'Unknown')}"
    elif 'local_df' in st.session_state:
        range_df = st.session_state['local_df']
        source_info = f"Local: {st.session_state.get('local_file_name', 'Unknown')}"
    else:
        # Try to load from repo file automatically
        repo_file, repo_df = find_latest_reconciliation_file()
        if repo_df is not None and repo_file:
            range_df = repo_df
            source_info = f"From repo: {repo_file.name}"
    
    if range_df is None:
        st.warning("⚠️ No reconciliation data available")
        st.info("""
        **Please choose one of these options:**
        
        1. **Upload Range Reconciliation file** - Upload a previously generated Excel file
        2. **Upload source files** - Upload JEEVES, CT, and STIBO files to generate reconciliation
        3. **Load local file** - If you've run `reconcile_products.py` locally
        """)
        return
    
    # Display data source info
    st.sidebar.markdown("---")
    st.sidebar.markdown("### 📊 Data Source")
    st.sidebar.caption(source_info if source_info else "Unknown")
    
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
        st.header("✅ Range Reconciliation")
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
        col1, col2, col3, col4 = st.columns(4)
        
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
        with col3:
            st.metric("In CT only", f"{len(range_pd[(range_pd['CT'] == 'X') & (range_pd['JEEVES'] == '') & (range_pd['STIBO'] == '')]):,}")
        with col4:
            st.metric("In JEEVES only", f"{len(range_pd[(range_pd['CT'] == '') & (range_pd['JEEVES'] == 'X') & (range_pd['STIBO'] == '')]):,}")
        
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
