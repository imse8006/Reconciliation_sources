"""Streamlit application to visualize reconciliation results"""
import streamlit as st
import polars as pl
import plotly.express as px
import plotly.graph_objects as go
from pathlib import Path
import glob
from datetime import datetime
import tempfile
import os

# Page configuration
st.set_page_config(
    page_title="Ekofisk Product Reconciliation",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Main title
st.title("📊 Ekofisk Product Reconciliation - JEEVES vs CT vs STIBO")
st.markdown("---")

@st.cache_data
def load_latest_analysis_files():
    """Load the most recent Range Reconciliation file"""
    # Search for most recent Range Reconciliation file (by modification date)
    range_files = list(Path(".").glob("Range_Reconciliation_*.xlsx"))
    
    range_file = None
    if range_files:
        range_file = max(range_files, key=lambda x: x.stat().st_mtime)
    
    try:
        range_df = None
        if range_file:
            range_df = pl.read_excel(range_file)
        
        return range_df, range_file
    except Exception as e:
        st.error(f"Error loading file: {e}")
        return None, None

def run_reconciliation(jeves_file, ct_file, stibo_file):
    """Run reconciliation with uploaded files"""
    import reconcile_products as rp
    
    # Save uploaded files temporarily
    temp_dir = Path(tempfile.mkdtemp())
    
    jeves_path = temp_dir / "jeves.xlsx"
    ct_path = temp_dir / "ct.xlsb"
    stibo_path = temp_dir / "stibo.xlsx"
    
    # Save files
    with open(jeves_path, "wb") as f:
        f.write(jeves_file.getvalue())
    with open(ct_path, "wb") as f:
        f.write(ct_file.getvalue())
    with open(stibo_path, "wb") as f:
        f.write(stibo_file.getvalue())
    
    # Load data
    jeves_df = rp.load_jeves_data(str(jeves_path))
    ct_df = rp.load_ct_data(str(ct_path))
    stibo_df = rp.load_stibo_data(str(stibo_path))
    
    # Create reconciliation
    reconciliation = rp.create_range_reconciliation(jeves_df, ct_df, stibo_df)
    
    # Cleanup
    import shutil
    shutil.rmtree(temp_dir)
    
    return reconciliation

def main():
    # Sidebar for options
    with st.sidebar:
        st.header("⚙️ Options")
        st.markdown("---")
        
        # File upload section
        st.subheader("📤 Upload Data Files")
        jeves_file = st.file_uploader("JEEVES File (.xlsx)", type=["xlsx"], key="jeves")
        ct_file = st.file_uploader("CT File (.xlsb)", type=["xlsb"], key="ct")
        stibo_file = st.file_uploader("STIBO File (.xlsx)", type=["xlsx"], key="stibo")
        
        if jeves_file and ct_file and stibo_file:
            if st.button("🔄 Run Reconciliation", use_container_width=True):
                with st.spinner("Running reconciliation..."):
                    try:
                        reconciliation_df = run_reconciliation(jeves_file, ct_file, stibo_file)
                        # Save to session state
                        st.session_state['reconciliation_df'] = reconciliation_df
                        st.session_state['reconciliation_source'] = 'uploaded'
                        st.success("Reconciliation completed!")
                        st.cache_data.clear()
                        st.rerun()
                    except Exception as e:
                        st.error(f"Error: {str(e)}")
        
        st.markdown("---")
        
        # Option to regenerate from local files
        if st.button("🔄 Regenerate from Local Files", use_container_width=True):
            st.info("Running reconciliation...")
            import subprocess
            import sys
            result = subprocess.run([sys.executable, "reconcile_products.py"], 
                                  capture_output=True, text=True)
            if result.returncode == 0:
                st.success("Reconciliation regenerated successfully!")
                st.cache_data.clear()
                st.rerun()
            else:
                st.error(f"Error: {result.stderr}")
    
    # Load data
    range_df = None
    range_file = None
    
    # Check if we have reconciliation from upload
    if 'reconciliation_df' in st.session_state:
        range_df = st.session_state['reconciliation_df']
        range_file = None
    else:
        # Try to load from file
        range_df, range_file = load_latest_analysis_files()
    
    if range_df is None:
        st.warning("⚠️ No Range Reconciliation file found. Please upload data files or run `reconcile_products.py`")
        st.info("""
        **Instructions:**
        1. Upload the three data files using the sidebar
        2. Click "Run Reconciliation" to generate the reconciliation
        3. Or run `reconcile_products.py` locally to generate the file
        """)
        return
    
    # Display loaded file info
    st.sidebar.markdown("---")
    st.sidebar.markdown("### 📁 File Info")
    if range_file:
        st.sidebar.caption(f"Range Recon: {Path(range_file).name}")
    elif 'reconciliation_source' in st.session_state and st.session_state['reconciliation_source'] == 'uploaded':
        st.sidebar.caption("Source: Uploaded files")
    
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
            st.plotly_chart(fig_pie, use_container_width=True)
        
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
            st.plotly_chart(fig_bar, use_container_width=True)
        
        # Data table
        st.subheader("Detailed Data")
        
        # Display options
        num_rows = st.slider("Number of rows to display", 10, min(1000, len(filtered_range)), 100)
        
        st.dataframe(
            filtered_range[["ProductCode", "CT", "JEEVES", "STIBO", "Absent_from"]],
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
            st.plotly_chart(fig_pie, use_container_width=True)
        
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
            st.plotly_chart(fig_bar, use_container_width=True)
    

if __name__ == "__main__":
    main()
