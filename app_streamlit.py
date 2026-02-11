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
        
        # Calculate key metrics
        total_products = len(range_pd)
        ct_count = len(range_pd[range_pd["CT"] == "X"])
        jeves_count = len(range_pd[range_pd["JEEVES"] == "X"])
        stibo_count = len(range_pd[range_pd["STIBO"] == "X"])
        all_three = len(range_pd[
            (range_pd["CT"] == "X") & 
            (range_pd["JEEVES"] == "X") & 
            (range_pd["STIBO"] == "X")
        ])
        problems_count = total_products - all_three
        
        # Visual Alert Banner for Problems
        if problems_count > 0:
            st.error(f"⚠️ **{problems_count} products have issues** (not present in all 3 sources)")
        else:
            st.success("✅ **All products are present in all 3 sources**")
        
        st.markdown("---")
        
        # Main visual metrics - Large and clear
        col1, col2, col3, col4, col5 = st.columns(5)
        
        with col1:
            st.metric("Total products", f"{total_products:,}")
        with col2:
            st.metric("✅ In all 3 sources", f"{all_three:,}", 
                     delta=f"{all_three/total_products*100:.1f}%" if total_products > 0 else "0%",
                     delta_color="normal")
        with col3:
            st.metric("⚠️ With issues", f"{problems_count:,}",
                     delta=f"{problems_count/total_products*100:.1f}%" if total_products > 0 else "0%",
                     delta_color="inverse")
        with col4:
            missing_ct = total_products - ct_count
            st.metric("Missing from CT", f"{missing_ct:,}", 
                     delta=f"-{missing_ct}" if missing_ct > 0 else None,
                     delta_color="inverse" if missing_ct > 0 else "off")
        with col5:
            missing_jeves = total_products - jeves_count
            missing_stibo = total_products - stibo_count
            st.metric("Missing from JEEVES/STIBO", f"{missing_jeves}/{missing_stibo}", 
                     delta=f"-{missing_jeves}/-{missing_stibo}" if (missing_jeves > 0 or missing_stibo > 0) else None,
                     delta_color="inverse" if (missing_jeves > 0 or missing_stibo > 0) else "off")
        
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
        
        # Detailed visualizations (below the quick overview) - Collapsible
        with st.expander("📊 Detailed Analysis", expanded=False):
            col_left, col_right = st.columns(2)
        
        with col_left:
            # Distribution chart - use filtered data but ensure correct calculations
            filtered_range_filled = filtered_range.fillna({'CT': '', 'JEEVES': '', 'STIBO': ''})
            
            # Count presence for each product
            presence_count = (
                (filtered_range_filled["CT"] == "X").astype(int) + 
                (filtered_range_filled["JEEVES"] == "X").astype(int) + 
                (filtered_range_filled["STIBO"] == "X").astype(int)
            )
            
            status_counts = {
                "✅ In all 3": len(filtered_range_filled[presence_count == 3]),
                "⚠️ In 2": len(filtered_range_filled[presence_count == 2]),
                "⚠️ In 1": len(filtered_range_filled[presence_count == 1]),
                "❌ In none": len(filtered_range_filled[presence_count == 0])
            }
            
            # Verify sum equals total
            total_calculated = sum(status_counts.values())
            if total_calculated != len(filtered_range_filled):
                st.warning(f"⚠️ Calculation mismatch: {total_calculated} vs {len(filtered_range_filled)}")
            
            fig_pie = px.pie(
                values=list(status_counts.values()),
                names=list(status_counts.keys()),
                title="Distribution by number of sources",
                color_discrete_map={
                    "✅ In all 3": "#28a745",
                    "⚠️ In 2": "#ffc107",
                    "⚠️ In 1": "#fd7e14",
                    "❌ In none": "#dc3545"
                }
            )
            st.plotly_chart(fig_pie, use_container_width=True, key="pie_range")
        
        with col_right:
            # Bar chart by source
            source_counts = {
                "CT": len(filtered_range[filtered_range["CT"] == "X"]),
                "JEEVES": len(filtered_range[filtered_range["JEEVES"] == "X"]),
                "STIBO": len(filtered_range[filtered_range["STIBO"] == "X"])
            }
            fig_bar = px.bar(
                x=list(source_counts.keys()),
                y=list(source_counts.values()),
                title="Number of products by source (filtered)",
                labels={"x": "Source", "y": "Number of products"},
                color=list(source_counts.keys()),
                color_discrete_map={
                    "CT": "#007bff",
                    "JEEVES": "#28a745",
                    "STIBO": "#17a2b8"
                }
            )
            st.plotly_chart(fig_bar, use_container_width=True, key="bar_range")
        
        # Filters
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
        
        # Download buttons
        col_dl1, col_dl2 = st.columns(2)
        
        with col_dl1:
            csv_range = filtered_range.to_csv(index=False)
            st.download_button(
                label="📥 Download all data",
                data=csv_range,
                file_name=f"range_reconciliation_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
                mime="text/csv",
                use_container_width=True
            )
        
        with col_dl2:
            # Filter products NOT in all 3 sources
            not_in_all_three = filtered_range[
                ~((filtered_range["CT"] == "X") & 
                  (filtered_range["JEEVES"] == "X") & 
                  (filtered_range["STIBO"] == "X"))
            ]
            csv_missing = not_in_all_three.to_csv(index=False)
            st.download_button(
                label="📥 Download Missing from Sources (CSV)",
                data=csv_missing,
                file_name=f"missing_from_sources_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
                mime="text/csv",
                use_container_width=True
            )
            st.caption(f"⚠️ {len(not_in_all_three)} products not in all 3 sources")
    
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
        
        # Histogram for presence patterns
        # Fill NaN for consistent comparison
        range_pd_filled = range_pd.fillna({'CT': '', 'JEEVES': '', 'STIBO': ''})
        
        # Count products by presence pattern
        presence_patterns = {
            "All 3": all_three,
            "CT + JEEVES": len(range_pd_filled[(range_pd_filled["CT"] == "X") & (range_pd_filled["JEEVES"] == "X") & (range_pd_filled["STIBO"] != "X")]),
            "CT + STIBO": len(range_pd_filled[(range_pd_filled["CT"] == "X") & (range_pd_filled["JEEVES"] != "X") & (range_pd_filled["STIBO"] == "X")]),
            "JEEVES + STIBO": len(range_pd_filled[(range_pd_filled["CT"] != "X") & (range_pd_filled["JEEVES"] == "X") & (range_pd_filled["STIBO"] == "X")]),
            "CT only": len(range_pd_filled[(range_pd_filled["CT"] == "X") & (range_pd_filled["JEEVES"] != "X") & (range_pd_filled["STIBO"] != "X")]),
            "JEEVES only": len(range_pd_filled[(range_pd_filled["CT"] != "X") & (range_pd_filled["JEEVES"] == "X") & (range_pd_filled["STIBO"] != "X")]),
            "STIBO only": len(range_pd_filled[(range_pd_filled["CT"] != "X") & (range_pd_filled["JEEVES"] != "X") & (range_pd_filled["STIBO"] == "X")]),
            "None": len(range_pd_filled[(range_pd_filled["CT"] != "X") & (range_pd_filled["JEEVES"] != "X") & (range_pd_filled["STIBO"] != "X")])
        }
        
        # Create horizontal bar chart (histogram)
        fig_hist = px.bar(
            x=list(presence_patterns.values()),
            y=list(presence_patterns.keys()),
            orientation='h',
            title="Product distribution by source combination",
            labels={"x": "Number of products", "y": "Source combination"},
            color=list(presence_patterns.keys()),
            color_discrete_map={
                "All 3": "#28a745",
                "CT + JEEVES": "#ffc107",
                "CT + STIBO": "#ffc107",
                "JEEVES + STIBO": "#ffc107",
                "CT only": "#fd7e14",
                "JEEVES only": "#fd7e14",
                "STIBO only": "#fd7e14",
                "None": "#dc3545"
            }
        )
        fig_hist.update_layout(showlegend=False, height=400)
        st.plotly_chart(fig_hist, use_container_width=True, key="hist_overview")
    

if __name__ == "__main__":
    main()
