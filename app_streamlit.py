"""Streamlit application to visualize reconciliation results"""
import streamlit as st
import polars as pl
import pandas as pd
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

OUTPUT_DIR = Path("output")


@st.cache_data
def list_output_versions():
    """List available versions = subdirs of output/ (e.g. 2302, 0203), sorted newest first."""
    if not OUTPUT_DIR.exists():
        return []
    versions = [d.name for d in OUTPUT_DIR.iterdir() if d.is_dir()]
    return sorted(versions, reverse=True)


def _product_key_col():
    return "ProductCode"


def _load_product_for_version(version: str, market: str = "Ekofisk"):
    """Load Product sheet from Reconciliation_{market}.xlsx in output/{version}/. Returns (path, pl.DataFrame or None)."""
    path = OUTPUT_DIR / version / f"Reconciliation_{market}.xlsx"
    if not path.exists():
        return None, None
    try:
        df = pl.read_excel(path, sheet_name="Product", raise_if_empty=False)
        return (path, df) if df is not None and df.height > 0 else (None, None)
    except Exception:
        return None, None


def _load_sheet_for_version(version: str, market: str, sheet_name: str):
    """Load one sheet from Reconciliation_{market}.xlsx in output/{version}/."""
    path = OUTPUT_DIR / version / f"Reconciliation_{market}.xlsx"
    if not path.exists():
        return None
    try:
        df = pl.read_excel(path, sheet_name=sheet_name, raise_if_empty=False)
        return df if df is not None and df.height > 0 else None
    except Exception:
        return None


def _diff_codes(df_old, df_new, key_col: str):
    """Compare two dataframes by key_col. Returns (added, removed, unchanged_count)."""
    if df_old is None or df_old.height == 0:
        old_set = set()
    else:
        old_set = set(df_old[key_col].drop_nulls().cast(pl.Utf8).to_list())
    if df_new is None or df_new.height == 0:
        new_set = set()
    else:
        new_set = set(df_new[key_col].drop_nulls().cast(pl.Utf8).to_list())
    added = sorted(new_set - old_set)
    removed = sorted(old_set - new_set)
    unchanged = len(old_set & new_set)
    return added, removed, unchanged


@st.cache_data
def find_latest_reconciliation_file(legal_entity: str = "Ekofisk"):
    """Find the latest Reconciliation file and load its Product sheet. Prefer output/{version}/Reconciliation_{market}.xlsx, else root."""
    market = _legal_entity_to_market(legal_entity)
    versions = list_output_versions()
    for v in versions:
        path, df = _load_product_for_version(v, market)
        if df is not None:
            return path, df
    path = Path(f"Reconciliation_{market}.xlsx")
    if path.exists():
        try:
            df = pl.read_excel(path, sheet_name="Product", raise_if_empty=False)
            if df is not None and df.height > 0:
                return path, df
        except Exception:
            pass
    return None, None


def _legal_entity_to_market(legal_entity: str) -> str:
    """Map sidebar legal entity to reconciliation file name segment."""
    return {"Ekofisk": "Ekofisk", "Fresh Direct": "Fresh", "Classic Drinks": "Classic"}.get(
        legal_entity, "Ekofisk"
    )


@st.cache_data
def load_invoice_ordering_reconciliation(legal_entity: str, focus: str, version: str | None = None):
    """Load Reconciliation_{market}.xlsx. focus='Vendor' -> Vendor Invoice + Vendor OS; focus='Customer' -> Customer Invoice + Customer OS.
    If version is set (e.g. 2302), load from output/{version}/; else prefer latest output version, then repo root."""
    market = _legal_entity_to_market(legal_entity)
    if version:
        path = OUTPUT_DIR / version / f"Reconciliation_{market}.xlsx"
    else:
        versions = list_output_versions()
        if versions:
            path = OUTPUT_DIR / versions[0] / f"Reconciliation_{market}.xlsx"
        else:
            path = Path(f"Reconciliation_{market}.xlsx")
    if not path.exists() and not version:
        path = Path(f"Reconciliation_{market}.xlsx")
    if not path.exists():
        return None, None
    inv_sheet = "Vendor Invoice" if focus == "Vendor" else "Customer Invoice"
    ord_sheet = "Vendor OS" if focus == "Vendor" else "Customer OS"
    try:
        invoice = pl.read_excel(path, sheet_name=inv_sheet, raise_if_empty=False)
    except Exception:
        invoice = None
    try:
        ordering = pl.read_excel(path, sheet_name=ord_sheet, raise_if_empty=False)
    except Exception:
        ordering = pl.DataFrame()
    if invoice is None or (hasattr(invoice, "height") and invoice.height == 0):
        invoice = None
    if ordering is None or (hasattr(ordering, "height") and ordering.height == 0):
        ordering = pl.DataFrame()
    if invoice is None and (ordering is None or ordering.height == 0):
        return None, None
    return (invoice if invoice is not None else pl.DataFrame({})), ordering


def _render_range_like_tab(pd_df, key_col: str, source_cols: list, tab_name: str, focus: str, key_suffix: str):
    """Same presentation as Product Range Reconciliation: alert, 4 metrics, filters, search, Detailed Analysis, table, downloads."""
    if key_col not in pd_df.columns or not source_cols:
        st.warning("Colonnes manquantes.")
        return
    pd_df = pd_df.fillna({c: "" for c in source_cols})
    total = len(pd_df)
    n_sources = len(source_cols)
    count_per_src = [len(pd_df[pd_df[c] == "X"]) for c in source_cols]
    mask_all = pd_df[source_cols[0]] == "X"
    for c in source_cols[1:]:
        mask_all = mask_all & (pd_df[c] == "X")
    in_all_sources = mask_all.sum()
    problems_count = total - in_all_sources

    st.header(f"{tab_name} Reconciliation")
    st.markdown("List of all codes with their presence in STIBO, CT and JEEVES")
    if problems_count > 0:
        st.error(f"⚠️ **{problems_count} codes have issues** (not present in all {n_sources} sources)")
    else:
        st.success(f"✅ **All codes are present in all {n_sources} sources**")
    st.markdown("---")

    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.metric("Total codes", f"{total:,}")
    with col2:
        st.metric(f"✅ In all {n_sources} sources", f"{in_all_sources:,}",
                  delta=f"{in_all_sources/total*100:.1f}%" if total > 0 else "0%", delta_color="normal")
    with col3:
        st.metric("⚠️ With issues", f"{problems_count:,}",
                  delta=f"{problems_count/total*100:.1f}%" if total > 0 else "0%", delta_color="inverse")
    with col4:
        missing_str = "/".join(str(total - c) for c in count_per_src)
        st.metric("Missing from STIBO/CT/JEEVES", missing_str)
    st.markdown("---")

    f1, f2, f3 = st.columns(3)
    with f1:
        filter_s1 = st.selectbox("STIBO", ["All", "X Present", "Absent"], index=0, key=f"stibo_{key_suffix}")
    with f2:
        filter_s2 = st.selectbox("CT", ["All", "X Present", "Absent"], index=0, key=f"ct_{key_suffix}")
    with f3:
        filter_s3 = st.selectbox("JEEVES", ["All", "X Present", "Absent"], index=0, key=f"jeves_{key_suffix}")
    search_term = st.text_input("🔍 Search code", "", key=f"search_{key_suffix}")

    filtered = pd_df.copy()
    if filter_s1 != "All":
        filtered = filtered[filtered[source_cols[0]] == ("X" if filter_s1 == "X Present" else "")]
    if filter_s2 != "All":
        filtered = filtered[filtered[source_cols[1]] == ("X" if filter_s2 == "X Present" else "")]
    if filter_s3 != "All":
        filtered = filtered[filtered[source_cols[2]] == ("X" if filter_s3 == "X Present" else "")]
    if search_term:
        filtered = filtered[filtered[key_col].astype(str).str.contains(search_term, case=False, na=False)]

    with st.expander("📊 Detailed Analysis", expanded=False):
        col_left, col_right = st.columns(2)
        with col_left:
            presence_count = (filtered[source_cols[0]] == "X").astype(int)
            for c in source_cols[1:]:
                presence_count = presence_count + (filtered[c] == "X").astype(int)
            status_counts = {
                f"✅ In all {n_sources}": len(filtered[presence_count == n_sources]),
                **{f"⚠️ In {k}": len(filtered[presence_count == k]) for k in range(n_sources - 1, 0, -1)},
                "❌ In none": len(filtered[presence_count == 0])
            }
            fig_pie = px.pie(
                values=list(status_counts.values()),
                names=list(status_counts.keys()),
                title="Distribution by number of sources",
                color_discrete_map={f"✅ In all {n_sources}": "#28a745", "⚠️ In 2": "#ffc107", "⚠️ In 1": "#fd7e14", "❌ In none": "#dc3545"}
            )
            st.plotly_chart(fig_pie, width="stretch", key=f"pie_{key_suffix}")
        with col_right:
            source_labels_short = ["STIBO", "CT", "JEEVES"][: len(source_cols)]
            source_counts = {lbl: len(filtered[filtered[c] == "X"]) for lbl, c in zip(source_labels_short, source_cols)}
            fig_bar = px.bar(
                x=list(source_counts.keys()),
                y=list(source_counts.values()),
                title="Number of codes by source (filtered)",
                labels={"x": "Source", "y": "Number of codes"},
                color=list(source_counts.keys()),
                color_discrete_map={"STIBO": "#17a2b8", "CT": "#007bff", "JEEVES": "#28a745"}
            )
            st.plotly_chart(fig_bar, width="stretch", key=f"bar_{key_suffix}")

    st.subheader("Detailed Data")
    source_labels = ["STIBO", "CT", "JEEVES"][: len(source_cols)]
    absent = filtered.apply(
        lambda r: ", ".join(n for n, c in zip(source_labels, source_cols) if r.get(c) != "X"),
        axis=1
    )
    filtered_display = filtered.copy()
    filtered_display["Absent_from"] = absent
    display_cols = [key_col] + source_cols + ["Absent_from"]
    filtered_sorted = filtered_display.sort_values(by="Absent_from", ascending=False, na_position="last")
    st.dataframe(filtered_sorted[display_cols], width="stretch", height=400)

    col_dl1, col_dl2 = st.columns(2)
    with col_dl1:
        st.download_button(
            label="📥 Download all data",
            data=filtered.to_csv(index=False),
            file_name=f"reconciliation_{focus}_{tab_name}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
            mime="text/csv",
            use_container_width=True,  # download_button has no width param
            key=f"dl_all_{key_suffix}"
        )
    with col_dl2:
        mask_all = filtered[source_cols[0]] == "X"
        for c in source_cols[1:]:
            mask_all = mask_all & (filtered[c] == "X")
        not_in_all = filtered[~mask_all]
        st.download_button(
            label="📥 Download Missing from Sources (CSV)",
            data=not_in_all.to_csv(index=False),
            file_name=f"missing_{focus}_{tab_name}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
            mime="text/csv",
            use_container_width=True,  # download_button
            key=f"dl_missing_{key_suffix}"
        )
        st.caption(f"⚠️ {len(not_in_all)} codes not in all {n_sources} sources")


def show_vendor_customer_reconciliation(legal_entity: str, focus: str):
    """Display Vendor or Customer reconciliation – same presentation as Product (alert, metrics, filters, charts, table)."""
    market = _legal_entity_to_market(legal_entity)
    title = "Vendor" if focus == "Vendor" else "Customer"
    st.title(f"{legal_entity} {title} Reconciliation - STIBO vs CT vs JEEVES")
    st.markdown("Invoice & Ordering-Shipping")
    st.markdown("---")

    invoice_df, ordering_df = load_invoice_ordering_reconciliation(legal_entity, focus)
    no_invoice = invoice_df is None or (hasattr(invoice_df, "height") and invoice_df.height == 0)
    no_ordering = ordering_df is None or (hasattr(ordering_df, "height") and ordering_df.height == 0)
    if no_invoice and no_ordering:
        st.warning("⚠️ No reconciliation data for this legal entity.")
        st.info(f"Run: `python run_reconciliation.py --market {market.lower()}` and ensure `Reconciliation_{market}.xlsx` exists.")
        return

    source_cols = ["STIBO_Vendor", "CT_Vendor", "JEEVES_Vendor"] if focus == "Vendor" else ["STIBO_Customer", "CT_Customer", "JEEVES_Customer"]
    key_col = "Code"

    tab_inv, tab_ord = st.tabs(["Invoice Reconciliation", "Ordering-Shipping Reconciliation"])

    with tab_inv:
        if no_invoice:
            st.info("No data for Invoice.")
        else:
            pd_inv = invoice_df.to_pandas()
            cols_inv = [c for c in source_cols if c in pd_inv.columns]
            if cols_inv and key_col in pd_inv.columns:
                _render_range_like_tab(pd_inv, key_col, cols_inv, "Invoice", focus, f"{focus}_inv")
            else:
                st.dataframe(pd_inv, width="stretch", height=400)

    with tab_ord:
        if no_ordering:
            st.info("No data for Ordering-Shipping.")
        else:
            pd_ord = ordering_df.to_pandas()
            cols_ord = [c for c in source_cols if c in pd_ord.columns]
            if cols_ord and key_col in pd_ord.columns:
                _render_range_like_tab(pd_ord, key_col, cols_ord, "Ordering-Shipping", focus, f"{focus}_ord")
            else:
                st.dataframe(pd_ord, width="stretch", height=400)

def show_history():
    """History: compare two output versions (output/{date}/) and show differences."""
    st.title("📜 History – Compare versions")
    st.markdown("Compare two runs from **output/** to see what changed (added / removed codes).")
    st.markdown("---")

    versions = list_output_versions()
    if not versions:
        st.warning("No versions found in **output/**. Run reconciliation with `--date 2302` (etc.) to create `output/2302/`.")
        return

    col_a, col_b, col_market, col_type = st.columns([1, 1, 1, 1])
    with col_a:
        version_old = st.selectbox("Version (older)", versions, index=min(1, len(versions) - 1), key="hist_old")
    with col_b:
        version_new = st.selectbox("Version (newer)", versions, index=0, key="hist_new")
    with col_market:
        market = st.selectbox("Market", ["Ekofisk", "Fresh", "Classic"], index=0, key="hist_market")
    with col_type:
        rec_type = st.selectbox(
            "Reconciliation type",
            ["Product", "Vendor Invoice", "Vendor OS", "Customer Invoice", "Customer OS"],
            index=0,
            key="hist_type",
        )

    if version_old == version_new:
        st.info("Choose two different versions to compare.")
        return

    key_col = _product_key_col() if rec_type == "Product" else "Code"
    df_old = None
    df_new = None

    if rec_type == "Product":
        _, df_old = _load_product_for_version(version_old, market)
        _, df_new = _load_product_for_version(version_new, market)
    else:
        sheet_map = {
            "Vendor Invoice": "Vendor Invoice",
            "Vendor OS": "Vendor OS",
            "Customer Invoice": "Customer Invoice",
            "Customer OS": "Customer OS",
        }
        sheet = sheet_map[rec_type]
        df_old = _load_sheet_for_version(version_old, market, sheet)
        df_new = _load_sheet_for_version(version_new, market, sheet)

    if df_old is None and df_new is None:
        st.warning(f"No data for **{rec_type}** in the selected versions.")
        return
    if key_col not in (df_old.columns if df_old is not None else []) and key_col not in (df_new.columns if df_new is not None else []):
        st.warning(f"Column **{key_col}** not found in the data.")
        return

    added, removed, unchanged = _diff_codes(df_old, df_new, key_col)

    st.subheader(f"Diff: **{version_old}** → **{version_new}**")
    c1, c2, c3 = st.columns(3)
    with c1:
        st.metric("➕ Added (in newer only)", len(added))
    with c2:
        st.metric("➖ Removed (in older only)", len(removed))
    with c3:
        st.metric("⏺ Unchanged", unchanged)
    st.markdown("---")

    tab_added, tab_removed = st.tabs(["➕ Added", "➖ Removed"])
    with tab_added:
        if added:
            st.dataframe(pd.DataFrame({key_col: added}), width="stretch", height=300)
            st.download_button(
                "📥 Download added (CSV)",
                key_col + "\n" + "\n".join(added),
                file_name=f"added_{version_old}_to_{version_new}_{rec_type.replace(' ', '_')}.csv",
                mime="text/csv",
                key="dl_added",
            )
        else:
            st.caption("No codes added.")
    with tab_removed:
        if removed:
            st.dataframe(pd.DataFrame({key_col: removed}), width="stretch", height=300)
            st.download_button(
                "📥 Download removed (CSV)",
                key_col + "\n" + "\n".join(removed),
                file_name=f"removed_{version_old}_to_{version_new}_{rec_type.replace(' ', '_')}.csv",
                mime="text/csv",
                key="dl_removed",
            )
        else:
            st.caption("No codes removed.")


def main():
    # Sidebar
    with st.sidebar:
        st.markdown("### Legal Entity")
        legal_entity = st.radio(
            "Legal Entity",
            ["Ekofisk", "Fresh Direct", "Classic Drinks"],
            index=0,
            key="legal_entity_selector",
            label_visibility="collapsed",
        )
        st.markdown("---")
        st.markdown("### Domain")
        domain = st.radio(
            "Domain",
            ["Product", "Vendor", "Customer", "History"],
            index=0,
            key="domain_selector",
            label_visibility="collapsed",
        )
        st.markdown("---")

    if domain == "History":
        show_history()
        return

    if legal_entity == "Ekofisk":
        if domain == "Product":
            show_product_reconciliation("Ekofisk")
        elif domain == "Vendor":
            show_vendor_customer_reconciliation("Ekofisk", "Vendor")
        elif domain == "Customer":
            show_vendor_customer_reconciliation("Ekofisk", "Customer")
    elif legal_entity == "Fresh Direct":
        if domain == "Product":
            st.title("Fresh Direct Product Reconciliation")
            st.info("🚧 Fresh Direct Product reconciliation coming soon...")
        elif domain == "Vendor":
            show_vendor_customer_reconciliation("Fresh Direct", "Vendor")
        elif domain == "Customer":
            show_vendor_customer_reconciliation("Fresh Direct", "Customer")
    elif legal_entity == "Classic Drinks":
        if domain == "Product":
            st.title("Classic Drinks Product Reconciliation")
            st.info("🚧 Classic Drinks Product reconciliation coming soon...")
        elif domain == "Vendor":
            show_vendor_customer_reconciliation("Classic Drinks", "Vendor")
        elif domain == "Customer":
            show_vendor_customer_reconciliation("Classic Drinks", "Customer")

def show_product_reconciliation(source_name="Ekofisk"):
    """Display Product reconciliation view"""
    st.title(f"{source_name} Product Reconciliation - JEEVES vs CT vs STIBO")
    st.markdown("---")
    
    # Load data automatically from repo
    repo_file, repo_df = find_latest_reconciliation_file(source_name)
    
    if repo_df is None or repo_file is None:
        st.warning("⚠️ No reconciliation data available")
        st.info("Run reconciliation to generate `output/{date}/Reconciliation_{market}.xlsx` (or place the file in the project root).")
        return
    
    range_df = repo_df
    
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
        col1, col2, col3, col4 = st.columns(4)
        
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
            missing_jeves = total_products - jeves_count
            missing_stibo = total_products - stibo_count
            st.metric("Missing from CT/JEEVES/STIBO", f"{missing_ct}/{missing_jeves}/{missing_stibo}")
        
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
            st.plotly_chart(fig_pie, width="stretch", key="pie_range")
        
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
            st.plotly_chart(fig_bar, width="stretch", key="bar_range")
        
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
        st.plotly_chart(fig_hist, width="stretch", key="hist_overview")
    

if __name__ == "__main__":
    main()
