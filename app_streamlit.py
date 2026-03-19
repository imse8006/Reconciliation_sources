"""Streamlit application to visualize reconciliation results"""
import streamlit as st
import polars as pl
import pandas as pd
import plotly.express as px
from pathlib import Path
from datetime import datetime

import market_config

st.set_page_config(
    page_title="Reconciliation",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

OUTPUT_DIR = Path("output")
_PRODUCT_NON_ERP_COLS = {"ProductCode", "CT", "STIBO", "Absent_from"}
_MONTHS = ["", "Jan", "Fév", "Mar", "Avr", "Mai", "Juin", "Juil", "Aoû", "Sep", "Oct", "Nov", "Déc"]


def _format_version(v: str) -> str:
    """'1803' → '18 Mar', '2302' → '23 Fév'"""
    try:
        day, month = int(v[:2]), int(v[2:])
        return f"{day} {_MONTHS[month]}"
    except Exception:
        return v


@st.cache_data
def list_output_versions() -> list[str]:
    """All version folders in output/, newest first."""
    if not OUTPUT_DIR.exists():
        return []
    return sorted([d.name for d in OUTPUT_DIR.iterdir() if d.is_dir()], reverse=True)


@st.cache_data
def _versions_for_market(market: str) -> list[str]:
    """Only versions that have a Reconciliation_{market}.xlsx file, newest first."""
    if not OUTPUT_DIR.exists():
        return []
    return sorted(
        [d.name for d in OUTPUT_DIR.iterdir()
         if d.is_dir() and (d / f"Reconciliation_{market}.xlsx").exists()],
        reverse=True,
    )


def _available_markets() -> list[str]:
    all_m = market_config.list_markets()
    if not OUTPUT_DIR.exists():
        return all_m
    found = {
        m for m in all_m
        for v_dir in OUTPUT_DIR.iterdir()
        if v_dir.is_dir() and (v_dir / f"Reconciliation_{m}.xlsx").exists()
    }
    return [m for m in all_m if m in found] or all_m


def _load_sheet(market: str, sheet: str, version: str | None = None) -> pl.DataFrame | None:
    if version:
        path = OUTPUT_DIR / version / f"Reconciliation_{market}.xlsx"
    else:
        for v in list_output_versions():
            p = OUTPUT_DIR / v / f"Reconciliation_{market}.xlsx"
            if p.exists():
                path = p
                break
        else:
            path = Path(f"Reconciliation_{market}.xlsx")
    if not path.exists():
        return None
    try:
        df = pl.read_excel(path, sheet_name=sheet, raise_if_empty=False)
        return df if df is not None and df.height > 0 else None
    except Exception:
        return None


def _detect_erp_col_product(columns: list[str]) -> str | None:
    for c in columns:
        if c not in _PRODUCT_NON_ERP_COLS:
            return c
    return None


def _detect_source_cols(columns: list[str], suffix: str) -> list[str]:
    cols = [c for c in columns if c.endswith(suffix)]
    return (
        [c for c in cols if c.startswith("STIBO")]
        + [c for c in cols if c.startswith("CT")]
        + [c for c in cols if not c.startswith("STIBO") and not c.startswith("CT")]
    )


# ─── Evolution (historique) ───────────────────────────────────────────────────

@st.cache_data
def _compute_product_evolution(market: str) -> pd.DataFrame:
    """Compute product metrics for each available version."""
    rows = []
    for v in sorted(list_output_versions()):  # chronological order
        df = _load_sheet(market, "Product", v)
        if df is None:
            continue
        pd_df = df.to_pandas()
        erp_col = _detect_erp_col_product(pd_df.columns.tolist())
        if erp_col is None:
            continue
        src_cols = ["CT", erp_col, "STIBO"]
        pd_df = pd_df.fillna({c: "" for c in src_cols})
        total = len(pd_df)
        mask_all = pd.Series([True] * total, index=pd_df.index)
        for c in src_cols:
            mask_all &= pd_df[c] == "X"
        in_all = int(mask_all.sum())
        rows.append({
            "Version": _format_version(v),
            "Version_raw": v,
            "Total produits": total,
            "Dans les 3 sources": in_all,
            "Avec écarts": total - in_all,
        })
    return pd.DataFrame(rows)


def _render_evolution_chart(market: str):
    evo = _compute_product_evolution(market)
    if evo.empty or len(evo) < 2:
        st.info("Pas assez de versions pour afficher l'évolution (minimum 2).")
        return

    fig = px.bar(
        evo,
        x="Version",
        y=["Dans les 3 sources", "Avec écarts"],
        barmode="stack",
        title=f"Évolution des produits — {market}",
        labels={"value": "Nombre de produits", "variable": "Statut"},
        color_discrete_map={"Dans les 3 sources": "#28a745", "Avec écarts": "#dc3545"},
    )
    fig.update_layout(legend=dict(orientation="h", yanchor="bottom", y=1.02))
    st.plotly_chart(fig, use_container_width=True)

    st.dataframe(
        evo.drop(columns=["Version_raw"]).set_index("Version"),
        use_container_width=True,
    )


# ─── Product reconciliation ───────────────────────────────────────────────────

def _render_product_tab(pd_df: pd.DataFrame, erp_col: str, market: str, version: str):
    source_cols = ["CT", erp_col, "STIBO"]
    key_col = "ProductCode"

    pd_df = pd_df.fillna({c: "" for c in source_cols + ["Absent_from"]})
    total = len(pd_df)
    counts = {c: int((pd_df[c] == "X").sum()) for c in source_cols}
    mask_all = pd.Series([True] * total, index=pd_df.index)
    for c in source_cols:
        mask_all &= pd_df[c] == "X"
    in_all = int(mask_all.sum())
    problems = total - in_all

    st.header("Range Reconciliation")

    if problems > 0:
        st.error(f"⚠️ **{problems} produits ont des écarts** (absents d'au moins une source)")
    else:
        st.success("✅ **Tous les produits sont présents dans les 3 sources**")
    st.markdown("---")

    col1, col2, col3, col4 = st.columns(4)
    col1.metric("Total produits", f"{total:,}")
    col2.metric("✅ Dans les 3 sources", f"{in_all:,}",
                delta=f"{in_all/total*100:.1f}%" if total else "0%", delta_color="normal")
    col3.metric("⚠️ Avec écarts", f"{problems:,}",
                delta=f"{problems/total*100:.1f}%" if total else "0%", delta_color="inverse")
    missing_str = "/".join(f"{total - counts[c]}" for c in source_cols)
    col4.metric(f"Absents CT/{erp_col}/STIBO", missing_str)
    st.markdown("---")

    fc1, fc2, fc3 = st.columns(3)
    f_ct    = fc1.selectbox("CT",    ["All", "X Present", "Absent"], key=f"f_ct_{market}_{version}")
    f_erp   = fc2.selectbox(erp_col, ["All", "X Present", "Absent"], key=f"f_erp_{market}_{version}")
    f_stibo = fc3.selectbox("STIBO", ["All", "X Present", "Absent"], key=f"f_stibo_{market}_{version}")
    search  = st.text_input("🔍 Rechercher un code produit", "", key=f"search_{market}_{version}")

    flt = pd_df.copy()
    for val, col in [(f_ct, "CT"), (f_erp, erp_col), (f_stibo, "STIBO")]:
        if val != "All":
            flt = flt[flt[col] == ("X" if val == "X Present" else "")]
    if search:
        flt = flt[flt[key_col].astype(str).str.contains(search, case=False, na=False)]

    with st.expander("📊 Analyse détaillée", expanded=False):
        cl, cr = st.columns(2)
        with cl:
            presence = sum((flt[c] == "X").astype(int) for c in source_cols)
            status = {
                "✅ Dans les 3": int((presence == 3).sum()),
                "⚠️ Dans 2":    int((presence == 2).sum()),
                "⚠️ Dans 1":    int((presence == 1).sum()),
                "❌ Dans aucune": int((presence == 0).sum()),
            }
            fig_pie = px.pie(
                values=list(status.values()), names=list(status.keys()),
                title="Distribution par nombre de sources",
                color_discrete_map={"✅ Dans les 3": "#28a745", "⚠️ Dans 2": "#ffc107",
                                     "⚠️ Dans 1": "#fd7e14", "❌ Dans aucune": "#dc3545"},
            )
            st.plotly_chart(fig_pie, use_container_width=True, key=f"pie_{market}_{version}")
        with cr:
            src_counts = {c: int((flt[c] == "X").sum()) for c in source_cols}
            fig_bar = px.bar(
                x=list(src_counts.keys()), y=list(src_counts.values()),
                title="Codes par source (filtré)",
                labels={"x": "Source", "y": "Nombre de produits"},
                color=list(src_counts.keys()),
            )
            st.plotly_chart(fig_bar, use_container_width=True, key=f"bar_{market}_{version}")

    st.subheader("Données détaillées")
    flt_sorted = flt.copy()
    flt_sorted["Absent_from"] = flt_sorted["Absent_from"].astype(str).replace({"nan": "", "None": ""})
    flt_sorted = flt_sorted.sort_values("Absent_from", ascending=False, na_position="last")
    st.dataframe(flt_sorted[[key_col, "CT", erp_col, "STIBO", "Absent_from"]],
                 use_container_width=True, height=400)

    dl1, dl2 = st.columns(2)
    dl1.download_button(
        "📥 Tout télécharger (CSV)", flt.to_csv(index=False),
        file_name=f"range_{market}_{version}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
        mime="text/csv", use_container_width=True, key=f"dl_all_{market}_{version}",
    )
    mask_all_flt = pd.Series([True] * len(flt), index=flt.index)
    for c in source_cols:
        mask_all_flt &= flt[c] == "X"
    not_in_all = flt[~mask_all_flt]
    dl2.download_button(
        "📥 Télécharger les écarts (CSV)", not_in_all.to_csv(index=False),
        file_name=f"ecarts_{market}_{version}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
        mime="text/csv", use_container_width=True, key=f"dl_missing_{market}_{version}",
    )
    dl2.caption(f"⚠️ {len(not_in_all)} produits absents d'au moins une source")


def show_product_reconciliation(market: str, version: str):
    erp = market_config.get_erp_name(market)
    st.title(f"{market} — Product Reconciliation — {_format_version(version)}")
    st.caption(f"CT / {erp} / STIBO  ·  version : {version}")

    df = _load_sheet(market, "Product", version)
    if df is None:
        st.warning("⚠️ Aucun fichier de réconciliation trouvé pour cette version.")
        st.info(f"Lance : `python run_reconciliation.py --market {market} --domains product --date {version}`")
        return

    pd_df = df.to_pandas()
    erp_col = _detect_erp_col_product(pd_df.columns.tolist())
    if erp_col is None:
        st.error(f"Colonne ERP non détectée. Colonnes disponibles : {list(pd_df.columns)}")
        return

    tab_range, tab_overview, tab_history = st.tabs(
        ["✅ Range Reconciliation", "📈 Vue d'ensemble", "📜 Historique"]
    )

    with tab_range:
        _render_product_tab(pd_df, erp_col, market, version)

    with tab_overview:
        st.header("Vue d'ensemble")
        src_cols = ["CT", erp_col, "STIBO"]
        pd_f = pd_df.fillna({c: "" for c in src_cols})
        total = len(pd_f)
        mask_all = pd.Series([True] * total, index=pd_f.index)
        for c in src_cols:
            mask_all &= pd_f[c] == "X"
        in_all = int(mask_all.sum())

        c1, c2, c3, c4, c5 = st.columns(5)
        c1.metric("Total produits", f"{total:,}")
        c2.metric("Dans les 3 sources", f"{in_all:,}",
                  delta=f"{in_all/total*100:.1f}%" if total else "0%")
        c3.metric(f"{src_cols[0]} only", str(len(pd_f[(pd_f[src_cols[0]]=="X")&(pd_f[src_cols[1]]!="X")&(pd_f[src_cols[2]]!="X")])))
        c4.metric(f"{src_cols[1]} only", str(len(pd_f[(pd_f[src_cols[0]]!="X")&(pd_f[src_cols[1]]=="X")&(pd_f[src_cols[2]]!="X")])))
        c5.metric("STIBO only",          str(len(pd_f[(pd_f[src_cols[0]]!="X")&(pd_f[src_cols[1]]!="X")&(pd_f[src_cols[2]]=="X")])))
        st.markdown("---")

        patterns = {
            "All 3": in_all,
            f"{src_cols[0]}+{src_cols[1]}": len(pd_f[(pd_f[src_cols[0]]=="X")&(pd_f[src_cols[1]]=="X")&(pd_f[src_cols[2]]!="X")]),
            f"{src_cols[0]}+STIBO":         len(pd_f[(pd_f[src_cols[0]]=="X")&(pd_f[src_cols[1]]!="X")&(pd_f[src_cols[2]]=="X")]),
            f"{src_cols[1]}+STIBO":         len(pd_f[(pd_f[src_cols[0]]!="X")&(pd_f[src_cols[1]]=="X")&(pd_f[src_cols[2]]=="X")]),
            f"{src_cols[0]} only":          len(pd_f[(pd_f[src_cols[0]]=="X")&(pd_f[src_cols[1]]!="X")&(pd_f[src_cols[2]]!="X")]),
            f"{src_cols[1]} only":          len(pd_f[(pd_f[src_cols[0]]!="X")&(pd_f[src_cols[1]]=="X")&(pd_f[src_cols[2]]!="X")]),
            "STIBO only":                   len(pd_f[(pd_f[src_cols[0]]!="X")&(pd_f[src_cols[1]]!="X")&(pd_f[src_cols[2]]=="X")]),
            "None":                         len(pd_f[(pd_f[src_cols[0]]!="X")&(pd_f[src_cols[1]]!="X")&(pd_f[src_cols[2]]!="X")]),
        }
        fig = px.bar(x=list(patterns.values()), y=list(patterns.keys()), orientation="h",
                     title="Distribution par combinaison de sources",
                     labels={"x": "Nombre de produits", "y": "Combinaison"},
                     color=list(patterns.keys()),
                     color_discrete_map={"All 3": "#28a745", "None": "#dc3545"})
        fig.update_layout(showlegend=False, height=400)
        st.plotly_chart(fig, use_container_width=True, key=f"overview_{market}_{version}")

    with tab_history:
        st.header(f"Évolution — {market}")
        _render_evolution_chart(market)


# ─── Vendor / Customer reconciliation ─────────────────────────────────────────

def _render_invoice_os_tab(pd_df: pd.DataFrame, source_cols: list[str], tab_name: str, key_suffix: str):
    key_col = "Code"
    if key_col not in pd_df.columns or not source_cols:
        st.warning("Colonnes manquantes.")
        return

    pd_df = pd_df.fillna({c: "" for c in source_cols})
    total = len(pd_df)
    mask_all = pd.Series([True] * total, index=pd_df.index)
    for c in source_cols:
        mask_all &= pd_df[c] == "X"
    in_all = int(mask_all.sum())
    problems = total - in_all
    src_labels = [c.rsplit("_", 1)[0] for c in source_cols]

    st.header(tab_name)
    if problems > 0:
        st.error(f"⚠️ **{problems} codes ont des écarts**")
    else:
        st.success(f"✅ **Tous les codes sont présents dans les {len(source_cols)} sources**")
    st.markdown("---")

    cols = st.columns(4)
    cols[0].metric("Total codes", f"{total:,}")
    cols[1].metric(f"✅ Dans les {len(source_cols)} sources", f"{in_all:,}",
                   delta=f"{in_all/total*100:.1f}%" if total else "0%", delta_color="normal")
    cols[2].metric("⚠️ Avec écarts", f"{problems:,}",
                   delta=f"{problems/total*100:.1f}%" if total else "0%", delta_color="inverse")
    missing_str = "/".join(f"{total - int((pd_df[c]=='X').sum())}" for c in source_cols)
    cols[3].metric("/".join(src_labels), missing_str)
    st.markdown("---")

    filter_cols = st.columns(len(source_cols))
    filters = {}
    for i, (c, lbl) in enumerate(zip(source_cols, src_labels)):
        filters[c] = filter_cols[i].selectbox(lbl, ["All", "X Present", "Absent"],
                                               key=f"{c}_{key_suffix}")
    search = st.text_input("🔍 Rechercher un code", "", key=f"search_{key_suffix}")

    flt = pd_df.copy()
    for c, val in filters.items():
        if val != "All":
            flt = flt[flt[c] == ("X" if val == "X Present" else "")]
    if search:
        flt = flt[flt[key_col].astype(str).str.contains(search, case=False, na=False)]

    st.subheader("Données détaillées")
    st.dataframe(flt[[key_col] + source_cols], use_container_width=True, height=400)

    dl1, dl2 = st.columns(2)
    dl1.download_button(
        "📥 Tout télécharger (CSV)", flt.to_csv(index=False),
        file_name=f"{tab_name.replace(' ', '_')}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
        mime="text/csv", use_container_width=True, key=f"dl_all_{key_suffix}",
    )
    mask_all_flt = pd.Series([True] * len(flt), index=flt.index)
    for c in source_cols:
        mask_all_flt &= flt[c] == "X"
    not_in_all = flt[~mask_all_flt]
    dl2.download_button(
        "📥 Télécharger les écarts (CSV)", not_in_all.to_csv(index=False),
        file_name=f"ecarts_{tab_name.replace(' ', '_')}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
        mime="text/csv", use_container_width=True, key=f"dl_missing_{key_suffix}",
    )
    dl2.caption(f"⚠️ {len(not_in_all)} codes absents d'au moins une source")


def show_vendor_customer_reconciliation(market: str, focus: str, version: str):
    erp = market_config.get_erp_name(market)
    st.title(f"{market} — {focus} Reconciliation — {_format_version(version)}")
    st.caption(f"STIBO / CT / {erp}  ·  version : {version}")

    suffix = "_Vendor" if focus == "Vendor" else "_Customer"
    inv_sheet = "Vendor Invoice" if focus == "Vendor" else "Customer Invoice"
    os_sheet  = "Vendor OS"      if focus == "Vendor" else "Customer OS"

    invoice_df = _load_sheet(market, inv_sheet, version)
    os_df      = _load_sheet(market, os_sheet,  version)

    if invoice_df is None and os_df is None:
        st.warning("⚠️ Aucune donnée de réconciliation pour ce marché / cette version.")
        st.info(f"Lance : `python run_reconciliation.py --market {market} --date {version}`")
        return

    tab_inv, tab_os = st.tabs(["Invoice Reconciliation", "Ordering-Shipping Reconciliation"])

    with tab_inv:
        if invoice_df is None:
            st.info("Pas de données Invoice.")
        else:
            pd_inv = invoice_df.to_pandas()
            src_cols = _detect_source_cols(pd_inv.columns.tolist(), suffix)
            if src_cols and "Code" in pd_inv.columns:
                _render_invoice_os_tab(pd_inv, src_cols, f"{focus} Invoice", f"{market}_{focus}_inv_{version}")
            else:
                st.dataframe(pd_inv, use_container_width=True, height=400)

    with tab_os:
        if os_df is None:
            st.info("Pas de données Ordering-Shipping.")
        else:
            pd_os = os_df.to_pandas()
            src_cols = _detect_source_cols(pd_os.columns.tolist(), suffix)
            if src_cols and "Code" in pd_os.columns:
                _render_invoice_os_tab(pd_os, src_cols, f"{focus} OS", f"{market}_{focus}_os_{version}")
            else:
                st.dataframe(pd_os, use_container_width=True, height=400)


# ─── History (diff entre deux versions) ───────────────────────────────────────

def show_history(market: str):
    st.title(f"📜 Historique — {market}")

    versions = list_output_versions()
    if not versions:
        st.warning("Aucune version trouvée dans **output/**.")
        return

    # Evolution chart
    st.subheader("Évolution des métriques produit")
    _render_evolution_chart(market)
    st.markdown("---")

    # Diff between two versions
    st.subheader("Comparer deux versions")
    if len(versions) < 2:
        st.info("Il faut au moins 2 versions pour comparer.")
        return

    col_a, col_b, col_t = st.columns(3)
    v_old    = col_a.selectbox("Version (ancienne)", versions, index=min(1, len(versions)-1), key="h_old")
    v_new    = col_b.selectbox("Version (récente)",  versions, index=0,                       key="h_new")
    rec_type = col_t.selectbox("Type", ["Product", "Vendor Invoice", "Vendor OS", "Customer Invoice", "Customer OS"], key="h_type")

    if v_old == v_new:
        st.info("Choisir deux versions différentes.")
        return

    key_col  = "ProductCode" if rec_type == "Product" else "Code"
    sheet    = "Product" if rec_type == "Product" else rec_type
    df_old   = _load_sheet(market, sheet, v_old)
    df_new   = _load_sheet(market, sheet, v_new)

    if df_old is None and df_new is None:
        st.warning(f"Aucune donnée pour **{rec_type}** dans les versions sélectionnées.")
        return

    def to_set(df):
        if df is None or key_col not in df.columns:
            return set()
        return set(df[key_col].drop_nulls().cast(pl.Utf8).to_list())

    old_set, new_set = to_set(df_old), to_set(df_new)
    added   = sorted(new_set - old_set)
    removed = sorted(old_set - new_set)
    unchanged = len(old_set & new_set)

    st.markdown(f"**{_format_version(v_old)}** → **{_format_version(v_new)}**")
    c1, c2, c3 = st.columns(3)
    c1.metric("➕ Ajoutés",    len(added))
    c2.metric("➖ Supprimés", len(removed))
    c3.metric("⏺ Inchangés", unchanged)
    st.markdown("---")

    tab_add, tab_rem = st.tabs(["➕ Ajoutés", "➖ Supprimés"])
    with tab_add:
        if added:
            st.dataframe(pd.DataFrame({key_col: added}), use_container_width=True, height=300)
            st.download_button("📥 Télécharger (CSV)", key_col + "\n" + "\n".join(added),
                               file_name=f"added_{v_old}_{v_new}.csv", mime="text/csv", key="dl_added")
        else:
            st.caption("Aucun code ajouté.")
    with tab_rem:
        if removed:
            st.dataframe(pd.DataFrame({key_col: removed}), use_container_width=True, height=300)
            st.download_button("📥 Télécharger (CSV)", key_col + "\n" + "\n".join(removed),
                               file_name=f"removed_{v_old}_{v_new}.csv", mime="text/csv", key="dl_removed")
        else:
            st.caption("Aucun code supprimé.")


# ─── Main ──────────────────────────────────────────────────────────────────────

def main():
    with st.sidebar:
        st.markdown("### Market")
        available = _available_markets()
        market = st.radio("Market", available, index=0, key="market_selector",
                          label_visibility="collapsed")

        st.markdown("---")
        st.markdown("### Domain")
        domain = st.radio("Domain", ["Product", "Vendor", "Customer", "Historique"],
                          index=0, key="domain_selector", label_visibility="collapsed")

        st.markdown("---")
        market_versions = _versions_for_market(market)
        if domain != "Historique" and market_versions:
            st.markdown("### Version")
            version = st.selectbox(
                "Version",
                market_versions,
                format_func=_format_version,
                index=0,
                key="version_selector",
                label_visibility="collapsed",
            )
        else:
            market_versions = list_output_versions()
            version = market_versions[0] if market_versions else "latest"

    if domain == "Historique":
        show_history(market)
    elif domain == "Product":
        show_product_reconciliation(market, version)
    elif domain == "Vendor":
        show_vendor_customer_reconciliation(market, "Vendor", version)
    elif domain == "Customer":
        show_vendor_customer_reconciliation(market, "Customer", version)


if __name__ == "__main__":
    main()
