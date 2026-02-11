"""Application Streamlit pour visualiser les résultats de réconciliation"""
import streamlit as st
import polars as pl
import plotly.express as px
import plotly.graph_objects as go
from pathlib import Path
import glob
from datetime import datetime

# Configuration de la page
st.set_page_config(
    page_title="Réconciliation Produits Ekofisk",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Titre principal
st.title("📊 Réconciliation Produits Ekofisk - JEEVES vs CT")
st.markdown("---")

@st.cache_data
def load_latest_analysis_files():
    """Charge le fichier Range Reconciliation le plus récent"""
    # Chercher le fichier Range Reconciliation le plus récent (par date de modification)
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
        st.error(f"Erreur lors du chargement du fichier: {e}")
        return None, None

def main():
    # Sidebar pour les options
    with st.sidebar:
        st.header("⚙️ Options")
        st.markdown("---")
        
        # Option pour régénérer les analyses
        if st.button("🔄 Régénérer les analyses", use_container_width=True):
            st.info("Exécution de la réconciliation...")
            import subprocess
            import sys
            result = subprocess.run([sys.executable, "reconcile_products.py"], 
                                  capture_output=True, text=True)
            if result.returncode == 0:
                st.success("Analyses régénérées avec succès!")
                st.cache_data.clear()
            else:
                st.error(f"Erreur: {result.stderr}")
    
    # Charger les données
    range_df, range_file = load_latest_analysis_files()
    
    if range_df is None:
        st.warning("⚠️ Aucun fichier Range Reconciliation trouvé. Veuillez d'abord exécuter `reconcile_products.py`")
        if st.button("Exécuter la réconciliation maintenant"):
            import subprocess
            import sys
            with st.spinner("Exécution en cours..."):
                result = subprocess.run([sys.executable, "reconcile_products.py"], 
                                      capture_output=True, text=True)
                if result.returncode == 0:
                    st.success("Réconciliation générée!")
                    st.cache_data.clear()
                    st.rerun()
                else:
                    st.error(f"Erreur: {result.stderr}")
        return
    
    # Afficher le fichier chargé
    st.sidebar.markdown("---")
    st.sidebar.markdown("### 📁 Fichier chargé")
    if range_file:
        st.sidebar.caption(f"Range Recon: {Path(range_file).name}")
    
    # Convertir en pandas pour Streamlit (plus facile pour l'affichage)
    range_pd = range_df.to_pandas()
    
    # Vérifier que les colonnes attendues existent
    required_cols = ["ProductCode", "CT", "JEEVES", "STIBO", "Absent_from"]
    missing_cols = [col for col in required_cols if col not in range_pd.columns]
    
    if missing_cols:
        st.error(f"Colonnes manquantes: {missing_cols}")
        st.info(f"Colonnes disponibles: {list(range_pd.columns)}")
        return
    
    # Onglets pour les différentes vues
    tab_range, tab_overview = st.tabs(["✅ Range Reconciliation", "📈 Vue d'ensemble"])
    
    with tab_range:
        st.header("✅ Range Reconciliation")
        st.markdown("Liste de tous les produits avec leur présence dans CT, JEEVES et STIBO")
        
        # Statistiques
        col1, col2, col3, col4 = st.columns(4)
        total_products = len(range_pd)
        ct_count = len(range_pd[range_pd["CT"] == "X"])
        jeves_count = len(range_pd[range_pd["JEEVES"] == "X"])
        stibo_count = len(range_pd[range_pd["STIBO"] == "X"])
        
        with col1:
            st.metric("Total produits", f"{total_products:,}")
        with col2:
            st.metric("Dans CT", f"{ct_count:,}")
        with col3:
            st.metric("Dans JEEVES", f"{jeves_count:,}")
        with col4:
            st.metric("Dans STIBO", f"{stibo_count:,}")
        
        st.markdown("---")
        
        # Filtres
        col1, col2, col3 = st.columns(3)
        with col1:
            filter_ct = st.selectbox("CT", ["All", "X Present", "Absent"], index=0)
        with col2:
            filter_jeves = st.selectbox("JEEVES", ["All", "X Present", "Absent"], index=0)
        with col3:
            filter_stibo = st.selectbox("STIBO", ["All", "X Present", "Absent"], index=0)
        
        # Recherche
        search_term = st.text_input("🔍 Rechercher un produit (code)", "")
        
        # Appliquer les filtres
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
        
        st.info(f"📊 {len(filtered_range)} produits affichés sur {total_products} au total")
        
        # Visualisations
        col_left, col_right = st.columns(2)
        
        with col_left:
            # Graphique de répartition
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
                title="Répartition par nombre de sources"
            )
            st.plotly_chart(fig_pie, use_container_width=True)
        
        with col_right:
            # Graphique en barres par source
            ct_count = len(range_pd[range_pd["CT"] == "X"])
            jeves_count = len(range_pd[range_pd["JEEVES"] == "X"])
            stibo_count = len(range_pd[range_pd["STIBO"] == "X"])
            source_counts = {
                "CT": ct_count,
                "JEEVES": jeves_count,
                "STIBO": stibo_count
            }
            fig_bar = px.bar(
                x=list(source_counts.keys()),
                y=list(source_counts.values()),
                title="Nombre de produits par source",
                labels={"x": "Source", "y": "Nombre de produits"}
            )
            st.plotly_chart(fig_bar, use_container_width=True)
        
        # Tableau de données
        st.subheader("Données détaillées")
        
        # Options d'affichage
        num_rows = st.slider("Nombre de lignes à afficher", 10, min(1000, len(filtered_range)), 100)
        
        st.dataframe(
            filtered_range[["ProductCode", "CT", "JEEVES", "STIBO", "Absent_from"]],
            use_container_width=True,
            height=400
        )
        
        # Téléchargement
        csv_range = filtered_range.to_csv(index=False)
        st.download_button(
            label="📥 Télécharger la Range Reconciliation (CSV)",
            data=csv_range,
            file_name=f"range_reconciliation_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
            mime="text/csv"
        )
    
    with tab_overview:
        st.header("Vue d'ensemble")
        
        # Métriques principales
        col1, col2, col3, col4 = st.columns(4)
        
        total_products = len(presence_pd)
        jeves_only = len(presence_pd[presence_pd["Statut"] == "JEEVES uniquement"])
        ct_only = len(presence_pd[presence_pd["Statut"] == "CT uniquement"])
        both = len(presence_pd[presence_pd["Statut"] == "Les deux"])
        
        with col1:
            st.metric("Total produits uniques", f"{total_products:,}")
        with col2:
            st.metric("Dans les deux sources", f"{both:,}", 
                     delta=f"{both/total_products*100:.1f}%")
        with col3:
            st.metric("JEEVES uniquement", f"{jeves_only:,}")
        with col4:
            st.metric("CT uniquement", f"{ct_only:,}")
        
        st.markdown("---")
        
        # Graphique en camembert pour la présence
        col_left, col_right = st.columns(2)
        
        with col_left:
            fig_pie = px.pie(
                presence_pd,
                names="Statut",
                title="Répartition des produits par source",
                color_discrete_map={
                    "Les deux": "#2ecc71",
                    "JEEVES uniquement": "#3498db",
                    "CT uniquement": "#e74c3c"
                }
            )
            fig_pie.update_traces(textposition='inside', textinfo='percent+label')
            st.plotly_chart(fig_pie, use_container_width=True)
        
        with col_right:
            # Graphique en barres
            status_counts = presence_pd["Statut"].value_counts()
            fig_bar = px.bar(
                x=status_counts.index,
                y=status_counts.values,
                title="Nombre de produits par statut",
                labels={"x": "Statut", "y": "Nombre de produits"},
                color=status_counts.index,
                color_discrete_map={
                    "Les deux": "#2ecc71",
                    "JEEVES uniquement": "#3498db",
                    "CT uniquement": "#e74c3c"
                }
            )
            fig_bar.update_layout(showlegend=False)
            st.plotly_chart(fig_bar, use_container_width=True)
        
    

if __name__ == "__main__":
    main()
