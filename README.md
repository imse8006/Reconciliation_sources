# Reconciliation Sources - Ekofisk Products

Outil de réconciliation des produits entre différentes sources de données (CT Ekofisk, JEEVES, STIBO).

## Description

Ce projet permet de comparer et réconcilier les données produits entre trois sources principales :
- **CT** : Données Ekofisk depuis le fichier CT
- **JEEVES** : Données depuis la feuille 3-STIBO-TRACKER
- **STIBO** : Données STIBO

Le script génère un fichier Excel de réconciliation qui liste tous les produits uniques et indique leur présence dans chaque source.

## Structure du projet

```
Reconciliation_Ekofisk_JEEVES/
├── JEEVES/
│   └── RECONC Product Data 2026-02-04.xlsx
├── CT/
│   └── P1 Data Cleansing - Product Ekofisk.xlsb
├── STIBO/
│   └── extract_stibo_all_products.xlsx
├── reconcile_products.py      # Script principal de réconciliation
├── app_streamlit.py            # Application Streamlit pour visualisation
└── README.md
```

## Prérequis

- Python 3.8+
- Bibliothèques Python :
  - `polars` (avec support Excel)
  - `openpyxl`
  - `pyxlsb`
  - `streamlit` (pour l'application web)
  - `plotly` (pour les graphiques)

## Installation

```bash
pip install polars[excel] openpyxl pyxlsb streamlit plotly
```

## Utilisation

### 1. Réconciliation des produits

Exécuter le script principal pour générer le fichier de réconciliation :

```bash
python reconcile_products.py
```

**Fichier généré** : `Range_Reconciliation_[timestamp].xlsx`

**Contenu du fichier** :
- `ProductCode` : Code produit unique (SUPC)
- `CT` : "X" si présent, vide si absent
- `JEEVES` : "X" si présent, vide si absent
- `STIBO` : "X" si présent, vide si absent
- `Absent_from` : Liste des sources où le produit est absent (ou "-" si présent partout)

### 2. Visualisation interactive (Streamlit)

Lancer l'application web pour visualiser les résultats :

```bash
streamlit run app_streamlit.py
```

L'application s'ouvre automatiquement dans votre navigateur à `http://localhost:8501`

**Fonctionnalités** :
- Vue d'ensemble avec statistiques
- Filtres par source (CT, JEEVES, STIBO)
- Recherche de produits
- Graphiques interactifs
- Export CSV

## Format des données sources

### JEEVES
- **Fichier** : `JEEVES/RECONC Product Data 2026-02-04.xlsx`
- **Feuille** : `3-STIBO-TRACKER`
- **Colonne produit** : `SUPC` (colonne A)
- **En-têtes** : Ligne 1
- **Données** : À partir de la ligne 2

### CT
- **Fichier** : `CT/P1 Data Cleansing - Product Ekofisk.xlsb`
- **Feuille** : `Item`
- **Colonne produit** : `SUPC` (colonne B)
- **En-têtes** : Ligne 6
- **Données** : À partir de la ligne 7, colonne B

### STIBO
- **Fichier** : `STIBO/extract_stibo_all_products.xlsx`
- **Colonne produit** : `SUPC` (colonne A)
- **En-têtes** : Ligne 1
- **Données** : À partir de la ligne 2

## Gestion des fichiers

Le script détecte automatiquement si les fichiers d'entrée ont changé :
- **Fichiers identiques** : Écrase le fichier de sortie existant
- **Fichiers modifiés** : Crée un nouveau fichier avec timestamp

Un fichier `.reconciliation_hash.json` est créé automatiquement pour suivre les changements.

## Exemple de sortie

| ProductCode | CT | JEEVES | STIBO | Absent_from |
|-------------|----|--------|-------|-------------|
| 205167      | X  | X      | X     | -           |
| 215455      | X  |        | X     | JEEVES      |
| 5021339     |    | X      |       | CT, STIBO   |

## Auteur

Développé pour la réconciliation des produits Ekofisk.

## Licence

Propriétaire - Usage interne uniquement.
