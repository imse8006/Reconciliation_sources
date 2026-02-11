# Reconciliation Sources - Ekofisk Products

Product reconciliation tool for comparing data across different sources (CT Ekofisk, JEEVES, STIBO).

## Description

This project enables comparison and reconciliation of product data across three main sources:
- **CT** : Ekofisk data from CT file
- **JEEVES** : Data from sheet 3-STIBO-TRACKER
- **STIBO** : STIBO data

The script generates an Excel reconciliation file that lists all unique products and indicates their presence in each source.

## Project Structure

```
Reconciliation_Ekofisk_JEEVES/
├── JEEVES/
│   └── RECONC Product Data 2026-02-04.xlsx
├── CT/
│   └── P1 Data Cleansing - Product Ekofisk.xlsb
├── STIBO/
│   └── extract_stibo_all_products.xlsx
├── reconcile_products.py      # Main reconciliation script
├── app_streamlit.py            # Streamlit visualization app
└── README.md
```

## Prerequisites

- Python 3.8+
- Python libraries:
  - `polars` (with Excel support)
  - `openpyxl`
  - `pyxlsb`
  - `streamlit` (for web application)
  - `plotly` (for charts)

## Installation

```bash
pip install polars[excel] openpyxl pyxlsb streamlit plotly
```

## Usage

### 1. Product Reconciliation

Run the main script to generate the reconciliation file:

```bash
python reconcile_products.py
```

**Generated file**: `Range_Reconciliation_[timestamp].xlsx`

**File contents**:
- `ProductCode`: Unique product code (SUPC)
- `CT`: "X" if present, empty if absent
- `JEEVES`: "X" if present, empty if absent
- `STIBO`: "X" if present, empty if absent
- `Absent_from`: List of sources where the product is absent (or "-" if present everywhere)

### 2. Interactive Visualization (Streamlit)

Launch the web application to visualize results:

```bash
streamlit run app_streamlit.py
```

The application opens automatically in your browser at `http://localhost:8501`

**Features**:
- Overview with statistics
- Filters by source (CT, JEEVES, STIBO)
- Product search
- Interactive charts
- CSV export

## Source Data Format

### JEEVES
- **File**: `JEEVES/RECONC Product Data 2026-02-04.xlsx`
- **Sheet**: `3-STIBO-TRACKER`
- **Product column**: `SUPC` (column A)
- **Headers**: Row 1
- **Data**: Starting from row 2

### CT
- **File**: `CT/P1 Data Cleansing - Product Ekofisk.xlsb`
- **Sheet**: `Item`
- **Product column**: `SUPC` (column B)
- **Headers**: Row 6
- **Data**: Starting from row 7, column B

### STIBO
- **File**: `STIBO/extract_stibo_all_products.xlsx`
- **Product column**: `SUPC` (column A)
- **Headers**: Row 1
- **Data**: Starting from row 2

## File Management

The script automatically detects if input files have changed:
- **Identical files**: Overwrites existing output file
- **Modified files**: Creates a new file with timestamp

A `.reconciliation_hash.json` file is automatically created to track changes.

## Output Example

| ProductCode | CT | JEEVES | STIBO | Absent_from |
|-------------|----|--------|-------|-------------|
| 205167      | X  | X      | X     | -           |
| 215455      | X  |        | X     | JEEVES      |
| 5021339     |    | X      |       | CT, STIBO   |

## Author

Developed for Ekofisk product reconciliation.

## License

Proprietary - Internal use only.
