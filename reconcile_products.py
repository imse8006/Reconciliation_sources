"""Product reconciliation script between JEEVES and Ekofisk CT"""
import polars as pl
from openpyxl import load_workbook
from pathlib import Path
from datetime import datetime
import hashlib
import json

def load_jeves_data(file_path: str) -> pl.DataFrame:
    """Load JEEVES Product data from sheet 3-STIBO-TRACKER
    Headers row 1, data from row 2+, column A = SUPC
    """
    wb = load_workbook(file_path, data_only=True)
    # Use sheet 3-STIBO-TRACKER
    ws = wb["3-STIBO-TRACKER"]
    
    # Read headers from row 1
    headers = [cell.value for cell in ws[1] if cell.value is not None]
    
    # Read data from row 2
    data = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if any(cell is not None for cell in row):
            data.append(row[:len(headers)])
    
    # Create DataFrame without strict schema to let Polars infer types
    if data:
        return pl.DataFrame(data, schema=headers, orient="row", infer_schema_length=None)
    else:
        return pl.DataFrame(schema=headers)

def load_ct_data(file_path: str) -> pl.DataFrame:
    """Load CT Ekofisk data
    Headers row 6, data starts at B7 (first SUPC)
    """
    from pyxlsb import open_workbook
    
    # Detect file format
    if file_path.endswith('.xlsb'):
        # Read .xlsb file
        with open_workbook(file_path) as wb:
            # Find sheet "Item" or use first sheet
            sheet_name = None
            for name in wb.sheets:
                if name.lower() == 'item':
                    sheet_name = name
                    break
            if sheet_name is None:
                sheet_name = wb.sheets[0]
            
            with wb.get_sheet(sheet_name) as sheet:
                rows = list(sheet.rows())
                
                # Read headers from row 6 (index 5), starting at column B (index 1)
                headers = []
                header_counts = {}  # To handle duplicates
                if len(rows) > 5:
                    header_row = rows[5]  # Row 6 (0-indexed = 5)
                    for idx, cell in enumerate(header_row):
                        if idx >= 1:  # Column B and following (SUPC is in column 2)
                            val = cell.v if cell.v is not None else None
                            if val is None:
                                header_name = f"Col_{idx+1}"
                            else:
                                # Convert to string and clean
                                header_name = str(val).strip() if isinstance(val, str) else str(val)
                            
                            # Handle duplicates
                            if header_name in header_counts:
                                header_counts[header_name] += 1
                                header_name = f"{header_name}_{header_counts[header_name]}"
                            else:
                                header_counts[header_name] = 0
                            
                            headers.append(header_name)
                
                # Read data from row 7 (index 6), column B (index 1)
                data = []
                for row_idx in range(6, len(rows)):
                    row = rows[row_idx]
                    # Check if column B (index 1, SUPC) has a numeric value
                    if len(row) > 1:
                        supc_val = row[1].v
                        # Accept only rows with valid SUPC (numeric)
                        if supc_val is not None and (isinstance(supc_val, (int, float)) or str(supc_val).strip()):
                            row_data = []
                            # Start at column B (index 1)
                            for col_idx in range(1, min(len(row), len(headers) + 1)):
                                cell_val = row[col_idx].v if col_idx < len(row) else None
                                row_data.append(cell_val)
                            # Complete with None if necessary
                            while len(row_data) < len(headers):
                                row_data.append(None)
                            data.append(row_data[:len(headers)])
                
                # Create DataFrame without strict schema to let Polars infer types
                if data:
                    return pl.DataFrame(data, schema=headers, orient="row", infer_schema_length=None)
                else:
                    return pl.DataFrame(schema=headers)
    else:
        # Use openpyxl for .xlsx files
        wb = load_workbook(file_path, data_only=True)
        # Find Product sheet or use active sheet
        ws = None
        for sheet_name in wb.sheetnames:
            if 'product' in sheet_name.lower():
                ws = wb[sheet_name]
                break
        if ws is None:
            ws = wb.active
        
        # Read headers from row 6, starting at column B
        headers = []
        for cell in ws[6]:
            if cell.column >= 2:  # Column B and following
                if cell.value is not None:
                    headers.append(cell.value)
                else:
                    headers.append(f"Col_{cell.column}")
        
        # Read data from row 7, column B
        data = []
        for row in ws.iter_rows(min_row=7, min_col=2, values_only=True):
            # Check if first column (B) has a value (SUPC)
            if row and row[0] is not None:
                data.append(row[:len(headers)])
        
        return pl.DataFrame(data, schema=headers, orient="row")

def load_stibo_data(file_path: str) -> pl.DataFrame:
    """Load STIBO Product data
    Headers row 1, data from row 2+
    """
    wb = load_workbook(file_path, data_only=True)
    ws = wb.active
    
    # Read headers from row 1
    headers = [cell.value for cell in ws[1] if cell.value is not None]
    
    # Read data from row 2
    data = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if any(cell is not None for cell in row):
            data.append(row[:len(headers)])
    
    return pl.DataFrame(data, schema=headers, orient="row")


def clean_product_code(value):
    """Clean a product code to remove .0 and normalize format"""
    if value is None:
        return None
    # Convert to string first
    str_val = str(value)
    # If it's a number with .0 at the end, convert to integer then string
    try:
        float_val = float(str_val)
        if float_val.is_integer():
            return str(int(float_val))
        return str_val
    except (ValueError, TypeError):
        return str_val.strip()

def create_range_reconciliation(jeves_df: pl.DataFrame, ct_df: pl.DataFrame, stibo_df: pl.DataFrame) -> pl.DataFrame:
    """Range Reconciliation: List all products with CT/JEEVES/STIBO columns and X marks"""
    # Identify product code column in each source
    # JEEVES: SUPC (from sheet 3-STIBO-TRACKER)
    # CT: SUPC
    # STIBO: SUPC
    
    # Find SUPC column in CT
    ct_product_col = None
    if "SUPC" in ct_df.columns:
        ct_product_col = "SUPC"
    else:
        ct_product_col = ct_df.columns[0]
    
    # Find SUPC column in STIBO
    stibo_product_col = "SUPC" if "SUPC" in stibo_df.columns else stibo_df.columns[0]
    
    # Create unique product lists from each source
    # Convert to string and clean format (remove .0)
    # JEEVES now uses SUPC from sheet 3-STIBO-TRACKER
    jeves_product_col = "SUPC" if "SUPC" in jeves_df.columns else jeves_df.columns[0]
    jeves_products = jeves_df.select([
        pl.col(jeves_product_col)
    ]).unique()
    
    ct_products = ct_df.select([
        pl.col(ct_product_col)
    ]).unique()
    
    stibo_products = stibo_df.select([
        pl.col(stibo_product_col)
    ]).unique()
    
    # Clean and convert to proper string
    def clean_and_convert(df, col_name):
        return df.with_columns([
            pl.col(col_name).map_elements(
                lambda x: clean_product_code(x),
                return_dtype=pl.Utf8
            ).alias("ProductCode")
        ]).select("ProductCode").unique()
    
    jeves_clean = clean_and_convert(jeves_products, jeves_product_col)
    ct_clean = clean_and_convert(ct_products, ct_product_col)
    stibo_clean = clean_and_convert(stibo_products, stibo_product_col)
    
    # Combine all unique products
    all_products = pl.concat([jeves_clean, ct_clean, stibo_clean]).unique("ProductCode")
    
    # Create lists for verification
    jeves_list = jeves_clean.to_series().to_list()
    ct_list = ct_clean.to_series().to_list()
    stibo_list = stibo_clean.to_series().to_list()
    
    # Create DataFrame with presence columns
    reconciliation = all_products.with_columns([
        pl.col("ProductCode").is_in(ct_list).alias("CT_present"),
        pl.col("ProductCode").is_in(jeves_list).alias("JEEVES_present"),
        pl.col("ProductCode").is_in(stibo_list).alias("STIBO_present")
    ]).with_columns([
        # Replace True with "X", False with empty string
        pl.when(pl.col("CT_present")).then(pl.lit("X")).otherwise(pl.lit("")).alias("CT"),
        pl.when(pl.col("JEEVES_present")).then(pl.lit("X")).otherwise(pl.lit("")).alias("JEEVES"),
        pl.when(pl.col("STIBO_present")).then(pl.lit("X")).otherwise(pl.lit("")).alias("STIBO")
    ]).with_columns([
        # Column summarizing absent sources (in English)
        pl.concat_str([
            pl.when(pl.col("CT_present") == False).then(pl.lit("CT")).otherwise(pl.lit("")),
            pl.when(pl.col("JEEVES_present") == False).then(pl.lit("JEEVES")).otherwise(pl.lit("")),
            pl.when(pl.col("STIBO_present") == False).then(pl.lit("STIBO")).otherwise(pl.lit(""))
        ], separator=", ").str.strip_chars_start(", ").str.strip_chars_end(", ").alias("Absent_from")
    ]).with_columns([
        # Replace empty strings with "-" for clarity
        pl.when(pl.col("Absent_from") == "").then(pl.lit("-")).otherwise(pl.col("Absent_from")).alias("Absent_from")
    ]).select([
        "ProductCode", "CT", "JEEVES", "STIBO", "Absent_from"
    ]).sort("ProductCode")
    
    return reconciliation

def get_file_hash(file_path: str) -> str:
    """Calculate MD5 hash of a file to detect changes"""
    hash_md5 = hashlib.md5()
    try:
        with open(file_path, "rb") as f:
            for chunk in iter(lambda: f.read(4096), b""):
                hash_md5.update(chunk)
        return hash_md5.hexdigest()
    except FileNotFoundError:
        return None

def get_input_files_hash() -> str:
    """Calculate combined hash of all input files"""
    input_files = [
        "JEEVES/RECONC Product Data 2026-02-04.xlsx",
        "CT/P1 Data Cleansing - Product Ekofisk.xlsb",
        "STIBO/extract_stibo_all_products.xlsx"
    ]
    
    hashes = []
    for file_path in input_files:
        file_hash = get_file_hash(file_path)
        if file_hash:
            hashes.append(f"{file_path}:{file_hash}")
        else:
            # If a file doesn't exist, return None to force creation of new file
            return None
    
    # Create combined hash
    combined = "|".join(hashes)
    return hashlib.md5(combined.encode()).hexdigest()

def find_existing_output_files() -> dict:
    """Find existing output files"""
    files = {}
    
    # Search for Range Reconciliation
    range_files = list(Path(".").glob("Range_Reconciliation_*.xlsx"))
    if range_files:
        files["range"] = max(range_files, key=lambda x: x.stat().st_mtime)
    
    return files

def save_hash_info(input_hash: str, output_file: Path):
    """Save input hash with output file name"""
    hash_file = Path(".reconciliation_hash.json")
    hash_info = {
        "input_hash": input_hash,
        "output_file": str(output_file)
    }
    with open(hash_file, "w") as f:
        json.dump(hash_info, f, indent=2)

def load_hash_info() -> dict:
    """Load previous hash"""
    hash_file = Path(".reconciliation_hash.json")
    if hash_file.exists():
        try:
            with open(hash_file, "r") as f:
                return json.load(f)
        except:
            return None
    return None

def main():
    print("Loading data...")
    
    # Check if input files have changed
    current_input_hash = get_input_files_hash()
    previous_hash_info = load_hash_info()
    
    # Determine output file name
    if current_input_hash and previous_hash_info and previous_hash_info.get("input_hash") == current_input_hash:
        # Input files are identical, overwrite existing file
        existing_files = find_existing_output_files()
        output_file_range = existing_files.get("range", Path("Range_Reconciliation.xlsx"))
        print("[INFO] Input files identical - overwriting existing file")
    else:
        # Files have changed or first run, create new file with timestamp
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_file_range = Path(f"Range_Reconciliation_{timestamp}.xlsx")
        if current_input_hash:
            print("[INFO] Input files modified - creating new file")
        else:
            print("[WARN] Some input files missing - creating new file")
    
    # Load data
    jeves_df = load_jeves_data("JEEVES/RECONC Product Data 2026-02-04.xlsx")
    print(f"JEEVES: {len(jeves_df)} products loaded")
    
    ct_df = load_ct_data("CT/P1 Data Cleansing - Product Ekofisk.xlsb")
    print(f"CT: {len(ct_df)} products loaded")
    
    stibo_df = load_stibo_data("STIBO/extract_stibo_all_products.xlsx")
    print(f"STIBO: {len(stibo_df)} products loaded")
    
    # Range Reconciliation: List all products with CT/JEEVES/STIBO
    print("\nCreating Range Reconciliation...")
    range_reconciliation = create_range_reconciliation(jeves_df, ct_df, stibo_df)
    print(f"Total unique products: {len(range_reconciliation)}")
    
    # Generate Excel file
    print(f"\nGenerating Excel file...")
    
    # Range Reconciliation file
    range_reconciliation.write_excel(output_file_range)
    print(f"[OK] {output_file_range}")
    
    # Save hash for next run
    if current_input_hash:
        save_hash_info(current_input_hash, output_file_range)
    
    print("\nDone!")

if __name__ == "__main__":
    main()
