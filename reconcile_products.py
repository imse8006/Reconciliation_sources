"""Script de réconciliation des produits entre JEEVES et Ekofisk CT"""
import polars as pl
from openpyxl import load_workbook
from pathlib import Path
from datetime import datetime
import hashlib
import json

def load_jeves_data(file_path: str) -> pl.DataFrame:
    """Charge les données JEEVES Product depuis la feuille 3-STIBO-TRACKER
    En-têtes ligne 1, données ligne 2+, colonne A = SUPC
    """
    wb = load_workbook(file_path, data_only=True)
    # Utiliser la feuille 3-STIBO-TRACKER
    ws = wb["3-STIBO-TRACKER"]
    
    # Lire les en-têtes de la ligne 1
    headers = [cell.value for cell in ws[1] if cell.value is not None]
    
    # Lire les données à partir de la ligne 2
    data = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if any(cell is not None for cell in row):
            data.append(row[:len(headers)])
    
    # Créer le DataFrame sans schéma strict pour laisser Polars inférer les types
    if data:
        return pl.DataFrame(data, schema=headers, orient="row", infer_schema_length=None)
    else:
        return pl.DataFrame(schema=headers)

def load_ct_data(file_path: str) -> pl.DataFrame:
    """Charge les données CT Ekofisk
    En-têtes ligne 6, données commencent B7 (premier SUPC)
    """
    from pyxlsb import open_workbook
    
    # Détecter le format du fichier
    if file_path.endswith('.xlsb'):
        # Lire le fichier .xlsb
        with open_workbook(file_path) as wb:
            # Trouver la feuille "Item" ou utiliser la première feuille
            sheet_name = None
            for name in wb.sheets:
                if name.lower() == 'item':
                    sheet_name = name
                    break
            if sheet_name is None:
                sheet_name = wb.sheets[0]
            
            with wb.get_sheet(sheet_name) as sheet:
                rows = list(sheet.rows())
                
                # Lire les en-têtes de la ligne 6 (index 5), en commençant à la colonne B (index 1)
                headers = []
                header_counts = {}  # Pour gérer les doublons
                if len(rows) > 5:
                    header_row = rows[5]  # Ligne 6 (0-indexed = 5)
                    for idx, cell in enumerate(header_row):
                        if idx >= 1:  # Colonne B et suivantes (SUPC est en colonne 2)
                            val = cell.v if cell.v is not None else None
                            if val is None:
                                header_name = f"Col_{idx+1}"
                            else:
                                # Convertir en string et nettoyer
                                header_name = str(val).strip() if isinstance(val, str) else str(val)
                            
                            # Gérer les doublons
                            if header_name in header_counts:
                                header_counts[header_name] += 1
                                header_name = f"{header_name}_{header_counts[header_name]}"
                            else:
                                header_counts[header_name] = 0
                            
                            headers.append(header_name)
                
                # Lire les données à partir de la ligne 7 (index 6), colonne B (index 1)
                data = []
                for row_idx in range(6, len(rows)):
                    row = rows[row_idx]
                    # Vérifier si la colonne B (index 1, SUPC) a une valeur numérique
                    if len(row) > 1:
                        supc_val = row[1].v
                        # Accepter seulement les lignes avec un SUPC valide (numérique)
                        if supc_val is not None and (isinstance(supc_val, (int, float)) or str(supc_val).strip()):
                            row_data = []
                            # Commencer à la colonne B (index 1)
                            for col_idx in range(1, min(len(row), len(headers) + 1)):
                                cell_val = row[col_idx].v if col_idx < len(row) else None
                                row_data.append(cell_val)
                            # Compléter avec None si nécessaire
                            while len(row_data) < len(headers):
                                row_data.append(None)
                            data.append(row_data[:len(headers)])
                
                # Créer le DataFrame sans schéma strict pour laisser Polars inférer les types
                if data:
                    return pl.DataFrame(data, schema=headers, orient="row", infer_schema_length=None)
                else:
                    return pl.DataFrame(schema=headers)
    else:
        # Utiliser openpyxl pour les fichiers .xlsx
        wb = load_workbook(file_path, data_only=True)
        # Trouver la feuille Product ou utiliser la feuille active
        ws = None
        for sheet_name in wb.sheetnames:
            if 'product' in sheet_name.lower():
                ws = wb[sheet_name]
                break
        if ws is None:
            ws = wb.active
        
        # Lire les en-têtes de la ligne 6, en commençant à la colonne B
        headers = []
        for cell in ws[6]:
            if cell.column >= 2:  # Colonne B et suivantes
                if cell.value is not None:
                    headers.append(cell.value)
                else:
                    headers.append(f"Col_{cell.column}")
        
        # Lire les données à partir de la ligne 7, colonne B
        data = []
        for row in ws.iter_rows(min_row=7, min_col=2, values_only=True):
            # Vérifier si la première colonne (B) a une valeur (SUPC)
            if row and row[0] is not None:
                data.append(row[:len(headers)])
        
        return pl.DataFrame(data, schema=headers, orient="row")

def load_stibo_data(file_path: str) -> pl.DataFrame:
    """Charge les données STIBO Product
    En-têtes ligne 1, données ligne 2+
    """
    wb = load_workbook(file_path, data_only=True)
    ws = wb.active
    
    # Lire les en-têtes de la ligne 1
    headers = [cell.value for cell in ws[1] if cell.value is not None]
    
    # Lire les données à partir de la ligne 2
    data = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if any(cell is not None for cell in row):
            data.append(row[:len(headers)])
    
    return pl.DataFrame(data, schema=headers, orient="row")


def clean_product_code(value):
    """Nettoie un code produit pour enlever les .0 et normaliser le format"""
    if value is None:
        return None
    # Convertir en string d'abord
    str_val = str(value)
    # Si c'est un nombre avec .0 à la fin, le convertir en entier puis string
    try:
        float_val = float(str_val)
        if float_val.is_integer():
            return str(int(float_val))
        return str_val
    except (ValueError, TypeError):
        return str_val.strip()

def create_range_reconciliation(jeves_df: pl.DataFrame, ct_df: pl.DataFrame, stibo_df: pl.DataFrame) -> pl.DataFrame:
    """Range Reconciliation: Liste tous les produits avec colonnes CT/JEEVES/STIBO et croix"""
    # Identifier la colonne de code produit dans chaque source
    # JEEVES: SUPC (depuis feuille 3-STIBO-TRACKER)
    # CT: SUPC
    # STIBO: SUPC
    
    # Trouver la colonne SUPC dans CT
    ct_product_col = None
    if "SUPC" in ct_df.columns:
        ct_product_col = "SUPC"
    else:
        ct_product_col = ct_df.columns[0]
    
    # Trouver la colonne SUPC dans STIBO
    stibo_product_col = "SUPC" if "SUPC" in stibo_df.columns else stibo_df.columns[0]
    
    # Créer les listes de produits uniques de chaque source
    # Convertir en string et nettoyer le format (enlever .0)
    # JEEVES utilise maintenant SUPC depuis la feuille 3-STIBO-TRACKER
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
    
    # Nettoyer et convertir en string propre
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
    
    # Combiner tous les produits uniques
    all_products = pl.concat([jeves_clean, ct_clean, stibo_clean]).unique("ProductCode")
    
    # Créer les listes pour la vérification
    jeves_list = jeves_clean.to_series().to_list()
    ct_list = ct_clean.to_series().to_list()
    stibo_list = stibo_clean.to_series().to_list()
    
    # Créer le DataFrame avec les colonnes de présence
    reconciliation = all_products.with_columns([
        pl.col("ProductCode").is_in(ct_list).alias("CT_present"),
        pl.col("ProductCode").is_in(jeves_list).alias("JEEVES_present"),
        pl.col("ProductCode").is_in(stibo_list).alias("STIBO_present")
    ]).with_columns([
        # Remplacer True par "X", False par chaîne vide
        pl.when(pl.col("CT_present")).then(pl.lit("X")).otherwise(pl.lit("")).alias("CT"),
        pl.when(pl.col("JEEVES_present")).then(pl.lit("X")).otherwise(pl.lit("")).alias("JEEVES"),
        pl.when(pl.col("STIBO_present")).then(pl.lit("X")).otherwise(pl.lit("")).alias("STIBO")
    ]).with_columns([
        # Colonne résumant les sources absentes (en anglais)
        pl.concat_str([
            pl.when(pl.col("CT_present") == False).then(pl.lit("CT")).otherwise(pl.lit("")),
            pl.when(pl.col("JEEVES_present") == False).then(pl.lit("JEEVES")).otherwise(pl.lit("")),
            pl.when(pl.col("STIBO_present") == False).then(pl.lit("STIBO")).otherwise(pl.lit(""))
        ], separator=", ").str.strip_chars_start(", ").str.strip_chars_end(", ").alias("Absent_from")
    ]).with_columns([
        # Remplacer les chaînes vides par "-" pour plus de clarté
        pl.when(pl.col("Absent_from") == "").then(pl.lit("-")).otherwise(pl.col("Absent_from")).alias("Absent_from")
    ]).select([
        "ProductCode", "CT", "JEEVES", "STIBO", "Absent_from"
    ]).sort("ProductCode")
    
    return reconciliation

    """Analyse 2: Alignement et comparaison des attributs selon le mapping"""
    # Le mapping a les colonnes: STIBO, Jeeves, CT
    # STIBO = nom standardisé de l'attribut
    # Jeeves = nom de la colonne dans JEEVES
    # CT = nom de la colonne dans CT
    
    results = []
    
    for row in mapping_df.iter_rows(named=True):
        stibo_attr = row.get("STIBO")
        jeves_col = row.get("Jeeves")
        ct_col = row.get("CT")
        
        if not stibo_attr or (not jeves_col and not ct_col):
            continue
        
        # Identifier les colonnes produit
        jeves_product_col = "ArtNr"
        ct_product_col = "SUPC" if "SUPC" in ct_df.columns else ct_df.columns[0]
        
        # Créer un DataFrame avec les produits et leurs valeurs pour cet attribut
        comparison_rows = []
        
        # Produits JEEVES
        if jeves_col and jeves_col in jeves_df.columns:
            jeves_attr = jeves_df.select([
                pl.col(jeves_product_col).cast(pl.Utf8).alias("ProductCode"),
                pl.col(jeves_col).cast(pl.Utf8).alias("Valeur")
            ]).with_columns([
                pl.lit("JEEVES").alias("Source"),
                pl.lit(stibo_attr).alias("Attribut")
            ])
            comparison_rows.append(jeves_attr)
        
        # Produits CT
        if ct_col and ct_col in ct_df.columns:
            # Éviter de sélectionner deux fois la même colonne
            if ct_col == ct_product_col:
                ct_attr = ct_df.select([
                    pl.col(ct_product_col).cast(pl.Utf8).alias("ProductCode"),
                    pl.col(ct_col).cast(pl.Utf8).alias("Valeur")
                ]).with_columns([
                    pl.lit("CT").alias("Source"),
                    pl.lit(stibo_attr).alias("Attribut")
                ])
            else:
                ct_attr = ct_df.select([
                    pl.col(ct_product_col).cast(pl.Utf8).alias("ProductCode"),
                    pl.col(ct_col).cast(pl.Utf8).alias("Valeur")
                ]).with_columns([
                    pl.lit("CT").alias("Source"),
                    pl.lit(stibo_attr).alias("Attribut")
                ])
            comparison_rows.append(ct_attr)
        
        if comparison_rows:
            attr_comparison = pl.concat(comparison_rows)
            results.append(attr_comparison)
    
    if results:
        # Combiner tous les résultats
        all_comparisons = pl.concat(results)
        
        # Créer une vue pivotée pour comparer JEEVES vs CT
        jeves_pivot = all_comparisons.filter(pl.col("Source") == "JEEVES").select([
            "ProductCode", "Attribut", "Valeur"
        ]).rename({"Valeur": "Valeur_JEEVES"})
        
        ct_pivot = all_comparisons.filter(pl.col("Source") == "CT").select([
            "ProductCode", "Attribut", "Valeur"
        ]).rename({"Valeur": "Valeur_CT"})
        
        # Joindre pour comparer
        comparison = jeves_pivot.join(
            ct_pivot,
            on=["ProductCode", "Attribut"],
            how="full"
        ).with_columns([
            (pl.col("Valeur_JEEVES") == pl.col("Valeur_CT")).alias("Match"),
            pl.when(pl.col("Valeur_JEEVES").is_null() & pl.col("Valeur_CT").is_not_null())
            .then(pl.lit("Manquant dans JEEVES"))
            .when(pl.col("Valeur_JEEVES").is_not_null() & pl.col("Valeur_CT").is_null())
            .then(pl.lit("Manquant dans CT"))
            .when(pl.col("Valeur_JEEVES") != pl.col("Valeur_CT"))
            .then(pl.lit("Différent"))
            .when(pl.col("Valeur_JEEVES") == pl.col("Valeur_CT"))
            .then(pl.lit("Identique"))
            .otherwise(pl.lit("Non défini"))
            .alias("Statut")
        ])
        
        return comparison
    
    return pl.DataFrame()

def get_file_hash(file_path: str) -> str:
    """Calcule le hash MD5 d'un fichier pour détecter les changements"""
    hash_md5 = hashlib.md5()
    try:
        with open(file_path, "rb") as f:
            for chunk in iter(lambda: f.read(4096), b""):
                hash_md5.update(chunk)
        return hash_md5.hexdigest()
    except FileNotFoundError:
        return None

def get_input_files_hash() -> str:
    """Calcule un hash combiné de tous les fichiers d'input"""
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
            # Si un fichier n'existe pas, retourner None pour forcer la création d'un nouveau fichier
            return None
    
    # Créer un hash combiné
    combined = "|".join(hashes)
    return hashlib.md5(combined.encode()).hexdigest()

def find_existing_output_files() -> dict:
    """Trouve les fichiers de sortie existants"""
    files = {}
    
    # Chercher Range Reconciliation
    range_files = list(Path(".").glob("Range_Reconciliation_*.xlsx"))
    if range_files:
        files["range"] = max(range_files, key=lambda x: x.stat().st_mtime)
    
    return files

def save_hash_info(input_hash: str, output_file: Path):
    """Sauvegarde le hash des inputs avec le nom du fichier de sortie"""
    hash_file = Path(".reconciliation_hash.json")
    hash_info = {
        "input_hash": input_hash,
        "output_file": str(output_file)
    }
    with open(hash_file, "w") as f:
        json.dump(hash_info, f, indent=2)

def load_hash_info() -> dict:
    """Charge le hash précédent"""
    hash_file = Path(".reconciliation_hash.json")
    if hash_file.exists():
        try:
            with open(hash_file, "r") as f:
                return json.load(f)
        except:
            return None
    return None

def main():
    print("Chargement des données...")
    
    # Vérifier si les fichiers d'input ont changé
    current_input_hash = get_input_files_hash()
    previous_hash_info = load_hash_info()
    
    # Déterminer le nom du fichier de sortie
    if current_input_hash and previous_hash_info and previous_hash_info.get("input_hash") == current_input_hash:
        # Les fichiers d'input sont identiques, écraser le fichier existant
        existing_files = find_existing_output_files()
        output_file_range = existing_files.get("range", Path("Range_Reconciliation.xlsx"))
        print("[INFO] Fichiers d'input identiques - ecrasement du fichier existant")
    else:
        # Les fichiers ont changé ou première exécution, créer un nouveau fichier avec timestamp
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_file_range = Path(f"Range_Reconciliation_{timestamp}.xlsx")
        if current_input_hash:
            print("[INFO] Fichiers d'input modifies - creation d'un nouveau fichier")
        else:
            print("[WARN] Certains fichiers d'input manquants - creation d'un nouveau fichier")
    
    # Charger les données
    jeves_df = load_jeves_data("JEEVES/RECONC Product Data 2026-02-04.xlsx")
    print(f"JEEVES: {len(jeves_df)} produits chargés")
    
    ct_df = load_ct_data("CT/P1 Data Cleansing - Product Ekofisk.xlsb")
    print(f"CT: {len(ct_df)} produits chargés")
    
    stibo_df = load_stibo_data("STIBO/extract_stibo_all_products.xlsx")
    print(f"STIBO: {len(stibo_df)} produits chargés")
    
    # Range Reconciliation: Liste tous les produits avec CT/JEEVES/STIBO
    print("\nCréation de la Range Reconciliation...")
    range_reconciliation = create_range_reconciliation(jeves_df, ct_df, stibo_df)
    print(f"Total produits uniques: {len(range_reconciliation)}")
    
    # Générer le fichier Excel
    print(f"\nGénération du fichier Excel...")
    
    # Fichier Range Reconciliation
    range_reconciliation.write_excel(output_file_range)
    print(f"[OK] {output_file_range}")
    
    # Sauvegarder le hash pour la prochaine exécution
    if current_input_hash:
        save_hash_info(current_input_hash, output_file_range)
    
    print("\nTerminé!")

if __name__ == "__main__":
    main()
