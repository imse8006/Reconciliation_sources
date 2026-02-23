"""
Launch reconciliation for all three domains (Invoice, Ordering-Shipping, Product).
Output: Reconciliation_{market}.xlsx with 5 tabs: Product, Vendor Invoice, Vendor OS, Customer Invoice, Customer OS.

Sources are read from dated folders: STIBO/{date}/, CT/{date}/, JEEVES/{date}/ (e.g. 2302 = 23 Feb, 0203 = 2 Mar).

Usage:
  python run_reconciliation.py --market ekofisk
  python run_reconciliation.py --market ekofisk --date 2302
  python run_reconciliation.py --market all --date 0203
  python run_reconciliation.py --market all --domains invoice_ordering --date 2302
"""
import argparse
from pathlib import Path

import polars as pl

MARKETS = ["Ekofisk", "Fresh", "Classic"]
DEFAULT_DATE = "2302"


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Run reconciliation. Output: Reconciliation_{market}.xlsx with 5 tabs (Product, Vendor Invoice, Vendor OS, Customer Invoice, Customer OS)."
    )
    parser.add_argument(
        "--market",
        choices=["ekofisk", "fresh", "classic", "all"],
        default="ekofisk",
        help="Market to run: one of ekofisk, fresh, classic, or all",
    )
    parser.add_argument(
        "--domains",
        choices=["invoice_ordering", "product", "all"],
        default="all",
        help="Domains to run: invoice_ordering, product (Range), or all",
    )
    parser.add_argument(
        "--date",
        default=DEFAULT_DATE,
        metavar="DDMM",
        help=f"Date folder for sources: STIBO/{{date}}/, CT/{{date}}/, JEEVES/{{date}}/ (e.g. 2302, 0203). Default: {DEFAULT_DATE}",
    )
    args = parser.parse_args()

    date_folder = args.date.strip()
    out_dir = Path("output") / date_folder
    out_dir.mkdir(parents=True, exist_ok=True)

    product_df: pl.DataFrame | None = None
    generated_files: list[Path] = []
    skipped: list[str] = []

    print("=" * 60)
    print("  RECONCILIATION")
    print("  Output: Reconciliation_{market}.xlsx (5 tabs)")
    print("  Tabs: Product | Vendor Invoice | Vendor OS | Customer Invoice | Customer OS")
    print("=" * 60)
    print(f"  Market(s): {args.market}")
    print(f"  Domains:   {args.domains}")
    print(f"  Date:      {date_folder} (sources: STIBO/{date_folder}/, CT/{date_folder}/, JEEVES/{date_folder}/)")
    print(f"  Output:   {out_dir}/")
    print()

    if args.domains in ("product", "all"):
        print("-" * 60)
        print("  [1/2] DOMAIN: PRODUCT (Range Reconciliation)")
        print("-" * 60)
        try:
            import reconcile_products
            product_df = reconcile_products.main(
                date_folder=date_folder, output_dir=out_dir, write_range_file=False
            )
            print(f"  -> Product data: {product_df.height} rows (pour onglet Product dans Reconciliation_*.xlsx)")
        except FileNotFoundError as e:
            skipped.append(f"Product: {e}")
            print(f"  SKIP: {e}")
            product_df = None
        except Exception as e:
            skipped.append(f"Product: {e}")
            print(f"  SKIP: {e}")
            product_df = None
        print()

    # Always run invoice_ordering when we have at least one market: single output = Reconciliation_{market}.xlsx
    if args.domains in ("invoice_ordering", "product", "all"):
        from reconcile_ekofisk_invoice_ordering import run_invoice_ordering_reconciliation

        if args.market == "all":
            markets_to_run = MARKETS
        else:
            markets_to_run = [args.market.capitalize()]

        print("-" * 60)
        print("  [2/2] DOMAIN: INVOICE + ORDERING-SHIPPING")
        print(f"         Markets: {', '.join(markets_to_run)}")
        print("-" * 60)
        for market in markets_to_run:
            try:
                out_path = run_invoice_ordering_reconciliation(
                    market, out_dir, product_df=product_df, date_folder=date_folder
                )
                generated_files.append(out_path)
            except FileNotFoundError as e:
                skipped.append(f"{market}: {e}")
                print(f"  SKIP {market}: {e}")
        print()

    print("=" * 60)
    print("  RÉSUMÉ")
    print("=" * 60)
    if generated_files:
        print("  Fichiers générés:")
        for p in generated_files:
            print(f"    - {p}")
    if skipped:
        print("  Ignorés (fichier source manquant ou erreur):")
        for s in skipped:
            print(f"    - {s}")
    if not generated_files and not skipped:
        print("  Aucune action (vérifiez --domains et --market).")
    print()
    print("  Terminé.")
    print("=" * 60)


if __name__ == "__main__":
    main()
