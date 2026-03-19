"""
Launch reconciliation for all three domains (Invoice, Ordering-Shipping, Product).
Output: Reconciliation_{market}.xlsx with 5 tabs: Product, Vendor Invoice, Vendor OS, Customer Invoice, Customer OS.

Sources are read from dated folders:
  STIBO/{date}/, CT/{date}/, ERP/{market}/{date}/  (e.g. 2302 = 23 Feb, 1003 = 10 Mar)

Markets and ERP associations are defined in markets.json.

Usage:
  python run_reconciliation.py --market Ekofisk
  python run_reconciliation.py --market Ekofisk --date 2302
  python run_reconciliation.py --market Fresh_Direct --date 1003
  python run_reconciliation.py --market all --date 1003
  python run_reconciliation.py --market all --domains invoice_ordering --date 2302
"""
import argparse
from pathlib import Path

import polars as pl

import market_config

DEFAULT_DATE = "2302"


def main() -> None:
    all_markets = market_config.list_markets()

    parser = argparse.ArgumentParser(
        description=(
            "Run reconciliation. Output: Reconciliation_{market}.xlsx with 5 tabs "
            "(Product, Vendor Invoice, Vendor OS, Customer Invoice, Customer OS)."
        )
    )
    parser.add_argument(
        "--market",
        default="Ekofisk",
        help=(
            f"Market name or 'all'. Known markets: {', '.join(all_markets)}. "
            "Market names are case-sensitive."
        ),
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
        help=(
            f"Date folder for sources: STIBO/{{date}}/, CT/{{date}}/, ERP/{{market}}/{{date}}/ "
            f"(e.g. 2302, 1003). Default: {DEFAULT_DATE}"
        ),
    )
    args = parser.parse_args()

    date_folder = args.date.strip()
    out_dir = Path("output") / date_folder
    out_dir.mkdir(parents=True, exist_ok=True)

    # Resolve market(s)
    if args.market.lower() == "all":
        markets_to_run = all_markets
    else:
        # Case-insensitive match
        match = next((m for m in all_markets if m.lower() == args.market.lower()), None)
        if match is None:
            parser.error(
                f"Unknown market '{args.market}'. Known markets: {', '.join(all_markets)}"
            )
        markets_to_run = [match]

    generated_files: list[Path] = []
    skipped: list[str] = []

    print("=" * 60)
    print("  RECONCILIATION")
    print("  Output: Reconciliation_{market}.xlsx (5 tabs)")
    print("  Tabs: Product | Vendor Invoice | Vendor OS | Customer Invoice | Customer OS")
    print("=" * 60)
    print(f"  Market(s): {', '.join(markets_to_run)}")
    print(f"  Domains:   {args.domains}")
    print(f"  Date:      {date_folder}")
    print(f"  Output:    {out_dir}/")
    print()

    from reconcile_ekofisk_invoice_ordering import (
        run_invoice_ordering_reconciliation,
        write_reconciliation_excel_5_tabs,
        KEY_COL,
    )
    import reconcile_products

    for market in markets_to_run:
        erp = market_config.get_erp_name(market)
        print(f"{'=' * 60}")
        print(f"  MARKET: {market}  (ERP: {erp})")
        print(f"{'=' * 60}")

        product_df: pl.DataFrame | None = None

        if args.domains in ("product", "all"):
            print("-" * 60)
            print("  [1/2] DOMAIN: PRODUCT (Range Reconciliation)")
            print("-" * 60)
            try:
                product_df = reconcile_products.main(
                    date_folder=date_folder,
                    output_dir=out_dir,
                    write_range_file=False,
                    market=market,
                )
                print(f"  -> Product data: {product_df.height} rows")
            except NotImplementedError as e:
                skipped.append(f"{market} / Product: {e}")
                print(f"  SKIP (ERP non supporté): {e}")
                product_df = None
            except FileNotFoundError as e:
                skipped.append(f"{market} / Product: {e}")
                print(f"  SKIP (fichier manquant): {e}")
                product_df = None
            except Exception as e:
                skipped.append(f"{market} / Product: {e}")
                print(f"  SKIP (erreur): {e}")
                product_df = None
            print()

        if args.domains == "product" and product_df is not None:
            # Product-only run: write Reconciliation file with empty Invoice/OS tabs
            erp = market_config.get_erp_name(market)
            empty = pl.DataFrame({KEY_COL: pl.Series([], dtype=pl.Utf8)})
            out_path = out_dir / f"Reconciliation_{market}.xlsx"
            write_reconciliation_excel_5_tabs(
                out_path, empty, empty, product_df=product_df, erp_name=erp
            )
            print(f"  -> écrit: {out_path}")
            generated_files.append(out_path)

        if args.domains in ("invoice_ordering", "all"):
            print("-" * 60)
            print("  [2/2] DOMAIN: INVOICE + ORDERING-SHIPPING")
            print("-" * 60)
            try:
                out_path = run_invoice_ordering_reconciliation(
                    market, out_dir, product_df=product_df, date_folder=date_folder
                )
                generated_files.append(out_path)
            except NotImplementedError as e:
                skipped.append(f"{market} / Invoice-OS: {e}")
                print(f"  SKIP (ERP non supporté): {e}")
            except FileNotFoundError as e:
                skipped.append(f"{market} / Invoice-OS: {e}")
                print(f"  SKIP (fichier manquant): {e}")
            print()

    print("=" * 60)
    print("  RÉSUMÉ")
    print("=" * 60)
    if generated_files:
        print("  Fichiers générés:")
        for p in generated_files:
            print(f"    - {p}")
    if skipped:
        print("  Ignorés (fichier manquant, ERP non supporté ou erreur):")
        for s in skipped:
            print(f"    - {s}")
    if not generated_files and not skipped:
        print("  Aucune action (vérifiez --domains et --market).")
    print()
    print("  Terminé.")
    print("=" * 60)


if __name__ == "__main__":
    main()
