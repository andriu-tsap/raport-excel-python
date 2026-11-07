"""Command line: python -m raport dane/*.csv -o raport.xlsx"""
import argparse
import sys
from pathlib import Path

from .build import build, load


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(prog="raport", description="Tworzy sformatowany raport sprzedaży w Excelu z eksportów zamówień (CSV/XLSX).")
    ap.add_argument("files", nargs="+", type=Path, help="pliki CSV lub XLSX z zamówieniami")
    ap.add_argument("-o", "--output", type=Path, default=Path("raport_sprzedazy.xlsx"), help="plik wynikowy (domyślnie raport_sprzedazy.xlsx)")
    ap.add_argument("-t", "--title", default="Raport sprzedaży", help="tytuł raportu")
    args = ap.parse_args(argv)

    missing = [str(f) for f in args.files if not f.exists()]
    if missing:
        print(f"Nie znaleziono plików: {', '.join(missing)}", file=sys.stderr)
        return 1
    try:
        report = load(args.files)
    except ValueError as e:
        print(f"Błąd danych: {e}", file=sys.stderr)
        return 2
    out = build(report, args.output, args.title)
    print(f"Gotowe: {out} ({len(report.orders)} pozycji, {report.orders['order_id'].nunique()} zamówień)")
    for p in report.problems:
        print(f"  uwaga: {p}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
