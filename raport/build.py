"""Build a formatted Excel sales report from raw order exports (CSV/XLSX)."""
from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, PieChart, Reference
from openpyxl.formatting.rule import CellIsRule, ColorScaleRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo

# Column names accepted in input files -> internal names.
# Lets the script read exports from different shops / Allegro without manual renaming.
ALIASES = {
    "order_id": ["order_id", "nr zamówienia", "numer zamówienia", "id", "nr"],
    "date": ["date", "data", "data zamówienia", "order date"],
    "product": ["product", "produkt", "nazwa produktu", "item"],
    "category": ["category", "kategoria"],
    "channel": ["channel", "kanał", "kanal", "źródło"],
    "qty": ["qty", "ilość", "ilosc", "quantity", "szt"],
    "price": ["price", "cena", "cena brutto", "unit price"],
}
REQUIRED = ["order_id", "date", "product", "qty", "price"]

INK = "1F2937"
ACCENT = "0F7C80"
HEADER_FILL = PatternFill("solid", fgColor=INK)
HEADER_FONT = Font(bold=True, color="FFFFFF")
TITLE_FONT = Font(bold=True, size=16, color=INK)
MONEY = '#,##0.00 "zł"'
THIN = Side(style="thin", color="D1D5DB")


@dataclass
class Report:
    orders: pd.DataFrame
    problems: list[str]


def _normalise_columns(df: pd.DataFrame) -> pd.DataFrame:
    lookup = {a: key for key, names in ALIASES.items() for a in names}
    renamed = {c: lookup[c.strip().lower()] for c in df.columns if c.strip().lower() in lookup}
    df = df.rename(columns=renamed)
    missing = [c for c in REQUIRED if c not in df.columns]
    if missing:
        raise ValueError(f"Brak wymaganych kolumn: {', '.join(missing)}")
    for opt, default in (("category", "Inne"), ("channel", "Nieznany")):
        if opt not in df.columns:
            df[opt] = default
    return df[list(ALIASES)]


def _to_number(s: pd.Series) -> pd.Series:
    # "1 234,50 zł" -> 1234.50
    cleaned = (s.astype(str).str.replace("zł", "", regex=False).str.replace(" ", "", regex=False)
               .str.replace(" ", "", regex=False).str.replace(",", ".", regex=False))
    return pd.to_numeric(cleaned, errors="coerce")


def load(paths: list[Path]) -> Report:
    frames, problems = [], []
    for p in paths:
        if p.suffix.lower() in (".xlsx", ".xls"):
            raw = pd.read_excel(p, dtype=str)
        else:
            raw = pd.read_csv(p, dtype=str, sep=None, engine="python", encoding="utf-8-sig")
        frames.append(_normalise_columns(raw))
    df = pd.concat(frames, ignore_index=True)

    iso = pd.to_datetime(df["date"], errors="coerce", format="%Y-%m-%d")
    df["date"] = iso.fillna(pd.to_datetime(df["date"].where(iso.isna()), errors="coerce", dayfirst=True, format="mixed"))
    df["qty"] = _to_number(df["qty"])
    df["price"] = _to_number(df["price"])
    for col in ("product", "category", "channel"):
        df[col] = df[col].fillna("").str.strip()

    bad = df["date"].isna() | df["qty"].isna() | df["price"].isna() | (df["qty"] <= 0)
    if bad.any():
        problems.append(f"Pominięto {_rows(int(bad.sum()))} z błędną datą, ilością lub ceną.")
    df = df[~bad]

    dupes = df.duplicated(subset=["order_id", "product"], keep="first")
    if dupes.any():
        problems.append(f"Usunięto {_items(int(dupes.sum()))} (ten sam numer zamówienia i produkt).")
    df = df[~dupes].copy()

    df["revenue"] = (df["qty"] * df["price"]).round(2)
    df["month"] = df["date"].dt.to_period("M").astype(str)
    return Report(df.sort_values("date").reset_index(drop=True), problems)


def _rows(n: int) -> str:
    if n == 1:
        return "1 wiersz"
    return f"{n} wiersze" if 2 <= n % 10 <= 4 and not 12 <= n % 100 <= 14 else f"{n} wierszy"


def _items(n: int) -> str:
    if n == 1:
        return "1 zduplikowaną pozycję"
    return f"{n} zduplikowane pozycje" if 2 <= n % 10 <= 4 and not 12 <= n % 100 <= 14 else f"{n} zduplikowanych pozycji"


def _header(ws, row: int, values: list[str]) -> None:
    for i, v in enumerate(values, start=1):
        c = ws.cell(row=row, column=i, value=v)
        c.fill, c.font = HEADER_FILL, HEADER_FONT
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)


def _widths(ws, widths: dict[int, int]) -> None:
    for col, w in widths.items():
        ws.column_dimensions[get_column_letter(col)].width = w


def _write_df(ws, df: pd.DataFrame, start_row: int, money_cols=(), int_cols=()) -> int:
    _header(ws, start_row, list(df.columns))
    for r, row in enumerate(df.itertuples(index=False), start=start_row + 1):
        for c, v in enumerate(row, start=1):
            cell = ws.cell(row=r, column=c, value=v)
            cell.border = Border(bottom=THIN)
            name = df.columns[c - 1]
            if name in money_cols:
                cell.number_format = MONEY
            elif name in int_cols:
                cell.number_format = "#,##0"
    return start_row + len(df)


def build(report: Report, out: Path, title: str = "Raport sprzedaży") -> Path:
    df = report.orders
    wb = Workbook()

    # ---------- Summary ----------
    ws = wb.active
    ws.title = "Podsumowanie"
    ws.sheet_view.showGridLines = False
    ws["A1"] = title
    ws["A1"].font = TITLE_FONT
    ws["A2"] = f"Okres: {df['date'].min():%d.%m.%Y} – {df['date'].max():%d.%m.%Y}"
    ws["A2"].font = Font(color="6B7280")

    kpis = [
        ("Przychód", df["revenue"].sum(), MONEY),
        ("Zamówienia", df["order_id"].nunique(), "#,##0"),
        ("Sprzedane sztuki", df["qty"].sum(), "#,##0"),
        ("Średnia wartość zamówienia", df.groupby("order_id")["revenue"].sum().mean(), MONEY),
    ]
    for i, (label, value, fmt) in enumerate(kpis):
        col = 1 + i
        ws.cell(row=4, column=col, value=label).font = Font(color="6B7280", size=10)
        v = ws.cell(row=5, column=col, value=float(value))
        v.font, v.number_format = Font(bold=True, size=18, color=ACCENT), fmt
    _widths(ws, {1: 22, 2: 22, 3: 22, 4: 26})
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth, ws.page_setup.fitToHeight = 1, 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True

    monthly = (df.groupby("month").agg(Przychód=("revenue", "sum"), Zamówienia=("order_id", "nunique"))
               .reset_index().rename(columns={"month": "Miesiąc"}))
    monthly["Zmiana m/m"] = monthly["Przychód"].pct_change().fillna(0)
    end = _write_df(ws, monthly, 8, money_cols=("Przychód",), int_cols=("Zamówienia",))
    for r in range(9, end + 1):
        ws.cell(row=r, column=4).number_format = "+0.0%;-0.0%;0.0%"
    ws.conditional_formatting.add(f"D9:D{end}", CellIsRule(operator="lessThan", formula=["0"], font=Font(color="B91C1C")))
    ws.conditional_formatting.add(f"D9:D{end}", CellIsRule(operator="greaterThan", formula=["0"], font=Font(color="15803D")))

    line = LineChart()
    line.title, line.height, line.width = "Przychód miesięczny", 8, 18
    line.y_axis.title, line.y_axis.numFmt = "zł", "#,##0"
    line.add_data(Reference(ws, min_col=2, min_row=8, max_row=end), titles_from_data=True)
    line.set_categories(Reference(ws, min_col=1, min_row=9, max_row=end))
    line.series[0].graphicalProperties.line.solidFill = ACCENT
    line.series[0].graphicalProperties.line.width = 28000
    line.legend = None
    ws.add_chart(line, "F8")

    if report.problems:
        r = end + 3
        ws.cell(row=r, column=1, value="Uwagi o danych").font = Font(bold=True)
        for i, p in enumerate(report.problems, start=1):
            ws.cell(row=r + i, column=1, value=f"• {p}").font = Font(color="92400E")

    # ---------- Products ----------
    wp = wb.create_sheet("Produkty")
    prods = (df.groupby(["product", "category"]).agg(Sztuki=("qty", "sum"), Przychód=("revenue", "sum"))
             .reset_index().sort_values("Przychód", ascending=False)
             .rename(columns={"product": "Produkt", "category": "Kategoria"}))
    prods["Udział"] = prods["Przychód"] / prods["Przychód"].sum()
    end_p = _write_df(wp, prods, 1, money_cols=("Przychód",), int_cols=("Sztuki",))
    for r in range(2, end_p + 1):
        wp.cell(row=r, column=5).number_format = "0.0%"
    wp.conditional_formatting.add(f"D2:D{end_p}", ColorScaleRule(start_type="min", start_color="FFFFFF", end_type="max", end_color="99D5D6"))
    _widths(wp, {1: 32, 2: 16, 3: 10, 4: 16, 5: 10})
    wp.freeze_panes = "A2"
    top = BarChart()
    top.type, top.title, top.height, top.width = "bar", "Top 10 produktów", 9, 16
    top.add_data(Reference(wp, min_col=4, min_row=1, max_row=min(11, end_p)), titles_from_data=True)
    top.set_categories(Reference(wp, min_col=1, min_row=2, max_row=min(11, end_p)))
    top.y_axis.numFmt, top.legend = "#,##0", None
    top.x_axis.scaling.orientation = "maxMin"
    top.series[0].graphicalProperties.solidFill = ACCENT
    wp.add_chart(top, "G2")

    # ---------- Channels ----------
    wc = wb.create_sheet("Kanały")
    ch = (df.groupby("channel").agg(Zamówienia=("order_id", "nunique"), Przychód=("revenue", "sum"))
          .reset_index().sort_values("Przychód", ascending=False).rename(columns={"channel": "Kanał"}))
    end_c = _write_df(wc, ch, 1, money_cols=("Przychód",), int_cols=("Zamówienia",))
    _widths(wc, {1: 18, 2: 14, 3: 16})
    pie = PieChart()
    pie.title, pie.height, pie.width = "Udział kanałów w przychodzie", 8, 12
    pie.add_data(Reference(wc, min_col=3, min_row=1, max_row=end_c), titles_from_data=True)
    pie.set_categories(Reference(wc, min_col=1, min_row=2, max_row=end_c))
    wc.add_chart(pie, "E2")

    # ---------- Clean data as an Excel table ----------
    wd = wb.create_sheet("Dane")
    data = df[["order_id", "date", "product", "category", "channel", "qty", "price", "revenue"]].copy()
    data.columns = ["Nr zamówienia", "Data", "Produkt", "Kategoria", "Kanał", "Ilość", "Cena", "Wartość"]
    data["Data"] = data["Data"].dt.date
    end_d = _write_df(wd, data, 1, money_cols=("Cena", "Wartość"), int_cols=("Ilość",))
    for r in range(2, end_d + 1):
        wd.cell(row=r, column=2).number_format = "DD.MM.YYYY"
    t = Table(displayName="Zamowienia", ref=f"A1:H{end_d}")
    t.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
    wd.add_table(t)
    wd.freeze_panes = "A2"
    _widths(wd, {1: 15, 2: 12, 3: 32, 4: 14, 5: 14, 6: 8, 7: 12, 8: 13})

    out.parent.mkdir(parents=True, exist_ok=True)
    wb.save(out)
    return out
