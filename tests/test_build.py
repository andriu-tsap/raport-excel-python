from pathlib import Path

import pandas as pd
import pytest
from openpyxl import load_workbook

from raport import build, load
from raport.__main__ import main


def write(tmp_path: Path, text: str, name="in.csv") -> Path:
    p = tmp_path / name
    p.write_text(text, encoding="utf-8")
    return p


def test_polish_headers_and_number_formats(tmp_path):
    f = write(tmp_path, "Nr zamówienia;Data;Produkt;Kanał;Ilość;Cena\n"
                        "1;05.03.2026;Sencha;Allegro;2;\"1 054,50 zł\"\n"
                        "2;06.03.2026;Mięta;Sklep www;1;22\n")
    r = load([f])
    assert list(r.orders["revenue"]) == [2109.0, 22.0]
    assert r.orders["category"].eq("Inne").all()
    assert r.orders["date"].iloc[0] == pd.Timestamp(2026, 3, 5)


def test_bad_rows_and_duplicates_are_reported(tmp_path):
    f = write(tmp_path, "order_id,date,product,qty,price\n"
                        "1,2026-01-02,A,1,10\n1,2026-01-02,A,1,10\n"
                        "2,nie-data,B,1,10\n3,2026-01-03,C,0,10\n")
    r = load([f])
    assert len(r.orders) == 1
    assert any("Pominięto 2 wiersze" in p for p in r.problems)
    assert any("1 zduplikowaną pozycję" in p for p in r.problems)


def test_missing_column_raises(tmp_path):
    f = write(tmp_path, "data,produkt,cena\n2026-01-01,A,1\n")
    with pytest.raises(ValueError, match="order_id"):
        load([f])


def test_workbook_structure(tmp_path):
    f = write(tmp_path, "order_id,date,product,category,channel,qty,price\n"
                        "1,2026-01-05,A,X,www,2,10\n2,2026-02-05,B,Y,Allegro,1,30\n")
    out = build(load([f]), tmp_path / "r.xlsx")
    wb = load_workbook(out)
    assert wb.sheetnames == ["Podsumowanie", "Produkty", "Kanały", "Dane"]
    assert wb["Podsumowanie"]["A5"].value == 50
    assert wb["Dane"].tables["Zamowienia"].ref == "A1:H3"


def test_cli(tmp_path, capsys):
    f = write(tmp_path, "order_id,date,product,qty,price\n1,2026-01-05,A,1,10\n")
    assert main([str(f), "-o", str(tmp_path / "x.xlsx")]) == 0
    assert "Gotowe" in capsys.readouterr().out
    assert main([str(tmp_path / "missing.csv")]) == 1
