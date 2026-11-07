"""Generate realistic sample exports (two files, mixed formats) for the demo."""
import csv
import random
from datetime import date, timedelta
from pathlib import Path

random.seed(7)
HERE = Path(__file__).parent
PRODUCTS = [("Sencha Asamushi", "Zielona", 54), ("Matcha ceremonialna", "Zielona", 79), ("Darjeeling First Flush", "Czarna", 62),
            ("Assam Halmari", "Czarna", 36), ("Dong Ding Oolong", "Oolong", 68), ("Mięta z Podkarpacia", "Zioła", 22),
            ("Rumianek i melisa", "Zioła", 24), ("Rooibos z pomarańczą", "Zioła", 28), ("Biała Bai Mu Dan", "Biała", 48)]
weights = [9, 8, 5, 7, 3, 8, 6, 6, 2]
start = date(2026, 1, 1)


def rows(n_days, channel_pick, oid):
    for d in range(n_days):
        day = start + timedelta(days=d)
        for _ in range(random.randint(4, 12) + d // 40):
            oid += 1
            for p in list(dict.fromkeys(random.choices(PRODUCTS, weights, k=random.randint(1, 3)))):
                yield oid, day, p, random.choice([1, 1, 1, 2, 3]), channel_pick()


# Shop export: English-ish headers, ISO dates, dot decimals
with open(HERE / "sklep_www.csv", "w", newline="", encoding="utf-8") as f:
    w = csv.writer(f)
    w.writerow(["order_id", "date", "product", "category", "channel", "qty", "price"])
    for oid, day, (name, cat, price), q, ch in rows(304, lambda: random.choice(["Sklep www", "Sklep www", "Instagram"]), 10000):
        w.writerow([oid, day.isoformat(), name, cat, ch, q, f"{price:.2f}"])

# Marketplace export: Polish headers, dd.mm.yyyy, "54,00 zł", semicolons, plus a few broken rows
with open(HERE / "allegro.csv", "w", newline="", encoding="utf-8") as f:
    w = csv.writer(f, delimiter=";")
    w.writerow(["Nr zamówienia", "Data zamówienia", "Nazwa produktu", "Kategoria", "Kanał", "Ilość", "Cena brutto"])
    for oid, day, (name, cat, price), q, _ in rows(304, lambda: "Allegro", 50000):
        w.writerow([f"A-{oid}", day.strftime("%d.%m.%Y"), name, cat, "Allegro", q, f"{price:.2f}".replace(".", ",") + " zł"])
    w.writerow(["A-99999", "brak", "Sencha Asamushi", "Zielona", "Allegro", "1", "54,00 zł"])
    w.writerow(["A-50001", (start).strftime("%d.%m.%Y"), "Sencha Asamushi", "Zielona", "Allegro", "1", "54,00 zł"])
print("Zapisano przyklad/sklep_www.csv i przyklad/allegro.csv")
