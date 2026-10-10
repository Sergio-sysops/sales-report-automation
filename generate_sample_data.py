"""Generate a deliberately messy sales CSV to demo the report automation.

Everything here is fake data: inconsistent date formats, messy product names,
prices with currency symbols, duplicated rows, empty cells and typos.
"""
import csv
import random
from datetime import date, timedelta

random.seed(42)

PRODUCTS = {
    "Wireless Mouse": 24.90,
    "Mechanical Keyboard": 79.00,
    "USB-C Hub": 39.50,
    "27\" Monitor": 219.00,
    "Laptop Stand": 32.00,
    "Webcam HD": 49.90,
}
REGIONS = ["North", "South", "East", "West"]
SELLERS = ["Ana", "Luis", "Marta", "Javier", "Lucía"]


def messy_date(d):
    fmt = random.choice(["%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y", "%b %d %Y"])
    return d.strftime(fmt)


def messy_product(name):
    return random.choice([name, name.upper(), name.lower(), f"  {name} ", name.replace(" ", "  ")])


def messy_price(p):
    return random.choice([f"{p:.2f}", f"{p:.2f} €", f"€{p:.2f}", f"{p:.2f}".replace(".", ",")])


def main(path="ventas_sucias.csv", n=1200):
    start = date(2026, 1, 1)
    rows = []
    for i in range(n):
        d = start + timedelta(days=random.randint(0, 272))
        product, price = random.choice(list(PRODUCTS.items()))
        rows.append([
            f"ORD-{10000 + i}",
            messy_date(d),
            messy_product(product),
            random.choice(REGIONS + [r.lower() for r in REGIONS]),
            random.choice(SELLERS),
            random.choice([1, 1, 1, 2, 2, 3, 5]),
            messy_price(price),
        ])
    # Real-world dirt: duplicates, blanks and impossible values
    rows += random.sample(rows, 40)
    for r in random.sample(rows, 25):
        r[random.choice([3, 4])] = ""
    for r in random.sample(rows, 6):
        r[5] = -r[5]
    rows += [[""] * 7 for _ in range(10)]
    random.shuffle(rows)

    with open(path, "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(["order_id", "date", "product", "region", "seller", "qty", "unit_price"])
        w.writerows(rows)
    print(f"Wrote {len(rows)} messy rows to {path}")


if __name__ == "__main__":
    main()
