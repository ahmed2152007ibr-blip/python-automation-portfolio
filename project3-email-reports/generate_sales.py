import pandas as pd
import random
from datetime import datetime, timedelta

# إعداد بيانات منتجات واقعية
products = [
    ("Wireless Earbuds", "Electronics", 450),
    ("Smart Watch", "Electronics", 1200),
    ("Phone Case", "Accessories", 120),
    ("Fast Charger 20W", "Accessories", 250),
    ("Laptop Stand", "Office", 380),
    ("Mechanical Keyboard", "Office", 950),
    ("USB-C Hub", "Accessories", 420),
    ("Bluetooth Speaker", "Electronics", 680),
]

random.seed(42)  # لتثبيت الأرقام عند التجربة
rows = []
start_date = datetime(2026, 11, 1)

for i in range(120):
    product, category, price = random.choice(products)
    quantity = random.randint(1, 5)
    date = (start_date + timedelta(days=random.randint(0, 18))).strftime("%Y-%m-%d")
    total = price * quantity
    rows.append({
        "Date": date,
        "Product": product,
        "Category": category,
        "Quantity": quantity,
        "Unit_Price": price,
        "Total_Sales": total
    })

df = pd.DataFrame(rows)
df = df.sort_values(by="Date").reset_index(drop=True)
df.to_excel("sales_data.xlsx", index=False)

print("✅ Created sales_data.xlsx with 120 sales records successfully!")