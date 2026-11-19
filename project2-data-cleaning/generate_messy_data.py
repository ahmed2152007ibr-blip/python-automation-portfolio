"""
Generates 3 messy customer Excel files (fake data) to practice cleaning on.

Problems built in on purpose:
- extra spaces and inconsistent capitalization in names
- phone numbers in many formats
- dates in many formats
- duplicate customers (within and across files)
- empty cells and fully empty rows
- different column names in each file
"""

import random
from datetime import date, timedelta
from pathlib import Path

import pandas as pd

random.seed(42)  # same fake data every run

OUTPUT_DIR = Path("raw_data")

FIRST_NAMES = ["Ahmed", "Mohamed", "Mahmoud", "Omar", "Youssef", "Karim", "Hassan",
               "Sara", "Mona", "Nour", "Salma", "Hana", "Dina", "Laila", "Yasmin"]
LAST_NAMES = ["Hassan", "Ibrahim", "Mostafa", "Ali", "Said", "Fouad", "Gamal",
              "Nabil", "Adel", "Tarek", "Samir", "Khaled"]
CITIES = ["Cairo", "Giza", "Alexandria", "Mansoura", "Tanta"]


def make_phone_digits():
    """Egyptian mobile: 01X + 8 digits, returned as 11 digits."""
    return "01" + random.choice("0125") + "".join(random.choice("0123456789") for _ in range(8))


def messy_phone(digits):
    """Return the same phone in a random messy format."""
    styles = [
        digits,                                   # 01012345678
        "+2" + digits,                            # +201012345678
        "2" + digits,                             # 201012345678
        f"{digits[:3]}-{digits[3:7]}-{digits[7:]}",   # 010-1234-5678
        f"{digits[:4]} {digits[4:7]} {digits[7:]}",   # 0101 234 5678
        f"({digits[:3]}) {digits[3:]}",           # (010) 12345678
    ]
    return random.choice(styles)


def messy_date(d):
    """Return the same date in a random messy format."""
    styles = [
        d.strftime("%Y-%m-%d"),     # 2024-03-15
        d.strftime("%d/%m/%Y"),     # 15/03/2024
        d.strftime("%d-%m-%Y"),     # 15-03-2024
        d.strftime("%B %d, %Y"),    # March 15, 2024
        d.strftime("%d %b %Y"),     # 15 Mar 2024
    ]
    return random.choice(styles)


def messy_name(name):
    """Random capitalization and stray spaces."""
    style = random.choice(["upper", "lower", "normal", "spaces", "spaces"])
    if style == "upper":
        return name.upper()
    if style == "lower":
        return name.lower()
    if style == "spaces":
        return "  " + name.replace(" ", "   ") + " "
    return name


def make_customer(customer_id):
    name = f"{random.choice(FIRST_NAMES)} {random.choice(LAST_NAMES)}"
    start = date(2023, 1, 1)
    return {
        "id": customer_id,
        "name": name,
        "phone": make_phone_digits(),
        "email": f"{name.lower().replace(' ', '.')}{customer_id}@example.com",
        "date": start + timedelta(days=random.randint(0, 700)),
        "city": random.choice(CITIES),
        "amount": random.randint(200, 5000),
    }


def build_file(customers, columns):
    """Turn clean customers into a messy DataFrame with the given column names."""
    rows = []
    for c in customers:
        rows.append({
            columns["name"]: messy_name(c["name"]),
            columns["phone"]: messy_phone(c["phone"]),
            columns["email"]: c["email"].upper() if random.random() < 0.2 else c["email"],
            columns["date"]: messy_date(c["date"]),
            columns["city"]: random.choice([c["city"], c["city"].lower(), " " + c["city"] + " "]),
            columns["amount"]: random.choice([c["amount"], f"{c['amount']:,} EGP", f"EGP {c['amount']}"]),
        })

    df = pd.DataFrame(rows)

    # Blank out random cells
    for _ in range(max(2, len(df) // 8)):
        row = random.randrange(len(df))
        col = random.choice([columns["email"], columns["city"], columns["amount"]])
        df.loc[row, col] = None

    # A row with no name AND no phone (useless row)
    junk = {col: None for col in df.columns}
    junk[columns["city"]] = "Cairo"
    df = pd.concat([df, pd.DataFrame([junk])], ignore_index=True)

    return df


def main():
    OUTPUT_DIR.mkdir(exist_ok=True)
    customers = [make_customer(i) for i in range(1, 91)]  # 90 unique customers

    parts = [customers[0:35], customers[30:65], customers[60:90]]
    # overlap on purpose: same customer appears in two files
    # plus a few exact repeats inside each file
    for part in parts:
        part.extend(random.sample(part, 4))

    column_sets = [
        {"name": "Name", "phone": "Phone", "email": "Email",
         "date": "Signup Date", "city": "City", "amount": "Amount"},
        {"name": "Full Name", "phone": "Mobile", "email": "E-mail",
         "date": "Date", "city": "Governorate", "amount": "Total Paid"},
        {"name": "customer_name", "phone": "phone_number", "email": "email_address",
         "date": "signup_date", "city": "city", "amount": "amount_egp"},
    ]
    file_names = ["customers_cairo.xlsx", "customers_giza.xlsx", "customers_alex.xlsx"]

    for name, part, cols in zip(file_names, parts, column_sets):
        random.shuffle(part)
        df = build_file(part, cols)
        df.to_excel(OUTPUT_DIR / name, index=False)
        print(f"Created {OUTPUT_DIR / name} ({len(df)} rows)")

    print("\nDone. Messy files are ready in the 'raw_data' folder.")


if __name__ == "__main__":
    main()