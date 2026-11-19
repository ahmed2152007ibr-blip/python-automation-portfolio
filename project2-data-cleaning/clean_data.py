"""
Cleans and merges messy customer Excel files into one clean file.

Input : all .xlsx files inside the 'raw_data' folder
Output: clean_customers.xlsx + cleaning_report.txt
"""

import re
from datetime import datetime
from pathlib import Path

import pandas as pd

INPUT_DIR = Path("raw_data")
OUTPUT_FILE = "clean_customers.xlsx"
REPORT_FILE = "cleaning_report.txt"

# Every column name we might meet -> one standard name
COLUMN_MAP = {
    "name": "name", "full name": "name", "customer_name": "name",
    "phone": "phone", "mobile": "phone", "phone_number": "phone",
    "email": "email", "e-mail": "email", "email_address": "email",
    "signup date": "signup_date", "date": "signup_date", "signup_date": "signup_date",
    "city": "city", "governorate": "city",
    "amount": "amount", "total paid": "amount", "amount_egp": "amount",
}
STANDARD_COLUMNS = ["name", "phone", "email", "signup_date", "city", "amount"]

DATE_FORMATS = ["%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y", "%B %d, %Y", "%d %b %Y", "%d/%m/%y"]


def clean_text(value):
    """Strip and collapse repeated spaces. Empty -> None."""
    if pd.isna(value):
        return None
    text = re.sub(r"\s+", " ", str(value)).strip()
    return text or None


def clean_name(value):
    text = clean_text(value)
    return text.title() if text else None


def clean_city(value):
    text = clean_text(value)
    return text.title() if text else None


def clean_email(value):
    text = clean_text(value)
    return text.lower() if text else None


def clean_phone(value):
    """Return +20XXXXXXXXXX for a valid Egyptian mobile, else None."""
    text = clean_text(value)
    if not text:
        return None
    digits = re.sub(r"\D", "", text)
    if digits.startswith("20") and len(digits) == 12:
        digits = "0" + digits[2:]
    if re.fullmatch(r"01[0125]\d{8}", digits):
        return "+2" + digits
    return None  # invalid number


def clean_date(value):
    text = clean_text(value)
    if not text:
        return None
    for fmt in DATE_FORMATS:
        try:
            return datetime.strptime(text, fmt).strftime("%Y-%m-%d")
        except ValueError:
            continue
    return None  # unknown format


def clean_amount(value):
    """'1,200 EGP' / 'EGP 1200' / 1200 -> 1200.0"""
    text = clean_text(value)
    if not text:
        return None
    number = re.sub(r"[^\d.]", "", text)
    try:
        return float(number)
    except ValueError:
        return None


def load_files():
    frames = []
    for path in sorted(INPUT_DIR.glob("*.xlsx")):
        df = pd.read_excel(path, dtype=object)
        df.columns = [COLUMN_MAP.get(str(c).strip().lower(), str(c).strip().lower())
                      for c in df.columns]
        df = df.reindex(columns=STANDARD_COLUMNS)
        df["source_file"] = path.name
        frames.append(df)
        print(f"Loaded {path.name}: {len(df)} rows")
    return pd.concat(frames, ignore_index=True) if frames else pd.DataFrame()


def main():
    raw = load_files()
    if raw.empty:
        print("No Excel files found in 'raw_data'. Run generate_messy_data.py first.")
        return

    total_in = len(raw)
    df = raw.copy()

    # Count how many values actually change (for the report)
    df["phone_raw"] = df["phone"].map(clean_text)
    df["date_raw"] = df["signup_date"].map(clean_text)

    df["name"] = df["name"].map(clean_name)
    df["phone"] = df["phone"].map(clean_phone)
    df["email"] = df["email"].map(clean_email)
    df["signup_date"] = df["signup_date"].map(clean_date)
    df["city"] = df["city"].map(clean_city)
    df["amount"] = df["amount"].map(clean_amount)

    phones_fixed = int(((df["phone"].notna()) & (df["phone"] != df["phone_raw"])).sum())
    dates_fixed = int(((df["signup_date"].notna()) & (df["signup_date"] != df["date_raw"])).sum())
    invalid_phones = int((df["phone"].isna() & df["phone_raw"].notna()).sum())

    # 1) Remove rows that have no name AND no phone (useless)
    useless = df["name"].isna() & df["phone"].isna()
    removed_empty = int(useless.sum())
    df = df[~useless]

    # 2) Remove duplicates: same phone = same customer. No phone -> use email.
    df["dedupe_key"] = df["phone"].fillna(df["email"])
    before = len(df)
    has_key = df["dedupe_key"].notna()
    df = pd.concat([df[has_key].drop_duplicates(subset="dedupe_key", keep="first"),
                    df[~has_key]]).sort_index()
    removed_dupes = before - len(df)

    final = df[STANDARD_COLUMNS].reset_index(drop=True)
    final.to_excel(OUTPUT_FILE, index=False)

    missing_cells = int(final.isna().sum().sum())

    report = f"""CLEANING REPORT
====================================
Files merged                : {raw['source_file'].nunique()}
Rows before cleaning        : {total_in}
Empty rows removed          : {removed_empty}
Duplicate rows removed      : {removed_dupes}
Rows after cleaning         : {len(final)}

Phone numbers standardized  : {phones_fixed}
Invalid phone numbers       : {invalid_phones} (left empty)
Dates standardized          : {dates_fixed}
Cells still empty           : {missing_cells}

Output file                 : {OUTPUT_FILE}
"""
    Path(REPORT_FILE).write_text(report, encoding="utf-8")
    print("\n" + report)


if __name__ == "__main__":
    main()