# Messy Excel Cleaner & Merger (Python)

Takes several messy customer Excel files, cleans them, and merges them into one clean file with a report of what was fixed.

## The problem
Real client data is rarely clean. In the sample files (115 rows across 3 files):
- Names with extra spaces and random capitalization
- Phone numbers in 6 different formats
- Dates in 5 different formats
- Different column names in every file ("Phone", "Mobile", "phone_number"...)
- Duplicate customers inside files and across files
- Empty cells and completely empty rows

## What the script does
- Standardizes column names across files
- Trims spaces and fixes name/city capitalization
- Converts every phone to `+20XXXXXXXXXX` (invalid numbers are flagged, not guessed)
- Converts every date to `YYYY-MM-DD`
- Converts amounts like `1,200 EGP` to numbers
- Removes empty rows and duplicate customers (same phone = same customer)
- Writes `clean_customers.xlsx` and `cleaning_report.txt`

## Result on the sample data
- Rows before: 115 -> after: 90 (3 empty rows and 22 duplicates removed)
- All phones and dates in one consistent format

## How to run
    pip install pandas openpyxl
    python generate_messy_data.py   # creates fake messy files in raw_data/
    python clean_data.py            # cleans and merges them

All customer data is randomly generated fake data.