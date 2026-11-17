# Books Scraper (Python)

Scrapes all 1,000 books from books.toscrape.com (50 pages) into a clean Excel file.

## What it extracts
Title, price (numeric), star rating (1-5), stock availability, and book URL.

## How it works
- Uses requests + BeautifulSoup to read each page
- Converts prices to numbers and ratings to integers
- Waits 1 second between pages to be polite to the server
- Skips and reports any page or book that fails instead of crashing
- Saves everything to books.xlsx with pandas

## How to run
    pip install requests beautifulsoup4 pandas openpyxl
    python scraper.py

## Output
books.xlsx with 1,000 rows and 5 columns.