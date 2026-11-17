"""
Books Scraper
Scrapes all books from books.toscrape.com (50 pages) and saves them to Excel.

Columns: title, price, rating, in_stock, url
"""

import re
import time
from urllib.parse import urljoin

import pandas as pd
import requests
from bs4 import BeautifulSoup

BASE_URL = "https://books.toscrape.com/"
PAGE_URL = BASE_URL + "catalogue/page-{}.html"
TOTAL_PAGES = 50
DELAY_SECONDS = 1
OUTPUT_FILE = "books.xlsx"

# Ratings are stored as words in the HTML class (e.g. "star-rating Three")
RATING_MAP = {"One": 1, "Two": 2, "Three": 3, "Four": 4, "Five": 5}

HEADERS = {"User-Agent": "Mozilla/5.0 (portfolio scraping demo)"}


def fetch_page(page_number):
    """Download one page. Returns HTML text, or None if it failed."""
    url = PAGE_URL.format(page_number)
    try:
        response = requests.get(url, headers=HEADERS, timeout=15)
        response.raise_for_status()
        response.encoding = "utf-8"  # avoids broken currency symbols
        return response.text
    except requests.RequestException as error:
        print(f"  [!] Page {page_number} failed: {error}")
        return None


def parse_books(html, page_number):
    """Extract all books from one page's HTML."""
    soup = BeautifulSoup(html, "html.parser")
    books = []

    for item in soup.select("article.product_pod"):
        try:
            link = item.h3.a
            title = link["title"]
            url = urljoin(PAGE_URL.format(page_number), link["href"])

            # "£51.77" -> 51.77 (as a real number, not text)
            price_text = item.select_one("p.price_color").get_text()
            price = float(re.sub(r"[^\d.]", "", price_text))

            rating_class = item.select_one("p.star-rating")["class"]
            rating = next((RATING_MAP[c] for c in rating_class if c in RATING_MAP), None)

            stock_text = item.select_one("p.instock.availability").get_text(strip=True)
            in_stock = "In stock" in stock_text

            books.append(
                {
                    "title": title,
                    "price": price,
                    "rating": rating,
                    "in_stock": in_stock,
                    "url": url,
                }
            )
        except (AttributeError, KeyError, ValueError) as error:
            print(f"  [!] Skipped one book on page {page_number}: {error}")

    return books


def main():
    all_books = []

    for page in range(1, TOTAL_PAGES + 1):
        print(f"Scraping page {page}/{TOTAL_PAGES}...")
        html = fetch_page(page)
        if html:
            page_books = parse_books(html, page)
            all_books.extend(page_books)
            print(f"  Found {len(page_books)} books (total: {len(all_books)})")
        time.sleep(DELAY_SECONDS)

    if not all_books:
        print("No data collected. Nothing to save.")
        return

    df = pd.DataFrame(all_books)
    df.to_excel(OUTPUT_FILE, index=False)
    print(f"\nDone! Saved {len(df)} books to {OUTPUT_FILE}")


if __name__ == "__main__":
    main()