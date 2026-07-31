from google.oauth2.service_account import Credentials
import gspread
import requests
from bs4 import BeautifulSoup
from datetime import datetime
import time
import os
import json
import re

# ======================================
# CONFIG
# ======================================

SHEET_NAME    = "topps-tracker"
HISTORY_SHEET = "price_history"

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive"
]

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
}

# ======================================
# UPDATE SETTINGS
# ======================================

# True  = skip rows already updated today
# False = refresh everything
SKIP_UPDATED_TODAY = True

# ======================================
# COLUMN INDICES (0-based)
# ======================================

COL_RELEASE_DATE  = 0
COL_YEAR          = 1
COL_CATEGORY      = 2
COL_PRODUCT       = 3
COL_BOX_TYPE      = 4
COL_MSRP          = 5
COL_WAXSTAT_AVG   = 6   # G
COL_NINETY_VALUE  = 7   # H
COL_GAIN          = 8
COL_ROI           = 9
COL_WAXSTAT_URL   = 10  # K
COL_LAST_UPDATED  = 11  # L
COL_EBAY_PROFIT   = 12  # M
COL_TRACKER_URL   = 13  # N
COL_EBAY_PRICE    = 14  # O
COL_EXCLUDE_KW    = 15  # P

# ======================================
# CONNECT TO GOOGLE SHEET
# ======================================

creds_env = os.environ.get("GOOGLE_CREDENTIALS_JSON")

if creds_env:
    creds_dict = json.loads(creds_env)
    creds = Credentials.from_service_account_info(creds_dict, scopes=SCOPES)
    print("Loaded credentials from environment variable")
else:
    CREDS_FILE = "topps-tracker-499814-874632f78fe4.json"
    creds = Credentials.from_service_account_file(CREDS_FILE, scopes=SCOPES)
    print("Loaded credentials from local file")

client = gspread.authorize(creds)
wb     = client.open(SHEET_NAME)
sheet  = wb.sheet1
rows   = sheet.get_all_values()

print(f"Found {len(rows)-1} data rows")

# ======================================
# PRICE HISTORY SETUP
# ======================================

today     = datetime.now()
today_str = today.strftime("%Y-%m-%d")
is_snapshot_day = today.day in (1, 15)

try:
    history_sheet = wb.worksheet(HISTORY_SHEET)
    print(f"Found existing '{HISTORY_SHEET}' sheet")
except gspread.exceptions.WorksheetNotFound:
    history_sheet = wb.add_worksheet(title=HISTORY_SHEET, rows=5000, cols=4)
    history_sheet.append_row(
        ["Date", "Product", "Box Type", "WaxStat Avg"],
        value_input_option="RAW"
    )
    print(f"Created new '{HISTORY_SHEET}' sheet")

if is_snapshot_day:
    existing_history = history_sheet.get_all_values()
    logged_keys = set()
    for h_row in existing_history[1:]:
        if len(h_row) >= 3:
            logged_keys.add((h_row[0], h_row[1], h_row[2]))
    print(f"Snapshot day — {len(logged_keys)} existing history records loaded")
else:
    logged_keys = set()
    print("Not a snapshot day — skipping price history")

# ======================================
# HELPER: SCRAPE SPORTSCARDSPRO PRICE
# ======================================

def get_sportscards_price(tracker_url, exclude_keywords=None):
    """
    Scrapes sold listings from a SportsCardsPro/PriceCharting product page.
    - Filters out listings whose titles contain any exclude_keywords
    - Returns median of last 10 clean sales (min 5)
    - Falls back to their market price if fewer than 5 clean sales exist
    - Returns None if no price can be determined
    """
    try:
        clean_url = tracker_url.split('?')[0]
        response = requests.get(clean_url, headers=HEADERS, timeout=30)
        response.raise_for_status()

        # Check for JS-based redirects by inspecting the page title
        # A valid product page has a specific title; a search/landing page does not
        soup_check = BeautifulSoup(response.text, "html.parser")
        page_title = soup_check.title.get_text(strip=True) if soup_check.title else ''

        # Landing/search pages have generic titles like "Hobby Box Card Prices"
        # or "Search Results". A real product page title contains the product name.
        bad_titles = ['search', 'results', 'too many results', 'card prices | hobby box card list']
        if any(bad in page_title.lower() for bad in bad_titles):
            print(f"  Landed on wrong page ('{page_title}') — skipping")
            return None

        soup = soup_check
        soup = BeautifulSoup(response.text, "html.parser")

        # Parse exclusion keywords (lowercase for case-insensitive matching)
        excludes = []
        if exclude_keywords:
            excludes = [k.strip().lower() for k in exclude_keywords.split(',') if k.strip()]

        # ── Scrape sold listings table ──────────────────────────────
        # Sold listings table has columns: Sale Date | TW | Title | Price
        # We look for rows with a date pattern and extract the price column
        date_pattern   = re.compile(r'^\d{4}-\d{2}-\d{2}$')
        dollar_pattern = re.compile(r'\$(\d{1,6}(?:,\d{3})*\.\d{2})')

        clean_prices = []

        tables = soup.find_all("table")
        for table in tables:
            rows = table.find_all("tr")
            for row in rows:
                cells = row.find_all("td")
                if len(cells) < 4:
                    continue

                # First cell should be a sale date
                date_text = cells[0].get_text(strip=True)
                if not date_pattern.match(date_text):
                    continue

                # Any cell after the first two could be title or price
                # Search all remaining cells for a price
                title_text = ''
                price = None
                for cell in cells[1:]:
                    text = cell.get_text(strip=True)
                    # Title cell is the longest text cell
                    if len(text) > len(title_text) and not dollar_pattern.search(text):
                        title_text = text.lower()
                    match = dollar_pattern.search(text)
                    if match and price is None:
                        candidate = float(match.group(1).replace(',', ''))
                        if 10 < candidate < 50000:
                            price = candidate

                if not title_text and not price:
                    continue

                # Skip if any exclusion keyword found in title
                if any(ex in title_text for ex in excludes):
                    continue

                if price is not None:
                    clean_prices.append(price)

            if clean_prices:
                break  # Stop after first table with valid sales data

        # ── Compute median ──────────────────────────────────────────
        if len(clean_prices) >= 5:
            sample = clean_prices[:10]  # Last 10 (most recent first)
            sample_sorted = sorted(sample)
            mid = len(sample_sorted) // 2
            if len(sample_sorted) % 2 == 0:
                median = (sample_sorted[mid - 1] + sample_sorted[mid]) / 2
            else:
                median = sample_sorted[mid]
            print(f"  Median of {len(sample)} clean sales: ${median:.2f}")
            return round(median, 2)

        # ── Fallback: their market price ────────────────────────────
        if len(clean_prices) > 0:
            print(f"  Only {len(clean_prices)} clean sale(s) found — falling back to market price")
        else:
            print(f"  No clean sales found — falling back to market price")

        # Re-parse the page fresh for the market price
        # (can't reuse soup — nav stripping may have corrupted it)
        response2 = requests.get(clean_url, headers=HEADERS, timeout=30)
        soup2 = BeautifulSoup(response2.text, "html.parser")
        page_title2 = soup2.title.get_text(strip=True) if soup2.title else ''
        if any(bad in page_title2.lower() for bad in bad_titles):
            print(f"  Fallback landed on wrong page — skipping")
            return None

        # Remove nav/header/footer/ul so $6/month links don't pollute
        for tag in soup2.find_all(['nav', 'header', 'footer', 'ul']):
            tag.decompose()

        for table in soup2.find_all("table"):
            for cell in table.find_all("td"):
                text = cell.get_text(strip=True)
                match = dollar_pattern.search(text)
                if match:
                    price = float(match.group(1).replace(',', ''))
                    if 10 < price < 50000:
                        print(f"  Market price fallback: ${price:.2f}")
                        return round(price, 2)

        print(f"  No price found at all")
        return None

    except Exception as e:
        print(f"  SportsCardsPro scrape error: {e}")
        return None

# ======================================
# PROCESS EACH ROW
# ======================================

snapshot_rows = []

for row_num in range(2, len(rows) + 1):

    try:
        row = rows[row_num - 1]

        # Pad short rows
        while len(row) < 16:
            row.append("")

        waxstat_url    = row[COL_WAXSTAT_URL].strip()
        tracker_url    = row[COL_TRACKER_URL].strip()
        last_updated   = row[COL_LAST_UPDATED].strip()
        product        = row[COL_PRODUCT].strip()
        box_type       = row[COL_BOX_TYPE].strip()
        exclude_kw     = row[COL_EXCLUDE_KW].strip()

        if not product:
            continue

        # Skip rows already updated today
        if SKIP_UPDATED_TODAY and last_updated:
            if last_updated.startswith(today_str):
                print(f"Row {row_num}: Already updated today — skipping")
                continue

        updated_something = False
        now_str = datetime.now().strftime("%Y-%m-%d %H:%M")

        # ======================================
        # WAXSTAT SCRAPE
        # ======================================

        if waxstat_url:
            print(f"Row {row_num}: Scraping WaxStat — {product}")
            try:
                response = requests.get(waxstat_url, headers=HEADERS, timeout=30)
                response.raise_for_status()
                soup = BeautifulSoup(response.text, "html.parser")

                avg_price = None
                labels = soup.find_all("div", class_="text-grey")
                for label in labels:
                    if "Average market price" in label.get_text():
                        price_div = label.find_next("div", class_="price")
                        if price_div:
                            avg_price = float(
                                price_div.get_text()
                                .replace("$", "")
                                .replace(",", "")
                            )
                            break

                if avg_price is not None:
                    ninety_value = round(avg_price * 0.90, 2)
                    sheet.update(
                        f"G{row_num}:H{row_num}",
                        [[avg_price, ninety_value]]
                    )
                    print(f"  WaxStat updated | Avg=${avg_price:.2f}")
                    updated_something = True

                    # Queue history snapshot
                    if is_snapshot_day:
                        key = (today_str, product, box_type)
                        if key not in logged_keys:
                            snapshot_rows.append([today_str, product, box_type, avg_price])
                            logged_keys.add(key)
                            print(f"  Queued history snapshot")
                else:
                    print(f"  WaxStat: price not found")

            except Exception as e:
                print(f"  WaxStat error: {e}")

            time.sleep(1)

        # ======================================
        # SPORTSCARDSPRO SCRAPE
        # ======================================

        if tracker_url:
            print(f"Row {row_num}: Scraping SportsCardsPro — {product}")
            ebay_price = get_sportscards_price(tracker_url, exclude_kw or None)

            if ebay_price is not None:
                sheet.update(
                    f"O{row_num}",
                    [[ebay_price]]
                )
                print(f"  SportsCardsPro updated | Price=${ebay_price:.2f}")
                updated_something = True
            else:
                print(f"  SportsCardsPro: price not found")

            time.sleep(1)

        # ======================================
        # UPDATE LAST UPDATED TIMESTAMP
        # ======================================

        if updated_something:
            sheet.update(f"L{row_num}", [[now_str]])

    except Exception as e:
        print(f"Row {row_num} failed: {e}")

# ======================================
# WRITE HISTORY BATCH
# ======================================

if snapshot_rows:
    history_sheet.append_rows(snapshot_rows, value_input_option="RAW")
    print(f"\nWrote {len(snapshot_rows)} snapshot rows to '{HISTORY_SHEET}'")

print("Done.")
