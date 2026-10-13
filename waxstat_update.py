from google.oauth2.service_account import Credentials
import gspread
import requests
from bs4 import BeautifulSoup
from datetime import datetime
import time
import os
import sys
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

# Sanity check threshold — must match the dashboard's ⚠️ divergence flag
DIVERGENCE_THRESHOLD = 0.40

# Sentinel returned by get_sportscards_price when the page loaded correctly
# but genuinely shows no price (vs. None = scrape failed / wrong page)
NO_PRICE_ON_PAGE = "NO_PRICE_ON_PAGE"

# Written to column Q when WaxStat shows N/A and column G holds a previous price
STALE_FLAG = "STALE"

# Fail the run if at least this share of WaxStat pages have no price label
LAYOUT_MISS_THRESHOLD = 0.50

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
COL_WAXSTAT_FLAG  = 16  # Q — "STALE" when G is a previous price (WaxStat showed N/A)

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

# Label the stale-price flag column the first time it's used
if rows and (len(rows[0]) <= COL_WAXSTAT_FLAG or not rows[0][COL_WAXSTAT_FLAG].strip()):
    sheet.update(range_name="Q1", values=[["WaxStat Status"]])

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
# HELPER: PARSE A SHEET CELL AS A NUMBER
# ======================================

def parse_num(value):
    """Converts a sheet cell like '$1,234.56' to a float. Returns 0.0 if blank/invalid."""
    try:
        return float(str(value).replace("$", "").replace(",", "").strip())
    except (ValueError, TypeError):
        return 0.0

# ======================================
# HELPER: PARSE WAXSTAT AVERAGE PRICE
# ======================================

PRICE_TEXT = re.compile(r'\$\s*(\d{1,3}(?:,\d{3})*(?:\.\d{2})?|\d+(?:\.\d{2})?)')

def _price_from_text(text):
    """'$1,067.49' -> 1067.49. Returns None for 'N/A' or anything without a $ amount."""
    match = PRICE_TEXT.search(text or "")
    return float(match.group(1).replace(",", "")) if match else None

def parse_waxstat_avg(soup):
    """
    Finds the average market price on a WaxStat box page. Locates values by
    their label text (not just CSS class) so a class rename alone won't break it.
    - Returns a float price when found
    - Returns NO_PRICE_ON_PAGE when the label is there but shows no price (N/A)
    - Returns None when no average-price label exists (layout changed again)
    """
    found_label = False
    candidates = []

    # Rail card (Oct 2026 layout): <div class="ws-lab">Average market price</div><div class="ws-big">$699.95</div>
    for label in soup.find_all(class_="ws-lab"):
        if "average market price" in label.get_text(" ", strip=True).lower():
            found_label = True
            value = label.find_next_sibling(class_="ws-big")
            if value:
                candidates.append(_price_from_text(value.get_text()))
            break

    # Price Stats tab (Oct 2026 layout): <p>Average Price</p><strong>$699.95</strong>
    for p in soup.select(".ws-stats p, .statistics p"):
        if p.get_text(" ", strip=True).lower() == "average price":
            found_label = True
            value = p.find_next_sibling("strong")
            if value:
                candidates.append(_price_from_text(value.get_text()))
            break

    # Pre-Oct 2026 layout, kept in case WaxStat rolls back
    for label in soup.find_all("div", class_="text-grey"):
        if "average market price" in label.get_text().lower():
            found_label = True
            value = label.find_next("div", class_="price")
            if value:
                candidates.append(_price_from_text(value.get_text()))
            break

    prices = [c for c in candidates if c is not None]
    if prices:
        if len(set(prices)) > 1:
            print(f"  WaxStat: average price sources disagree {prices} — using the first")
        return prices[0]
    if found_label:
        return NO_PRICE_ON_PAGE
    return None

# ======================================
# HELPER: SCRAPE SPORTSCARDSPRO PRICE
# ======================================

def get_sportscards_price(tracker_url, exclude_keywords=None):
    """
    Scrapes sold listings from a SportsCardsPro/PriceCharting product page.
    - Filters out listings whose titles contain any exclude_keywords
    - Returns median of last 10 clean sales (min 5)
    - Falls back to their market price if fewer than 5 clean sales exist
    - Returns NO_PRICE_ON_PAGE if the correct page loaded but shows no price
    - Returns None if the scrape failed (error, wrong page, redirect)
    """
    try:
        clean_url = tracker_url.split('?')[0]
        response = requests.get(clean_url, headers=HEADERS, timeout=30)
        response.raise_for_status()

        soup = BeautifulSoup(response.text, "html.parser")

        # Valid product pages have a canonical URL containing '/game/'
        # Landing/search pages redirect to /search-products or similar
        canonical = soup.find('link', rel='canonical')
        canonical_url = canonical['href'] if canonical else ''
        if '/game/' not in canonical_url:
            print(f"  Wrong page (canonical: {canonical_url}) — skipping")
            return None
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
        response2.raise_for_status()
        soup2 = BeautifulSoup(response2.text, "html.parser")
        canonical2 = soup2.find('link', rel='canonical')
        canonical_url2 = canonical2['href'] if canonical2 else ''
        if '/game/' not in canonical_url2:
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

        # Correct page loaded, but it has no sales and no market price
        print(f"  No price found at all")
        return NO_PRICE_ON_PAGE

    except Exception as e:
        print(f"  SportsCardsPro scrape error: {e}")
        return None

# ======================================
# PROCESS EACH ROW
# ======================================

snapshot_rows = []

# Track WaxStat parse health so a layout change fails the run instead of
# silently leaving column G stale
waxstat_attempts      = 0
waxstat_layout_misses = 0

for row_num in range(2, len(rows) + 1):

    try:
        row = rows[row_num - 1]

        # Pad short rows
        while len(row) < 17:
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

        # Reseller price used for the sanity check — starts as the stored
        # column G value, replaced below if WaxStat returns a fresh price
        reseller_price = parse_num(row[COL_WAXSTAT_AVG])

        # ======================================
        # WAXSTAT SCRAPE
        # ======================================

        if waxstat_url:
            print(f"Row {row_num}: Scraping WaxStat — {product}")
            try:
                response = requests.get(waxstat_url, headers=HEADERS, timeout=30)
                response.raise_for_status()
                soup = BeautifulSoup(response.text, "html.parser")

                waxstat_attempts += 1
                avg_price = parse_waxstat_avg(soup)

                if avg_price is None:
                    waxstat_layout_misses += 1
                    print(f"  WaxStat: average price label not found — page layout may have changed")
                elif avg_price == NO_PRICE_ON_PAGE:
                    if parse_num(row[COL_WAXSTAT_AVG]) > 0:
                        if row[COL_WAXSTAT_FLAG].strip() != STALE_FLAG:
                            sheet.update(range_name=f"Q{row_num}", values=[[STALE_FLAG]])
                        print(f"  WaxStat: page shows N/A — keeping previous price, flagged stale")
                    else:
                        print(f"  WaxStat: page shows no average price (N/A)")
                else:
                    ninety_value = round(avg_price * 0.90, 2)
                    sheet.update(
                        range_name=f"G{row_num}:H{row_num}",
                        values=[[avg_price, ninety_value]]
                    )
                    if row[COL_WAXSTAT_FLAG].strip():
                        sheet.update(range_name=f"Q{row_num}", values=[[""]])
                    print(f"  WaxStat updated | Avg=${avg_price:.2f}")
                    updated_something = True
                    reseller_price = avg_price

                    # Queue history snapshot
                    if is_snapshot_day:
                        key = (today_str, product, box_type)
                        if key not in logged_keys:
                            snapshot_rows.append([today_str, product, box_type, avg_price])
                            logged_keys.add(key)
                            print(f"  Queued history snapshot")

            except Exception as e:
                print(f"  WaxStat error: {e}")

            time.sleep(1)

        # ======================================
        # SPORTSCARDSPRO SCRAPE
        # ======================================

        if tracker_url:
            print(f"Row {row_num}: Scraping SportsCardsPro — {product}")
            ebay_price = get_sportscards_price(tracker_url, exclude_kw or None)

            if ebay_price == NO_PRICE_ON_PAGE:
                # Page shows no price — check the stored value against
                # the same 40% sanity check the dashboard uses
                stored_ebay = parse_num(row[COL_EBAY_PRICE])

                if stored_ebay > 0 and reseller_price > 0:
                    diff = abs(stored_ebay - reseller_price) / reseller_price
                    if diff >= DIVERGENCE_THRESHOLD:
                        sheet.update(range_name=f"O{row_num}", values=[[0]])
                        print(f"  Stored eBay price ${stored_ebay:.2f} is {diff*100:.0f}% off "
                              f"reseller ${reseller_price:.2f} — set to 0")
                        updated_something = True
                    else:
                        print(f"  Stored eBay price ${stored_ebay:.2f} within {diff*100:.0f}% "
                              f"of reseller — kept")
                else:
                    print(f"  No stored eBay price or no reseller price to compare — left as is")

            elif ebay_price is not None:
                sheet.update(
                    range_name=f"O{row_num}",
                    values=[[ebay_price]]
                )
                print(f"  SportsCardsPro updated | Price=${ebay_price:.2f}")
                updated_something = True

            else:
                print(f"  SportsCardsPro: scrape failed — stored value left as is")

            time.sleep(1)

        # ======================================
        # UPDATE LAST UPDATED TIMESTAMP
        # ======================================

        if updated_something:
            sheet.update(range_name=f"L{row_num}", values=[[now_str]])

    except Exception as e:
        print(f"Row {row_num} failed: {e}")

# ======================================
# WRITE HISTORY BATCH
# ======================================

if snapshot_rows:
    history_sheet.append_rows(snapshot_rows, value_input_option="RAW")
    print(f"\nWrote {len(snapshot_rows)} snapshot rows to '{HISTORY_SHEET}'")

print("Done.")

# ======================================
# FAIL LOUDLY ON LAYOUT CHANGE
# ======================================

if waxstat_attempts and waxstat_layout_misses / waxstat_attempts >= LAYOUT_MISS_THRESHOLD:
    print(f"\nERROR: WaxStat price label missing on {waxstat_layout_misses} of "
          f"{waxstat_attempts} pages — the page layout has likely changed. "
          f"Update parse_waxstat_avg().")
    sys.exit(1)
