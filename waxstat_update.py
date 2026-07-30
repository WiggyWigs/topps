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

def get_sportscards_price(tracker_url):
    """
    Scrapes the Ungraded market price from a SportsCardsPro product page.
    Returns a float price or None if not found.
    """
    try:
        clean_url = tracker_url.split('?')[0]
        response = requests.get(clean_url, headers=HEADERS, timeout=30)
        response.raise_for_status()
        soup = BeautifulSoup(response.text, "html.parser")

        dollar_pattern = re.compile(r'\$(\d{1,6}(?:,\d{3})*\.\d{2})')

        # Remove nav/header/footer so their $6/month links don't pollute results
        for tag in soup.find_all(['nav', 'header', 'footer', 'ul']):
            tag.decompose()

        # The main price table is the first <table> on the page after cleanup
        # It contains the Ungraded / Grade 7 / Grade 8 etc columns
        tables = soup.find_all("table")
        for table in tables:
            cells = table.find_all("td")
            for cell in cells:
                text = cell.get_text(strip=True)
                match = dollar_pattern.search(text)
                if match:
                    price = float(match.group(1).replace(',', ''))
                    # Sanity check: real box prices are between $10 and $50,000
                    if 10 < price < 50000:
                        return price

        print(f"  No price found in tables")
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
        while len(row) < 15:
            row.append("")

        waxstat_url  = row[COL_WAXSTAT_URL].strip()
        tracker_url  = row[COL_TRACKER_URL].strip()
        last_updated = row[COL_LAST_UPDATED].strip()
        product      = row[COL_PRODUCT].strip()
        box_type     = row[COL_BOX_TYPE].strip()

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
            ebay_price = get_sportscards_price(tracker_url)

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
