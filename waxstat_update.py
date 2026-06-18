from google.oauth2.service_account import Credentials
import gspread
import requests
from bs4 import BeautifulSoup
from datetime import datetime
import time
import os
import json

# ======================================
# CONFIG
# ======================================

SHEET_NAME = "topps-tracker"

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

# True = skip rows already updated today
# False = refresh everything

SKIP_UPDATED_TODAY = True

# ======================================
# CONNECT TO GOOGLE SHEET
# ======================================

# Load credentials from environment variable (GitHub Actions secret)
# Falls back to local file for running locally
creds_env = os.environ.get("GOOGLE_CREDENTIALS_JSON")

if creds_env:
    creds_dict = json.loads(creds_env)
    creds = Credentials.from_service_account_info(
        creds_dict,
        scopes=SCOPES
    )
    print("Loaded credentials from environment variable")
else:
    # Local fallback — uses the original JSON file
    CREDS_FILE = "topps-tracker-499814-874632f78fe4.json"
    creds = Credentials.from_service_account_file(
        CREDS_FILE,
        scopes=SCOPES
    )
    print("Loaded credentials from local file")

client = gspread.authorize(creds)

sheet = client.open(SHEET_NAME).sheet1

rows = sheet.get_all_values()

print(f"Found {len(rows)-1} rows")

# ======================================
# PROCESS EACH ROW
# ======================================

for row_num in range(2, len(rows) + 1):

    try:

        row = rows[row_num - 1]

        # Column K = URL
        url = row[10].strip() if len(row) > 10 else ""

        # Column L = Last Updated
        last_updated_cell = row[11].strip() if len(row) > 11 else ""

        if not url:
            print(f"Row {row_num}: No URL")
            continue

        # Skip rows already updated today

        if SKIP_UPDATED_TODAY and last_updated_cell:

            today = datetime.now().strftime("%Y-%m-%d")

            if last_updated_cell.startswith(today):

                print(
                    f"Row {row_num}: Already updated today"
                )

                continue

        print(f"Processing row {row_num}")

        response = requests.get(
            url,
            headers=HEADERS,
            timeout=30
        )

        response.raise_for_status()

        soup = BeautifulSoup(
            response.text,
            "html.parser"
        )

        avg_price = None

        labels = soup.find_all(
            "div",
            class_="text-grey"
        )

        for label in labels:

            if "Average market price" in label.get_text():

                price_div = label.find_next(
                    "div",
                    class_="price"
                )

                if price_div:

                    avg_price = float(
                        price_div.get_text()
                        .replace("$", "")
                        .replace(",", "")
                    )

                    break

        if avg_price is None:

            print(
                f"Row {row_num}: Average market price not found"
            )

            continue

        ninety_value = round(
            avg_price * 0.90,
            2
        )

        last_updated = datetime.now().strftime(
            "%Y-%m-%d %H:%M"
        )

        # ======================================
        # UPDATE SHEET
        # ======================================

        # G = WaxStat Avg
        # H = 90% Value

        sheet.update(
            f"G{row_num}:H{row_num}",
            [[
                avg_price,
                ninety_value
            ]]
        )

        # L = Last Updated

        sheet.update(
            f"L{row_num}",
            [[last_updated]]
        )

        print(
            f"Updated row {row_num} | Avg=${avg_price:.2f}"
        )

        # Avoid hammering WaxStat

        time.sleep(1)

    except Exception as e:

        print(
            f"Row {row_num} failed: {e}"
        )

print("Done.")
