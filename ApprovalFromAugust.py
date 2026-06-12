import os
import json
import time
import requests
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from gspread.exceptions import WorksheetNotFound
from requests.exceptions import ConnectionError as RequestsConnectionError

# ── CONFIG ────────────────────────────────────────────────────────────────────
BASE_URL   = "https://pw.jotform.com/API"
FORM_ID    = os.environ["JOTFORM_FORM_ID"]
API_KEY    = os.environ["JOTFORM_API_KEY"]
START_DATE = os.environ.get("START_DATE", "2024-01-01 00:00:00")
SPREADSHEET_NAME = os.environ.get("SPREADSHEET_NAME",'STN')     
WORKSHEET_NAME   = os.environ.get("WORKSHEET_NAME",'Approval Status')      
CREDENTIALS  = os.environ.get('GOOGLE_CREDENTIALS_JSON', 'credentials.json')

PAGE_SIZE           = 300
SLEEP_BETWEEN_CALLS = 1
MAX_PAGES           = 500
WRITE_BATCH_SIZE    = 500

# ── GOOGLE SHEETS ─────────────────────────────────────────────────────────────
scope = [
    'https://spreadsheets.google.com/feeds',
    'https://www.googleapis.com/auth/drive'
]

if not os.path.exists(CREDENTIALS):
    raise FileNotFoundError(f"Credentials file not found: {CREDENTIALS}")

with open(CREDENTIALS, 'r') as f:
    creds_dict = json.load(f)

creds  = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
client = gspread.authorize(creds)

spreadsheet = client.open(SPREADSHEET_NAME)

try:
    sheet = spreadsheet.worksheet(WORKSHEET_NAME)
    print(f"✅ Found existing worksheet '{WORKSHEET_NAME}' — clearing it for fresh load")
    sheet.clear()
except WorksheetNotFound:
    sheet = spreadsheet.add_worksheet(title=WORKSHEET_NAME, rows=1000, cols=10)
    print(f"➕ Created new worksheet '{WORKSHEET_NAME}'")

HEADERS = ['Unique ID', 'Created at', 'Updated at', 'Approval Status']
sheet.update(range_name='A1', values=[HEADERS])
print(f"✔ Headers written: {HEADERS}")

# ── HELPERS ───────────────────────────────────────────────────────────────────
def fetch_submissions(offset=0, limit=300):
    url    = f"{BASE_URL}/form/{FORM_ID}/submissions"
    params = {
        'apiKey':             API_KEY,
        'limit':              limit,
        'offset':             offset,
        'orderby':            'created_at',
        'direction':          'ASC',
        'addWorkflowStatus':  1,
        'filter':             json.dumps({'created_at:gt': START_DATE}),
    }
    for attempt in range(5):
        try:
            resp = requests.get(url, params=params, timeout=60)
            resp.raise_for_status()
            data = resp.json()
            if data.get('responseCode') != 200:
                raise Exception(f"JotForm API error: {data}")
            return data.get('content', [])
        except (RequestsConnectionError, requests.exceptions.Timeout) as e:
            wait = 5 * (attempt + 1)
            print(f"⚠️  Fetch error (attempt {attempt+1}/5), retrying in {wait}s... [{e}]")
            time.sleep(wait)
        except requests.exceptions.HTTPError as e:
            if e.response is not None and e.response.status_code == 429:
                wait = 15 * (attempt + 1)
                print(f"⚠️  Rate limited, retrying in {wait}s...")
                time.sleep(wait)
            else:
                raise
    raise Exception(f"❌ Failed to fetch submissions at offset {offset}")


def extract_unique_id(answers):
    for _, meta in answers.items():
        if meta.get('name') == 'uniqueId' or meta.get('text') == 'Unique ID':
            return meta.get('answer', '')
    return ''


def append_with_retry(sheet, batch, retries=3):
    for attempt in range(retries):
        try:
            sheet.append_rows(batch, value_input_option='RAW')
            return
        except Exception as e:
            if attempt < retries - 1:
                wait = 5 * (attempt + 1)
                print(f"⚠️  Write failed (attempt {attempt+1}/{retries}), retrying in {wait}s... [{e}]")
                time.sleep(wait)
            else:
                raise

# ── FETCH ALL DATA FIRST ──────────────────────────────────────────────────────
all_rows = []
offset   = 0
page     = 0

print(f"\n🚀 Starting initial load from {START_DATE} ...")

while page < MAX_PAGES:
    submissions = fetch_submissions(offset=offset, limit=PAGE_SIZE)

    if not submissions:
        print("✔ No more submissions.")
        break

    for sub in submissions:
        answers = sub.get('answers', {})
        all_rows.append([
            extract_unique_id(answers),
            sub.get('created_at', ''),
            sub.get('updated_at', ''),
            sub.get('workflowStatus', ''),
        ])

    offset += PAGE_SIZE
    page   += 1
    print(f"✔ Page {page} pulled | {len(all_rows)} rows total so far")
    time.sleep(SLEEP_BETWEEN_CALLS)

# ── SORT OLDEST FIRST ─────────────────────────────────────────────────────────
all_rows.sort(key=lambda r: r[1])  # sort by 'Created at' (ISO-like timestamps sort correctly as strings)

print(f"\n📊 Total rows fetched: {len(all_rows)} — sorted oldest first")

# ── WRITE IN BATCHES ───────────────────────────────────────────────────────────
total_written = 0
for i in range(0, len(all_rows), WRITE_BATCH_SIZE):
    batch = all_rows[i:i + WRITE_BATCH_SIZE]
    append_with_retry(sheet, batch)
    total_written += len(batch)
    print(f"📝 Written {total_written}/{len(all_rows)} rows so far...")
    time.sleep(2)

print(f"\n✅ DONE — {total_written} rows written to '{SPREADSHEET_NAME}' → '{WORKSHEET_NAME}'")