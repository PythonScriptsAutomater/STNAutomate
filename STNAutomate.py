import os
import json
import time
import requests
import gspread
from datetime import datetime
from oauth2client.service_account import ServiceAccountCredentials
from gspread.exceptions import WorksheetNotFound
from requests.exceptions import ConnectionError as RequestsConnectionError

BASE_URL   = os.environ.get("JOTFORM_BASE_URL", "https://pw.jotform.com/API")
FORM_ID    = os.environ["JOTFORM_FORM_ID"]
API_KEY    = os.environ["JOTFORM_API_KEY"]
SPREADSHEET_NAME = os.environ.get("SPREADSHEET_NAME",'STN')     
WORKSHEET_NAME   = os.environ.get("WORKSHEET_NAME",'Approval Status')      
CREDENTIALS  = os.environ.get('GOOGLE_CREDENTIALS_JSON', 'credentials.json')
PAGE_SIZE           = 300
SLEEP_BETWEEN_CALLS = 1
MAX_PAGES           = 500
WRITE_BATCH_SIZE    = 500

HEADERS = ['Unique ID', 'Created at', 'Updated at', 'Approval Status']

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
    print(f"✅ Opened: '{SPREADSHEET_NAME}' → '{WORKSHEET_NAME}'")
except WorksheetNotFound:
    raise Exception(
        f"❌ Worksheet '{WORKSHEET_NAME}' not found. Run stock_initial_load.py first."
    )

# ── READ EXISTING DATA ────────────────────────────────────────────────────────
all_rows  = sheet.get_all_values()
if not all_rows:
    raise Exception("❌ Sheet is empty. Run stock_initial_load.py first.")

existing_headers = all_rows[0]
data_rows        = all_rows[1:]

if not data_rows:
    raise Exception("❌ Sheet has no data rows. Run stock_initial_load.py first.")

# Column indices (robust — looks up by header name)
try:
    uid_index  = existing_headers.index('Unique ID')
    date_index = existing_headers.index('Created at')
except ValueError as e:
    raise Exception(f"❌ Required column not found: {e}")

# ── DETERMINE CUTOFF ──────────────────────────────────────────────────────────
# Collect all Unique IDs already in the sheet and find the latest created_at.
# We normalise the date to 'YYYY-MM-DD HH:MM:SS' so it compares correctly
# against JotForm's created_at regardless of how Google Sheets formatted it.

def normalise_date(raw: str) -> str:
    raw = raw.strip()
    if not raw:
        return ''
    formats = [
        '%Y-%m-%d %H:%M:%S',
        '%Y-%m-%dT%H:%M:%S',
        '%Y-%m-%d',
        '%d/%m/%Y %H:%M:%S',
        '%d/%m/%Y',
        '%m/%d/%Y %H:%M:%S',
        '%m/%d/%Y',
        '%d-%m-%Y %H:%M:%S',
        '%d-%m-%Y',
    ]
    for fmt in formats:
        try:
            return datetime.strptime(raw, fmt).strftime('%Y-%m-%d %H:%M:%S')
        except ValueError:
            continue
    print(f"⚠️  Could not parse date: {repr(raw)}")
    return ''


known_unique_ids: set[str] = set()
last_known_date: str = ''

for row in data_rows:
    # Collect Unique IDs
    uid = row[uid_index].strip() if len(row) > uid_index else ''
    if uid:
        known_unique_ids.add(uid)

    # Track latest normalised date
    raw_date = row[date_index].strip() if len(row) > date_index else ''
    nd = normalise_date(raw_date)
    if nd and nd > last_known_date:
        last_known_date = nd

if not last_known_date:
    raise Exception("❌ Could not determine last known date from sheet.")

# Diagnostic: show raw date samples so any format mismatch is immediately visible
sample_raw = [
    row[date_index].strip()
    for row in data_rows[-5:]
    if len(row) > date_index and row[date_index].strip()
]
print(f"🔍 Sample raw dates from sheet (last 5) : {sample_raw}")
print(f"📌 Last known date (normalised)          : {last_known_date}")
print(f"📌 Known Unique IDs in sheet             : {len(known_unique_ids)}")

# ── HELPERS ───────────────────────────────────────────────────────────────────
def fetch_submissions_page(offset=0, limit=300):
    url    = f"{BASE_URL}/form/{FORM_ID}/submissions"
    params = {
        'apiKey':            API_KEY,
        'limit':             limit,
        'offset':            offset,
        'orderby':           'created_at',
        'direction':         'DESC',          # newest first → stop early
        'addWorkflowStatus': 1,
        # No date filter here — we control stopping ourselves
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


def is_new(sub: dict, uid: str) -> bool:
    """True if this submission is not already in the sheet."""
    created = sub.get('created_at', '')
    if created > last_known_date:
        return True
    if created == last_known_date:
        return uid not in known_unique_ids
    return False


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

# ── FETCH NEW SUBMISSIONS (DESC, stop at cutoff) ──────────────────────────────
offset        = 0
page          = 0
fetched_total = 0
all_new_subs: list[dict] = []

print(f"\n🚀 Fetching new submissions (newest-first)...")

while page < MAX_PAGES:
    submissions = fetch_submissions_page(offset=offset, limit=PAGE_SIZE)

    if not submissions:
        print("✔ No more submissions from JotForm.")
        break

    fetched_total += len(submissions)

    # Since direction=DESC: submissions[0] is newest, submissions[-1] is oldest
    newest_created = submissions[0].get('created_at', '')
    oldest_created = submissions[-1].get('created_at', '')

    if page == 0:
        print(f"🔍 First page — newest: {newest_created!r}, oldest: {oldest_created!r}, cutoff: {last_known_date!r}")

    # Entire page is older than cutoff → stop
    if newest_created < last_known_date:
        print(f"🛑 Entire page predates cutoff — stopping.")
        break

    # Filter this page to only new submissions
    new_on_page = []
    for sub in submissions:
        uid = extract_unique_id(sub.get('answers', {}))
        if is_new(sub, uid):
            new_on_page.append(sub)

    all_new_subs.extend(new_on_page)

    # If the oldest row on this page is already known, no point going further
    if oldest_created <= last_known_date:
        print(f"🎯 Boundary page: collected {len(new_on_page)} new rows — stopping.")
        break

    print(f"✔ Page {page+1}: {len(new_on_page)} new rows | {len(all_new_subs)} total new | scanned {fetched_total}")
    offset += PAGE_SIZE
    page   += 1
    time.sleep(SLEEP_BETWEEN_CALLS)

# ── WRITE NEW ROWS ────────────────────────────────────────────────────────────
if not all_new_subs:
    print("\n✅ Sheet is already up to date — no new submissions found.")
else:
    # Deduplicate (safety net for boundary overlaps)
    seen: set[str] = set()
    deduped: list[dict] = []
    for sub in all_new_subs:
        sid = sub.get('id', '')
        if sid not in seen:
            seen.add(sid)
            deduped.append(sub)

    # Sort oldest-first so sheet stays chronological
    deduped.sort(key=lambda s: s.get('created_at', ''))

    print(f"\n📦 {len(deduped)} new submissions to write...")

    rows_buffer   = []
    total_written = 0

    for sub in deduped:
        answers = sub.get('answers', {})
        rows_buffer.append([
            extract_unique_id(answers),
            sub.get('created_at', ''),
            sub.get('updated_at', ''),
            sub.get('workflowStatus', ''),
        ])

        if len(rows_buffer) >= WRITE_BATCH_SIZE:
            append_with_retry(sheet, rows_buffer)
            total_written += len(rows_buffer)
            print(f"📝 Written {total_written} / {len(deduped)} rows...")
            rows_buffer = []
            time.sleep(2)

    if rows_buffer:
        append_with_retry(sheet, rows_buffer)
        total_written += len(rows_buffer)

    print(f"\n✅ DONE — {total_written} new rows appended (scanned {fetched_total} total submissions)")