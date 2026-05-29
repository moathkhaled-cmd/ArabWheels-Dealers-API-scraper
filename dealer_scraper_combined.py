"""
ArabWheels — Multi-Dealer Scraper  (Combined v1)
=================================================
Dealers covered
---------------
  1. Alba Cars     — REST API  (albacars.ae)
  2. GTA Cars      — WP REST API (gtacars.ae)
  3. AutoMax       — REST API  (automaxgroup.me) — auto-probes pagination

For each dealer the script:
  • Fetches all live listings
  • Reconciles against the dealer's previous sheet tab (new / updated / unchanged / removed)
  • Dumps everything into Google Sheets:
        Sheet name  → <DealerName> <YYYY-MM-DD>   e.g. "Alba Cars 2026-05-29"
        The target spreadsheet ID is set in CONFIG below.
  • Sends a Gmail summary email via a Google Apps Script Web App URL (also in CONFIG).

Setup
-----
  1. Create a Google Cloud service account, share the target spreadsheet with its email,
     and download the JSON key → set SERVICE_ACCOUNT_JSON below.
  2. Deploy the Apps Script email sender (see companion script email_sender.gs)
     and paste its Web App URL → APPS_SCRIPT_EMAIL_URL below.
  3. pip install google-auth google-auth-httplib2 google-api-python-client requests
  4. python dealer_scraper_combined.py

Scheduling (run daily)
----------------------
  Linux/macOS cron example:
    0 7 * * * /usr/bin/python3 /path/to/dealer_scraper_combined.py >> /var/log/dealer_scraper.log 2>&1
  Windows Task Scheduler: point to pythonw.exe + script path.
"""

# ── stdlib ────────────────────────────────────────────────────────────────────
import html
import json
import os
import re
import sys
import time
import urllib.request
import urllib.parse
import urllib.error
from datetime import date, datetime

# ── third-party ───────────────────────────────────────────────────────────────
try:
    import requests
except ImportError:
    sys.exit("Missing 'requests'. Run: pip install requests")

try:
    from google.oauth2 import service_account
    from googleapiclient.discovery import build
    from googleapiclient.errors import HttpError
except ImportError:
    sys.exit(
        "Missing Google client libs. Run:\n"
        "  pip install google-auth google-auth-httplib2 google-api-python-client"
    )


# =============================================================================
# ██████╗ ██████╗ ███╗   ██╗███████╗██╗ ██████╗
# ██╔════╝██╔═══██╗████╗  ██║██╔════╝██║██╔════╝
# ██║     ██║   ██║██╔██╗ ██║█████╗  ██║██║  ███╗
# ██║     ██║   ██║██║╚██╗██║██╔══╝  ██║██║   ██║
# ╚██████╗╚██████╔╝██║ ╚████║██║     ██║╚██████╔╝
#  ╚═════╝ ╚═════╝ ╚═╝  ╚═══╝╚═╝     ╚═╝ ╚═════╝
# =============================================================================

CONFIG = {
    # ── Google Sheets ─────────────────────────────────────────────────────────
    # Reads from env var SPREADSHEET_ID when running in GitHub Actions.
    # For local runs, either set the env var or replace the fallback string.
    "SPREADSHEET_ID": os.environ.get("SPREADSHEET_ID", "YOUR_SPREADSHEET_ID_HERE"),

    # Path to the service account JSON key file.
    # GitHub Actions writes this file from the GOOGLE_SERVICE_ACCOUNT_JSON secret.
    "SERVICE_ACCOUNT_JSON": "service_account.json",

    # ── Email (Google Apps Script Web App) ────────────────────────────────────
    # Reads from env var APPS_SCRIPT_EMAIL_URL when running in GitHub Actions.
    # For local runs, either set the env var or replace the fallback string.
    "APPS_SCRIPT_EMAIL_URL": os.environ.get(
        "APPS_SCRIPT_EMAIL_URL", "YOUR_APPS_SCRIPT_WEB_APP_URL_HERE"
    ),

    # ── General ───────────────────────────────────────────────────────────────
    "REQUEST_TIMEOUT": 30,
    "RETRY_ATTEMPTS":  3,
    "REQUEST_DELAY":   0.3,   # seconds between paginated API calls
}


# =============================================================================
# GOOGLE SHEETS CLIENT
# =============================================================================

SHEETS_SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]


def get_sheets_service():
    creds = service_account.Credentials.from_service_account_file(
        CONFIG["SERVICE_ACCOUNT_JSON"], scopes=SHEETS_SCOPES
    )
    return build("sheets", "v4", credentials=creds, cache_discovery=False)


# =============================================================================
# SHARED UTILITIES
# =============================================================================

def strip_html(raw: str) -> str:
    if not raw:
        return ""
    text = re.sub(r"<br\s*/?>", " ", raw, flags=re.IGNORECASE)
    text = re.sub(r"</?p>", " ", text, flags=re.IGNORECASE)
    text = re.sub(r"<[^>]+>", "", text)
    text = html.unescape(text)
    return re.sub(r"\s+", " ", text).strip()


def fetch_json_urllib(url: str, retries: int = 3, delay: float = 2.0):
    """urllib-based fetch (no requests dependency for WP API calls)."""
    for attempt in range(1, retries + 1):
        try:
            req = urllib.request.Request(
                url, headers={"User-Agent": "ArabWheels-Scraper/1.0"}
            )
            with urllib.request.urlopen(req, timeout=CONFIG["REQUEST_TIMEOUT"]) as resp:
                return json.loads(resp.read().decode("utf-8")), resp.headers
        except Exception as exc:
            print(f"    Attempt {attempt}/{retries} failed for {url}: {exc}")
            if attempt < retries:
                time.sleep(delay)
    return None, None


def fetch_json_requests(url: str, params: dict = None) -> dict | None:
    """requests-based fetch with exponential back-off."""
    for attempt in range(1, CONFIG["RETRY_ATTEMPTS"] + 1):
        try:
            resp = requests.get(url, params=params, timeout=CONFIG["REQUEST_TIMEOUT"])
            resp.raise_for_status()
            return resp.json()
        except requests.RequestException as exc:
            print(f"    Attempt {attempt}/{CONFIG['RETRY_ATTEMPTS']} failed: {exc}")
            if attempt < CONFIG["RETRY_ATTEMPTS"]:
                time.sleep(2 ** attempt)
    return None


# =============================================================================
# RECONCILIATION  (shared across all dealers)
# =============================================================================

def reconcile(new_rows: list[dict], previous: dict[str, dict], tracked: list[str]) -> list[dict]:
    """
    Assign new / updated / unchanged to new_rows; append removed rows.
    previous = {str(id): row_dict}
    """
    new_ids = {str(r["id"]) for r in new_rows}

    for row in new_rows:
        rid = str(row["id"])
        if rid not in previous:
            row["status"] = "new"
        else:
            changed = any(
                str(row.get(f, "")) != str(previous[rid].get(f, ""))
                for f in tracked
            )
            row["status"] = "updated" if changed else "unchanged"

    removed = []
    for rid, prev_row in previous.items():
        if rid not in new_ids:
            prev_row["status"] = "removed"
            removed.append(prev_row)

    return new_rows + removed


def count_statuses(rows: list[dict]) -> dict:
    counts = {"new": 0, "updated": 0, "unchanged": 0, "removed": 0}
    for r in rows:
        s = r.get("status", "")
        if s in counts:
            counts[s] += 1
    counts["total"] = len(rows)
    return counts


# =============================================================================
# GOOGLE SHEETS READ / WRITE
# =============================================================================

def sheet_tab_name(dealer_name: str, today: str) -> str:
    return f"{dealer_name} {today}"


def get_previous_rows(service, spreadsheet_id: str, dealer_name: str) -> dict[str, dict]:
    """
    Find the most recent tab for this dealer and load it as {id: row_dict}.
    Tabs are named '<DealerName> YYYY-MM-DD'; we sort and pick the last one.
    """
    try:
        meta = service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
        sheets = meta.get("sheets", [])
        dealer_tabs = sorted(
            [s["properties"]["title"] for s in sheets
             if s["properties"]["title"].startswith(dealer_name + " ")],
        )
    except HttpError as exc:
        print(f"  Could not list sheets: {exc}")
        return {}

    if not dealer_tabs:
        return {}

    latest = dealer_tabs[-1]
    print(f"  Previous tab found: '{latest}'")

    try:
        result = (
            service.spreadsheets()
            .values()
            .get(spreadsheetId=spreadsheet_id, range=f"'{latest}'!A1:ZZ")
            .execute()
        )
        values = result.get("values", [])
    except HttpError as exc:
        print(f"  Could not read previous tab: {exc}")
        return {}

    if not values or len(values) < 2:
        return {}

    headers = values[0]
    previous = {}
    for row_vals in values[1:]:
        # pad short rows
        row_vals += [""] * (len(headers) - len(row_vals))
        row_dict = dict(zip(headers, row_vals))
        ad_id = str(row_dict.get("id", "")).strip()
        if ad_id:
            previous[ad_id] = row_dict
    return previous


def write_to_sheet(service, spreadsheet_id: str, tab_name: str, columns: list[str], rows: list[dict]):
    """Create a new tab and write header + data rows."""
    # ── Create the tab ────────────────────────────────────────────────────────
    try:
        service.spreadsheets().batchUpdate(
            spreadsheetId=spreadsheet_id,
            body={
                "requests": [{
                    "addSheet": {
                        "properties": {"title": tab_name}
                    }
                }]
            },
        ).execute()
        print(f"  Created tab: '{tab_name}'")
    except HttpError as exc:
        if "already exists" in str(exc):
            print(f"  Tab '{tab_name}' already exists — will overwrite.")
        else:
            raise

    # ── Build values matrix ───────────────────────────────────────────────────
    header_row = columns
    data_rows  = []
    for row in rows:
        data_rows.append([str(row.get(col, "") or "") for col in columns])

    all_values = [header_row] + data_rows

    # ── Write in one call ─────────────────────────────────────────────────────
    service.spreadsheets().values().update(
        spreadsheetId=spreadsheet_id,
        range=f"'{tab_name}'!A1",
        valueInputOption="RAW",
        body={"values": all_values},
    ).execute()

    # ── Auto-resize columns ───────────────────────────────────────────────────
    meta = service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
    sheet_id = next(
        (s["properties"]["sheetId"] for s in meta["sheets"]
         if s["properties"]["title"] == tab_name),
        None,
    )
    if sheet_id is not None:
        service.spreadsheets().batchUpdate(
            spreadsheetId=spreadsheet_id,
            body={
                "requests": [{
                    "autoResizeDimensions": {
                        "dimensions": {
                            "sheetId": sheet_id,
                            "dimension": "COLUMNS",
                            "startIndex": 0,
                            "endIndex": len(columns),
                        }
                    }
                }]
            },
        ).execute()

    print(f"  Wrote {len(rows)} row(s) to '{tab_name}'.")


# =============================================================================
# ██████╗ ███████╗ █████╗ ██╗     ███████╗██████╗
# ██╔══██╗██╔════╝██╔══██╗██║     ██╔════╝██╔══██╗
# ██║  ██║█████╗  ███████║██║     █████╗  ██████╔╝
# ██║  ██║██╔══╝  ██╔══██║██║     ██╔══╝  ██╔══██╗
# ██████╔╝███████╗██║  ██║███████╗███████╗██║  ██║
# ╚═════╝ ╚══════╝╚═╝  ╚═╝╚══════╝╚══════╝╚═╝  ╚═╝
# ── ALBA CARS ────────────────────────────────────────────────────────────────
# =============================================================================

ALBA_API_BASE       = "https://backup-api.albacars.ae/api/v1/vehicle"
ALBA_IMAGE_CDN_BASE = "https://storage.albacars.ae/"
ALBA_PAGE_SIZE      = 50
ALBA_TRACKED        = ["status", "price", "mileage", "trim",
                       "exteriorColor", "interiorColor", "year", "images"]
ALBA_COLUMNS        = [
    "id", "refId", "referenceNumber",
    "year", "make", "model", "trim",
    "exteriorColor", "interiorColor",
    "mileage", "price", "fuelType",
    "transmission", "driveType", "engineCapacity", "cylinders",
    "doors", "seats", "vehicleSpec",
    "warranty", "warrantyYears",
    "serviceContract", "serviceContractYears",
    "marketingTitle", "vin", "chassisNumber",
    "dubiCarsAdId", "adStatus", "isBrandNew", "isOnSale",
    "images",
    "extracted_date", "status",
]


def alba_build_image_url(media_url: str) -> str:
    if media_url.startswith("http://") or media_url.startswith("https://"):
        return media_url
    return ALBA_IMAGE_CDN_BASE + media_url.lstrip("/")


def alba_collect_images(vehicle: dict) -> str:
    media = [m for m in (vehicle.get("media") or []) if m.get("mediaType") == "image"]
    media.sort(key=lambda m: (not m.get("isPrimary", False), m.get("mediaIndex", 999)))
    return ", ".join(
        alba_build_image_url(m["mediaUrl"]) for m in media if m.get("mediaUrl")
    )


def alba_parse_make_model(vehicle: dict) -> tuple[str, str]:
    ref   = vehicle.get("referenceNumber", "")
    parts = ref.split("-")
    title = vehicle.get("title", "")
    make  = parts[0].strip() if len(parts) >= 2 else title.split(" ", 1)[0]
    model = parts[1].strip() if len(parts) >= 2 else (title.split(" ", 1)[1] if " " in title else "")
    known_multi = ["Mercedes-Benz", "Land Rover", "Alfa Romeo", "Aston Martin", "Rolls-Royce"]
    for mm in known_multi:
        if title.startswith(mm):
            make  = mm
            rest  = title[len(mm):].strip()
            model = rest.split(" ")[0] if rest else model
            break
    return make, model


def alba_extract_row(vehicle: dict, today_str: str) -> dict:
    make, model = alba_parse_make_model(vehicle)
    price_raw   = vehicle.get("price")
    price       = f"AED {price_raw:,}" if price_raw is not None else ""
    return {
        "id":                   vehicle.get("id"),
        "refId":                vehicle.get("refId"),
        "referenceNumber":      vehicle.get("referenceNumber"),
        "year":                 vehicle.get("year"),
        "make":                 make,
        "model":                model,
        "trim":                 vehicle.get("trim", ""),
        "exteriorColor":        vehicle.get("color", ""),
        "interiorColor":        vehicle.get("interiorColor", ""),
        "mileage":              vehicle.get("mileage"),
        "price":                price,
        "fuelType":             vehicle.get("fuelType", ""),
        "transmission":         vehicle.get("transmission", ""),
        "driveType":            vehicle.get("driveType", ""),
        "engineCapacity":       vehicle.get("engineCapacity", ""),
        "cylinders":            vehicle.get("numberOfCylinders", ""),
        "doors":                vehicle.get("doors", ""),
        "seats":                vehicle.get("seats", ""),
        "vehicleSpec":          vehicle.get("vehicleSpec", ""),
        "warranty":             vehicle.get("warranty", ""),
        "warrantyYears":        vehicle.get("warrantyYears", ""),
        "serviceContract":      vehicle.get("serviceContract", ""),
        "serviceContractYears": vehicle.get("serviceContractYears", ""),
        "marketingTitle":       vehicle.get("marketingTitle", ""),
        "vin":                  vehicle.get("vin", ""),
        "chassisNumber":        vehicle.get("chassisNumber", ""),
        "dubiCarsAdId":         vehicle.get("dubiCarsAdId", ""),
        "adStatus":             vehicle.get("status", ""),
        "isBrandNew":           vehicle.get("isBrandNew", False),
        "isOnSale":             vehicle.get("isOnSale", False),
        "images":               alba_collect_images(vehicle),
        "extracted_date":       today_str,
        "status":               "",
    }


def run_alba(service, today_str: str) -> dict:
    print("\n" + "="*60)
    print("  ALBA CARS")
    print("="*60)

    # Fetch
    vehicles = []
    offset   = 0
    while True:
        data = fetch_json_requests(ALBA_API_BASE, params={"limit": ALBA_PAGE_SIZE, "offset": offset})
        if not data:
            print("  API fetch failed.")
            break
        payload = data.get("data", {})
        batch   = payload.get("result", [])
        total   = payload.get("total", len(batch))
        vehicles.extend(batch)
        print(f"  Fetched {len(vehicles)}/{total} vehicles …")
        if not batch or len(vehicles) >= total:
            break
        offset += ALBA_PAGE_SIZE
        time.sleep(CONFIG["REQUEST_DELAY"])

    if not vehicles:
        return {"dealer": "Alba Cars", "error": "No vehicles fetched", "counts": {}}

    new_rows = [alba_extract_row(v, today_str) for v in vehicles]

    # Reconcile
    previous = get_previous_rows(service, CONFIG["SPREADSHEET_ID"], "Alba Cars")
    if previous:
        all_rows = reconcile(new_rows, previous, ALBA_TRACKED)
    else:
        for r in new_rows:
            r["status"] = "new"
        all_rows = new_rows

    counts = count_statuses(all_rows)

    # Write
    tab = sheet_tab_name("Alba Cars", today_str)
    write_to_sheet(service, CONFIG["SPREADSHEET_ID"], tab, ALBA_COLUMNS, all_rows)

    return {"dealer": "Alba Cars", "counts": counts, "tab": tab}


# =============================================================================
# ██████╗ ████████╗ █████╗      ██████╗ █████╗ ██████╗ ███████╗
# ██╔════╝╚══██╔══╝██╔══██╗    ██╔════╝██╔══██╗██╔══██╗██╔════╝
# ██║  ███╗  ██║   ███████║    ██║     ███████║██████╔╝███████╗
# ██║   ██║  ██║   ██╔══██║    ██║     ██╔══██║██╔══██╗╚════██║
# ╚██████╔╝  ██║   ██║  ██║    ╚██████╗██║  ██║██║  ██║███████║
#  ╚═════╝   ╚═╝   ╚═╝  ╚═╝     ╚═════╝╚═╝  ╚═╝╚═╝  ╚═╝╚══════╝
# ── GTA CARS ─────────────────────────────────────────────────────────────────
# =============================================================================

GTA_SITE_URL      = "https://gtacars.ae"
GTA_SHOWROOMS_API = f"{GTA_SITE_URL}/wp-json/wp/v2/gta-showrooms"
GTA_MEDIA_API     = f"{GTA_SITE_URL}/wp-json/wp/v2/media"
GTA_PER_PAGE      = 16
GTA_BATCH_SIZE    = 100
GTA_TRACKED       = ["price_aed", "mileage_km", "color", "specs",
                     "warranty", "service_history", "title", "status_publish"]
GTA_COLUMNS       = [
    "id", "title", "link", "year", "make", "model", "trim",
    "body_type", "color", "specs", "mileage_km",
    "price_aed", "monthly_installment_aed",
    "fuel", "cylinders", "transmission",
    "warranty", "service_history", "service_contract",
    "description", "all_image_urls",
    "phone", "whatsapp",
    "date_posted", "date_modified", "status_publish",
    "extracted_date", "status",
]


def gta_extract_trim(title, make, model, year):
    trim = title
    for token in [year, make, model]:
        if token:
            trim = re.sub(re.escape(str(token)), "", trim, flags=re.IGNORECASE)
    trim = trim.split(",")[0]
    return re.sub(r"\s+", " ", trim).strip(" -") or ""


def gta_collect_media_ids(ads: list) -> set:
    ids = set()
    for ad in ads:
        meta = ad.get("meta") or {}
        raw  = meta.get("choose-car-images", "")
        for part in str(raw).split(","):
            part = part.strip()
            if part.isdigit():
                ids.add(int(part))
    return ids


def gta_resolve_media(all_ids: set) -> dict:
    if not all_ids:
        return {}
    resolved = {}
    id_list  = sorted(all_ids)
    for i in range(0, len(id_list), GTA_BATCH_SIZE):
        batch  = id_list[i: i + GTA_BATCH_SIZE]
        params = urllib.parse.urlencode({
            "include":  ",".join(str(x) for x in batch),
            "per_page": len(batch),
            "_fields":  "id,source_url",
        })
        data, _ = fetch_json_urllib(f"{GTA_MEDIA_API}?{params}")
        if data and isinstance(data, list):
            for item in data:
                resolved[item["id"]] = item.get("source_url", "")
        if i + GTA_BATCH_SIZE < len(id_list):
            time.sleep(CONFIG["REQUEST_DELAY"])
    return resolved


def gta_build_image_cell(ad: dict, media_map: dict) -> str:
    urls = []
    hero = (ad.get("featured_image_url") or "").strip()
    if hero:
        urls.append(hero)
    meta = ad.get("meta") or {}
    raw  = meta.get("choose-car-images", "")
    for part in str(raw).split(","):
        part = part.strip()
        if part.isdigit():
            mid = int(part)
            url = media_map.get(mid, "")
            if url and url not in urls:
                urls.append(url)
            elif not url:
                ph = f"[media_id:{mid}]"
                if ph not in urls:
                    urls.append(ph)
    return ", ".join(urls)


def gta_extract_row(ad: dict, media_map: dict, today_str: str) -> dict:
    meta      = ad.get("meta") or {}
    year      = ad.get("year") or meta.get("select-year", "")
    make      = ad.get("make", "")
    model     = ad.get("model", "")
    body      = (ad.get("body") or meta.get("select-vehicle-type", "")).strip().lstrip("\ufeff")
    title_raw = ad.get("title") or {}
    title     = (
        strip_html(title_raw.get("rendered", ""))
        if isinstance(title_raw, dict)
        else str(title_raw)
    )
    guid = ad.get("guid") or {}
    link = ad.get("link") or (guid.get("rendered", "") if isinstance(guid, dict) else str(guid))
    return {
        "id":                      ad.get("id"),
        "title":                   title,
        "link":                    link,
        "year":                    year,
        "make":                    make,
        "model":                   model,
        "trim":                    gta_extract_trim(title, make, model, str(year)),
        "body_type":               body,
        "color":                   ad.get("color") or meta.get("car-color", ""),
        "specs":                   meta.get("select-specs", ""),
        "mileage_km":              meta.get("kilometer-reading") or ad.get("km", ""),
        "price_aed":               meta.get("car-price", ""),
        "monthly_installment_aed": meta.get("car-monthly-installmentamount", ""),
        "fuel":                    meta.get("select-fuel-type", ""),
        "cylinders":               meta.get("number-of-cylinders", ""),
        "transmission":            meta.get("select-gear-box", ""),
        "warranty":                meta.get("_warranty", ""),
        "service_history":         meta.get("_servicehistory", ""),
        "service_contract":        meta.get("_servicecontract", ""),
        "description":             strip_html(meta.get("car-description", "")),
        "all_image_urls":          gta_build_image_cell(ad, media_map),
        "phone":                   meta.get("_phone", ""),
        "whatsapp":                meta.get("_whatsappnumber", ""),
        "date_posted":             (ad.get("date") or "")[:10],
        "date_modified":           (ad.get("modified") or "")[:10],
        "status_publish":          ad.get("status", ""),
        "extracted_date":          today_str,
        "status":                  "",
    }


def run_gta(service, today_str: str) -> dict:
    print("\n" + "="*60)
    print("  GTA CARS")
    print("="*60)

    # Paginate
    ads  = []
    page = 1
    while True:
        params = urllib.parse.urlencode({
            "_fields":  "id,guid,title,link,meta,featured_image_url,date,modified,status,year,make,model,body,color,km",
            "page":     page,
            "per_page": GTA_PER_PAGE,
        })
        url  = f"{GTA_SHOWROOMS_API}?{params}"
        data, hdrs = fetch_json_urllib(url)
        if not data or not isinstance(data, list) or len(data) == 0:
            break
        ads.extend(data)
        print(f"  Page {page}: {len(data)} listing(s) (total so far: {len(ads)})")

        total_pages = None
        if hdrs:
            tp = hdrs.get("X-WP-TotalPages") or hdrs.get("x-wp-totalpages")
            if tp:
                try:
                    total_pages = int(tp)
                except ValueError:
                    pass
        if total_pages is not None:
            if page >= total_pages:
                break
        elif len(data) < GTA_PER_PAGE:
            break
        page += 1
        time.sleep(CONFIG["REQUEST_DELAY"])

    if not ads:
        return {"dealer": "GTA Cars", "error": "No listings fetched", "counts": {}}

    # Resolve media
    media_ids = gta_collect_media_ids(ads)
    media_map = gta_resolve_media(media_ids) if media_ids else {}

    new_rows  = [gta_extract_row(ad, media_map, today_str) for ad in ads]

    # Reconcile
    previous = get_previous_rows(service, CONFIG["SPREADSHEET_ID"], "GTA Cars")
    if previous:
        all_rows = reconcile(new_rows, previous, GTA_TRACKED)
    else:
        for r in new_rows:
            r["status"] = "new"
        all_rows = new_rows

    counts = count_statuses(all_rows)

    # Write
    tab = sheet_tab_name("GTA Cars", today_str)
    write_to_sheet(service, CONFIG["SPREADSHEET_ID"], tab, GTA_COLUMNS, all_rows)

    return {"dealer": "GTA Cars", "counts": counts, "tab": tab}


# =============================================================================
#  █████╗ ██╗   ██╗████████╗ ██████╗ ███╗   ███╗ █████╗ ██╗  ██╗
# ██╔══██╗██║   ██║╚══██╔══╝██╔═══██╗████╗ ████║██╔══██╗╚██╗██╔╝
# ███████║██║   ██║   ██║   ██║   ██║██╔████╔██║███████║ ╚███╔╝
# ██╔══██║██║   ██║   ██║   ██║   ██║██║╚██╔╝██║██╔══██║ ██╔██╗
# ██║  ██║╚██████╔╝   ██║   ╚██████╔╝██║ ╚═╝ ██║██║  ██║██╔╝ ██╗
# ╚═╝  ╚═╝ ╚═════╝    ╚═╝    ╚═════╝ ╚═╝     ╚═╝╚═╝  ╚═╝╚═╝  ╚═╝
# ── AUTOMAX ───────────────────────────────────────────────────────────────────
# =============================================================================

AUTOMAX_API_URL  = "https://automaxgroup.me/api/cars"
AUTOMAX_TRACKED  = ["inStock", "price", "mileage", "trim",
                    "exteriorColor", "interiorColor", "badge", "description"]
AUTOMAX_COLUMNS  = [
    "inStock", "stockNumber", "catalogPdfUrl", "id",
    "year", "make", "model", "trim",
    "exteriorColor", "interiorColor",
    "mileage", "price", "description", "mainImages",
    "fuel", "engine", "cylinders", "transmission",
    "regionalSpec", "badge",
    "extracted_date", "status",
]
AUTOMAX_IMAGE_ORDER = [
    "front", "frontLeft", "frontRight",
    "back", "backRight", "leftSide",
    "dashboard", "steering",
]


def automax_collect_images(ad: dict) -> str:
    main = ad.get("mainImages") or {}
    urls = []
    for key in AUTOMAX_IMAGE_ORDER:
        entry = main.get(key)
        if entry and entry.get("url"):
            urls.append(entry["url"])
    for key, entry in main.items():
        if key not in AUTOMAX_IMAGE_ORDER and entry and entry.get("url"):
            urls.append(entry["url"])
    additional = [
        img["url"]
        for img in (ad.get("additionalImages") or [])
        if img.get("url")
    ]
    return ", ".join(urls + additional)


def automax_format_prices(ad: dict) -> str:
    price_map = {
        "USD": ad.get("price_usd"),
        "AED": ad.get("price_aed"),
        "SAR": ad.get("price_sar"),
        "KSA/DXB": ad.get("price_ksa_dxb"),
        "KSA+VAT": ad.get("price_ksa_vat"),
        "Base": ad.get("price"),
    }
    return " | ".join(f"{k}: {v}" for k, v in price_map.items() if v is not None)


def automax_extract_row(ad: dict, today_str: str) -> dict:
    specs = ad.get("specifications") or {}
    return {
        "id":             ad.get("id"),
        "inStock":        ad.get("inStock"),
        "stockNumber":    ad.get("stockNumber"),
        "catalogPdfUrl":  ad.get("catalogPdfUrl"),
        "year":           ad.get("year"),
        "make":           ad.get("make"),
        "model":          ad.get("model"),
        "trim":           ad.get("trim"),
        "exteriorColor":  ad.get("exteriorColor"),
        "interiorColor":  ad.get("interiorColor"),
        "mileage":        ad.get("mileage"),
        "price":          automax_format_prices(ad),
        "description":    (ad.get("description") or "").replace("\n", " "),
        "mainImages":     automax_collect_images(ad),
        "fuel":           ad.get("fuel"),
        "engine":         specs.get("engine"),
        "cylinders":      ad.get("cylinders"),
        "transmission":   ad.get("transmission"),
        "regionalSpec":   specs.get("regionalSpec"),
        "badge":          ad.get("badge"),
        "extracted_date": today_str,
        "status":         "",
    }


def automax_probe_and_fetch() -> list[dict]:
    """
    Auto-detect the API's pagination scheme:
      1. Single JSON list → return it.
      2. Wrapper object with data/cars/results key → return that list.
      3. Paginated (page param) → paginate until empty.
      4. Paginated (offset/limit) → paginate until empty.
    """
    print("  Probing AutoMax API …")
    raw = fetch_json_requests(AUTOMAX_API_URL)
    if raw is None:
        return []

    # ── Case 1: bare list ─────────────────────────────────────────────────────
    if isinstance(raw, list):
        print(f"  Single-response list: {len(raw)} ads.")
        return raw

    # ── Case 2: wrapper dict ──────────────────────────────────────────────────
    if isinstance(raw, dict):
        for key in ("data", "cars", "ads", "results", "items", "vehicles"):
            if key in raw and isinstance(raw[key], list):
                ads = raw[key]
                total = raw.get("total") or raw.get("totalCount") or raw.get("count")
                print(f"  Wrapper key '{key}' found: {len(ads)} ads (total reported: {total})")

                # If total > len(ads) it's probably paginated — try page-based
                if total and int(total) > len(ads):
                    print("  Looks paginated — trying page-based pagination …")
                    all_ads = list(ads)
                    page    = 2
                    while len(all_ads) < int(total):
                        batch = fetch_json_requests(AUTOMAX_API_URL, params={"page": page})
                        if not batch:
                            break
                        # Unwrap if needed
                        if isinstance(batch, dict) and key in batch:
                            batch = batch[key]
                        if not isinstance(batch, list) or len(batch) == 0:
                            break
                        all_ads.extend(batch)
                        print(f"  Page {page}: {len(all_ads)}/{total} …")
                        page += 1
                        time.sleep(CONFIG["REQUEST_DELAY"])
                    return all_ads

                return ads

        # Flat dict with numeric keys → values are ads
        values = list(raw.values())
        if all(isinstance(v, dict) for v in values):
            print(f"  Flat dict of ads: {len(values)}")
            return values

    print("  Unexpected AutoMax API structure — returning empty.")
    return []


def run_automax(service, today_str: str) -> dict:
    print("\n" + "="*60)
    print("  AUTOMAX")
    print("="*60)

    ads = automax_probe_and_fetch()
    if not ads:
        return {"dealer": "AutoMax", "error": "No ads fetched", "counts": {}}

    print(f"  Total ads: {len(ads)}")
    new_rows = [automax_extract_row(ad, today_str) for ad in ads]

    # Reconcile
    previous = get_previous_rows(service, CONFIG["SPREADSHEET_ID"], "AutoMax")
    if previous:
        all_rows = reconcile(new_rows, previous, AUTOMAX_TRACKED)
    else:
        for r in new_rows:
            r["status"] = "new"
        all_rows = new_rows

    counts = count_statuses(all_rows)

    # Write
    tab = sheet_tab_name("AutoMax", today_str)
    write_to_sheet(service, CONFIG["SPREADSHEET_ID"], tab, AUTOMAX_COLUMNS, all_rows)

    return {"dealer": "AutoMax", "counts": counts, "tab": tab}


# =============================================================================
# EMAIL  (via Google Apps Script Web App)
# =============================================================================

def send_summary_email(results: list[dict], today_str: str):
    """
    POST structured data to the Apps Script Web App.
    The Apps Script builds and sends the full professional HTML email.
    Payload matches what email_sender_v2.gs expects.
    """
    url = CONFIG["APPS_SCRIPT_EMAIL_URL"]
    if not url or url.startswith("YOUR_"):
        print("\n  [Email] No Apps Script URL configured — skipping email.")
        return

    spreadsheet_url = (
        f"https://docs.google.com/spreadsheets/d/{CONFIG['SPREADSHEET_ID']}/edit"
    )

    # Build the dealers list for the Apps Script
    dealers_payload = []
    for r in results:
        dealer_entry = {
            "name":     r.get("dealer", "Unknown"),
            "tab":      r.get("tab", ""),
            "sheetUrl": spreadsheet_url,
            "counts":   r.get("counts", {}),
            "error":    r.get("error", None),
        }
        dealers_payload.append(dealer_entry)

    run_time = datetime.now().strftime("%H:%M:%S")

    payload = {
        "runDate":  today_str,
        "runTime":  run_time,
        "sheetUrl": spreadsheet_url,
        "dealers":  dealers_payload,
    }

    try:
        resp = requests.post(url, json=payload, timeout=30)
        if resp.status_code == 200:
            print(f"\n  ✅  Summary email sent via Apps Script")
        else:
            print(f"\n  ⚠  Email Web App returned {resp.status_code}: {resp.text[:200]}")
    except Exception as exc:
        print(f"\n  ⚠  Email failed: {exc}")


# =============================================================================
# MAIN
# =============================================================================

def main():
    print("\n" + "█"*60)
    print("  ArabWheels Multi-Dealer Scraper")
    print(f"  {datetime.now().strftime('%Y-%m-%d  %H:%M:%S')}")
    print("█"*60)

    today_str = date.today().isoformat()

    # Validate config
    if CONFIG["SPREADSHEET_ID"].startswith("YOUR_"):
        sys.exit(
            "\nERROR: Set SPREADSHEET_ID in CONFIG before running.\n"
            "  Open this script and fill in the CONFIG dict at the top."
        )
    if not os.path.exists(CONFIG["SERVICE_ACCOUNT_JSON"]):
        sys.exit(
            f"\nERROR: Service account JSON not found: {CONFIG['SERVICE_ACCOUNT_JSON']}\n"
            "  Download it from Google Cloud Console and place it next to this script."
        )

    # Google Sheets service
    print("\nConnecting to Google Sheets …")
    service = get_sheets_service()
    print("  Connected ✅")

    # Run each dealer
    results = []
    results.append(run_alba(service, today_str))
    results.append(run_gta(service, today_str))
    results.append(run_automax(service, today_str))

    # Print summary table
    print("\n\n" + "="*60)
    print("  SUMMARY")
    print("="*60)
    print(f"  {'Dealer':<15} {'New':>6} {'Updated':>9} {'Unchanged':>11} {'Removed':>9} {'Total':>7}")
    print("  " + "-"*58)
    for r in results:
        dealer = r.get("dealer", "?")
        if "error" in r:
            print(f"  {dealer:<15}  ERROR: {r['error']}")
            continue
        c = r["counts"]
        print(
            f"  {dealer:<15}"
            f" {c.get('new',0):>6}"
            f" {c.get('updated',0):>9}"
            f" {c.get('unchanged',0):>11}"
            f" {c.get('removed',0):>9}"
            f" {c.get('total',0):>7}"
        )

    # Send email
    send_summary_email(results, today_str)

    print("\n✅  All done.\n")


if __name__ == "__main__":
    main()
