"""
BigMint price tracker — web app backend.

Paste fresh bigmint.co cookies into the frontend textbox, hit Run, and this
fetches the current price for every URL in data/urls.csv, updates
data/tracker.xlsx (the running two-year tracker), regenerates a snapshot
image, and writes a CSV of the run. If the cookies are missing/expired, it
reports that clearly instead of guessing.

Prices come from bigmint.co's own JSON price-graph endpoint
(/prices_tg/graph/{itemID}/{currency}/{priceType}) via plain HTTP requests
with the session cookies — no headless browser involved. That endpoint is
what the price detail page itself calls client-side to draw its chart, and
its item ID/price type/currency are embedded in each tracked URL's slug.
Skipping a full browser (previously Playwright + Chromium) avoids the RAM a
headless browser needs, which was getting OOM-killed on Render's free tier.
"""

import asyncio
import json
import os
import re
import csv
import shutil
import traceback
from datetime import datetime, timezone, timedelta
from pathlib import Path

from curl_cffi.requests import AsyncSession
from fastapi import FastAPI, Request
from fastapi.responses import JSONResponse, FileResponse
import openpyxl
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt

BASE_DIR = Path(__file__).parent
DATA_DIR = BASE_DIR / "data"
STATIC_DIR = BASE_DIR / "static"
URLS_FILE = DATA_DIR / "urls.csv"
TRACKER_FILE = DATA_DIR / "tracker.xlsx"
BACKUP_DIR = DATA_DIR / "backups"
SNAPSHOT_DIR = DATA_DIR / "snapshots"
OUTPUT_CSV = DATA_DIR / "latest_output.csv"
COOKIES_FILE = DATA_DIR / "session_cookies.json"

EXCEL_SHEET = "Monthly"
EXCEL_URL_HEADER = "url"
EXCEL_CURRENT_HEADER = "current"
BIGMINT_HOME = "https://www.bigmint.co/"
MONTHS_BACK_IN_SNAPSHOT = 12

APP_PASSWORD = os.environ.get("APP_PASSWORD", "")  # optional shared-secret gate
IST = timezone(timedelta(hours=5, minutes=30))

app = FastAPI(title="BigMint Price Tracker")

BACKUP_DIR.mkdir(parents=True, exist_ok=True)
SNAPSHOT_DIR.mkdir(parents=True, exist_ok=True)

# Scraping runs as a background task instead of inline in the request handler
# because it can take longer than Render's proxy will hold an HTTP request
# open, which was truncating the JSON response. The frontend polls
# /api/run/status instead. Single in-memory job slot is fine — this app has
# one user running one scrape at a time.
JOB_STATE = {"status": "idle"}


# ---------------------------------------------------------------------------
# Helpers ported from the original bigmint_playwright.py / bigmint_screenshot.py
# ---------------------------------------------------------------------------

def clean_number(value):
    if value is None:
        return ""
    value = value.replace(",", "")
    m = re.search(r"(\d+(?:\.\d+)?)", value)
    return m.group(1) if m else ""


def map_same_site(value):
    if not value:
        return "Lax"
    value = str(value).lower()
    if value in ["no_restriction", "none"]:
        return "None"
    if value == "strict":
        return "Strict"
    return "Lax"


def normalize_cookies(raw):
    """Accepts a raw parsed cookie export (list, or {"cookies": [...]}) and
    returns a list of Playwright-format cookie dicts for *.bigmint.co only."""
    if isinstance(raw, dict):
        if isinstance(raw.get("cookies"), list):
            raw = raw["cookies"]
        else:
            raise ValueError("Unsupported cookie JSON shape — expected a list or {'cookies': [...]}")
    if not isinstance(raw, list):
        raise ValueError("Cookie JSON must be a list of cookie objects.")

    out = []
    for c in raw:
        try:
            domain = c.get("domain", "")
            if not domain or "bigmint.co" not in domain:
                continue
            entry = {
                "name": c["name"],
                "value": c["value"],
                "domain": domain,
                "path": c.get("path", "/"),
                "httpOnly": bool(c.get("httpOnly", False)),
                "secure": bool(c.get("secure", False)),
                "sameSite": map_same_site(c.get("sameSite")),
            }
            expiry = c.get("expirationDate") or c.get("expires")
            if expiry not in [None, "", 0]:
                try:
                    entry["expires"] = int(float(expiry))
                except Exception:
                    pass
            out.append(entry)
        except Exception:
            continue
    return out


def load_saved_cookies():
    if not COOKIES_FILE.exists():
        return None
    try:
        return json.loads(COOKIES_FILE.read_text(encoding="utf-8"))
    except Exception:
        return None


def save_cookies(cookies):
    COOKIES_FILE.write_text(json.dumps(cookies), encoding="utf-8")


def read_urls():
    urls = []
    with open(URLS_FILE, newline="", encoding="utf-8-sig") as f:
        reader = csv.DictReader(f)
        if reader.fieldnames and any((h or "").strip().lower() == "url" for h in reader.fieldnames):
            for row in reader:
                u = (row.get("url") or row.get("URL") or "").strip()
                if u:
                    urls.append(u)
        else:
            f.seek(0)
            for line in f:
                u = line.strip()
                if u and u.lower() != "url":
                    urls.append(u)
    return urls


def slug_from_url(url):
    return url.rstrip("/").split("/")[-1] or "unknown"


def text_looks_logged_in(text):
    """The login modal's markup (password field etc.) is present in the raw
    HTML even when you ARE logged in, so we can't just look for 'password'.
    The reliable signal is the personalised greeting that's server-rendered
    directly into the page for an authenticated session (e.g. "Hi, Mr ...")."""
    t = text.lower()
    return ("hi, mr" in t or "hi, ms" in t or "my portfolio" in t) and "enter your phone number" not in t


def text_looks_shield_blocked(text):
    """bigmint.co sits behind a Bunny Shield anti-bot check that can
    interstitial a request with an "Establishing a secure connection..."
    page instead of the real content."""
    t = text.lower()
    return "establishing" in t and "secure connection" in t


ITEM_URL_RE = re.compile(r"-(\d+)-([a-zA-Z])-([A-Za-z]{3})/?$")


def parse_item_from_url(url):
    """Tracked URLs look like
    .../prices/detail/pig-iron-dap-raipur-india-1102-f-INR — the trailing
    -{itemID}-{priceType}-{currency} maps directly onto bigmint.co's own
    price-graph API path."""
    m = ITEM_URL_RE.search(url.rstrip("/").split("?")[0])
    if not m:
        return None
    item_id, price_type, currency = m.groups()
    return {"item_id": item_id, "price_type": price_type, "currency": currency.upper()}



async def fetch_current_price(client, item_id, currency, price_type, market="ferrous"):
    url = f"https://www.bigmint.co/prices_tg/graph/{item_id}/{currency}/{price_type}"
    resp = await client.get(url, params={"market": market})
    resp.raise_for_status()
    data = resp.json()
    points = data.get("point") or (data.get("data") or {}).get("point") or []
    if not points:
        return None
    latest = max(points, key=lambda p: p[0])
    return latest[1]


def find_header_column(ws, header_row, predicate, max_col=60):
    for col in range(1, max_col + 1):
        val = ws.cell(row=header_row, column=col).value
        if val is not None and predicate(str(val)):
            return col
    return None


def find_month_column(ws, header_row, target_date, max_col=60):
    target_abbr = target_date.strftime("%b").lower()
    target_yy = target_date.strftime("%y")

    def is_target_month(text):
        s = text.strip().lower()
        m = re.match(r"[a-z]+", s)
        digits = re.findall(r"\d+", s)
        if not m or not digits:
            return False
        prefix = m.group()[:3]
        yy = digits[-1][-2:].zfill(2)
        return prefix == target_abbr and yy == target_yy

    return find_header_column(ws, header_row, is_target_month, max_col=max_col)


def update_tracker(results, today):
    if not TRACKER_FILE.exists():
        return {"updated": 0, "skipped": len(results), "not_found": list(results.keys()),
                "error": f"Tracker file not found at {TRACKER_FILE}"}

    wb = openpyxl.load_workbook(TRACKER_FILE)
    if EXCEL_SHEET not in wb.sheetnames:
        return {"updated": 0, "skipped": len(results), "not_found": [],
                "error": f"Sheet '{EXCEL_SHEET}' not found in tracker."}
    ws = wb[EXCEL_SHEET]

    url_col = find_header_column(ws, 1, lambda s: s.strip().lower() == EXCEL_URL_HEADER)
    current_col = find_header_column(ws, 1, lambda s: s.strip().lower().startswith(EXCEL_CURRENT_HEADER))
    month_col = find_month_column(ws, 1, today)

    if not url_col or not current_col or not month_col:
        return {"updated": 0, "skipped": len(results), "not_found": [],
                "error": f"Could not locate required columns (url_col={url_col}, current_col={current_col}, month_col={month_col})."}

    url_to_row = {}
    for r in range(2, ws.max_row + 1):
        u = ws.cell(row=r, column=url_col).value
        if u:
            url_to_row[str(u).strip()] = r

    backup_name = f"{TRACKER_FILE.stem}_{today.strftime('%Y%m%d_%H%M%S')}{TRACKER_FILE.suffix}"
    shutil.copy2(TRACKER_FILE, BACKUP_DIR / backup_name)

    updated, skipped, not_found = 0, 0, []
    for url, row in results.items():
        r = url_to_row.get(url.strip())
        if not r:
            not_found.append(url)
            skipped += 1
            continue
        if not row["current_price"]:
            skipped += 1
            continue
        try:
            val = float(row["current_price"])
            ws.cell(row=r, column=current_col).value = val
            ws.cell(row=r, column=month_col).value = val
            updated += 1
        except ValueError:
            skipped += 1

    wb.save(TRACKER_FILE)
    return {"updated": updated, "skipped": skipped, "not_found": not_found, "error": None}


def write_csv(results, today):
    fields = ["url", "item", "current_price", "status", "notes"]
    rows = []
    for url, data in results.items():
        rows.append({
            "url": url,
            "item": slug_from_url(url),
            "current_price": data["current_price"],
            "status": data["status"],
            "notes": data["notes"],
        })
    with open(OUTPUT_CSV, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=fields)
        writer.writeheader()
        writer.writerows(rows)
    dated = DATA_DIR / f"bigmint_prices_output_{today.strftime('%Y-%m-%d')}.csv"
    shutil.copy2(OUTPUT_CSV, dated)


def generate_snapshot(today):
    wb = openpyxl.load_workbook(TRACKER_FILE, data_only=True)
    ws = wb[EXCEL_SHEET]

    last_row = ws.max_row
    while last_row > 1 and ws.cell(row=last_row, column=1).value is None:
        last_row -= 1

    month_col = find_month_column(ws, 1, today)
    current_col = find_header_column(ws, 1, lambda s: s.strip().lower().startswith(EXCEL_CURRENT_HEADER))
    if not month_col or not current_col:
        return None

    start_col = max(2, month_col - MONTHS_BACK_IN_SNAPSHOT)
    month_cols = list(range(start_col, month_col + 1))
    display_cols = [1] + month_cols
    if current_col not in display_cols:
        display_cols.append(current_col)

    headers = [ws.cell(row=1, column=c).value or "" for c in display_cols]
    rows = []
    for r in range(2, last_row + 1):
        material = ws.cell(row=r, column=1).value
        if material is None or str(material).strip() == "":
            continue

        def fmt(val):
            if val is None:
                return ""
            try:
                v = float(str(val).replace(",", ""))
                return f"{v:,.0f}" if v.is_integer() else f"{v:,.2f}"
            except (ValueError, TypeError):
                return str(val)

        rows.append([fmt(ws.cell(row=r, column=c).value) for c in display_cols])

    n_rows = len(rows) + 1
    n_cols = len(display_cols)
    fig_width = max(10, n_cols * 1.3 + 2)
    fig_height = max(2, n_rows * 0.4)

    fig, ax = plt.subplots(figsize=(fig_width, fig_height))
    ax.axis("off")
    table = ax.table(cellText=rows, colLabels=headers, loc="center", cellLoc="right")
    table.auto_set_font_size(False)
    table.set_fontsize(9)
    table.scale(1, 1.6)
    table.auto_set_column_width(list(range(n_cols)))

    for col_idx in range(n_cols):
        cell = table[0, col_idx]
        cell.set_facecolor("#92cddc")
        cell.set_text_props(weight="bold", color="black", ha="center")


    out_path = SNAPSHOT_DIR / f"snapshot_{today.strftime('%Y-%m-%d')}.png"
    fig.savefig(out_path, dpi=200, bbox_inches="tight")
    plt.close(fig)
    return out_path


# ---------------------------------------------------------------------------
# Routes
# ---------------------------------------------------------------------------

@app.get("/")
def root():
    return FileResponse(STATIC_DIR / "index.html")


@app.get("/api/health")
def health():
    return {"status": "ok", "urls_configured": URLS_FILE.exists(), "tracker_present": TRACKER_FILE.exists()}


@app.get("/api/config")
def config():
    return {"password_required": bool(APP_PASSWORD), "cookies_saved": COOKIES_FILE.exists()}


@app.get("/api/run/status")
def run_status():
    return JOB_STATE


@app.post("/api/run")
async def run_scrape(request: Request):
    payload = await request.json()

    if APP_PASSWORD and payload.get("password") != APP_PASSWORD:
        return JSONResponse({"status": "error", "message": "Incorrect password."}, status_code=401)

    if JOB_STATE.get("status") == "running":
        return JSONResponse({"status": "error", "message": "A scrape is already running."}, status_code=409)

    cookies_text = (payload.get("cookies_text") or "").strip()
    used_saved_cookies = False

    if cookies_text:
        try:
            raw = json.loads(cookies_text)
        except Exception as e:
            return JSONResponse({"status": "error", "message": f"That doesn't look like valid JSON: {e}"}, status_code=400)

        try:
            cookies = normalize_cookies(raw)
        except Exception as e:
            return JSONResponse({"status": "error", "message": str(e)}, status_code=400)

        if not cookies:
            return JSONResponse({"status": "error", "message": "No bigmint.co cookies found in that JSON."}, status_code=400)
    else:
        cookies = load_saved_cookies()
        if not cookies:
            return JSONResponse({
                "status": "error",
                "message": "No saved cookies yet — paste your bigmint.co cookies JSON once to get started.",
            }, status_code=400)
        used_saved_cookies = True

    if not URLS_FILE.exists():
        return JSONResponse({"status": "error", "message": "Server has no urls.csv configured — see README."}, status_code=500)

    JOB_STATE.clear()
    JOB_STATE.update({"status": "running", "phase": "starting browser", "progress": {"done": 0, "total": 0}})
    asyncio.create_task(_run_scrape_job(cookies, used_saved_cookies))
    return {"status": "started"}


async def _run_scrape_job(cookies, used_saved_cookies):
    try:
        await _do_scrape(cookies, used_saved_cookies)
    except Exception as e:
        JOB_STATE.clear()
        JOB_STATE.update({
            "status": "error",
            "message": f"Unexpected error: {e}",
            "traceback": traceback.format_exc(),
        })


REQUEST_HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/128.0.0.0 Safari/537.36"
    ),
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,application/json;q=0.9,*/*;q=0.8",
    "X-Requested-With": "XMLHttpRequest",
}


async def _do_scrape(cookies, used_saved_cookies):
    urls = read_urls()
    today = datetime.now(IST)
    JOB_STATE["progress"] = {"done": 0, "total": len(urls)}

    results = {}

    async with AsyncSession(
        impersonate="chrome110", headers=REQUEST_HEADERS, timeout=30
    ) as client:
        for c in cookies:
            client.cookies.set(c["name"], c["value"], domain=c["domain"], path=c.get("path", "/"))

        JOB_STATE["phase"] = "checking login"
        try:
            resp = await client.get(BIGMINT_HOME)
        except Exception as e:
            JOB_STATE.clear()
            JOB_STATE.update({"status": "error", "message": f"Could not reach bigmint.co: {e}",
                               "traceback": traceback.format_exc()})
            return

        body = resp.text
        if not text_looks_logged_in(body):
            blocked_by_shield = text_looks_shield_blocked(body)
            if blocked_by_shield:
                message = (
                    "bigmint.co's bot-protection page blocked this request — this isn't your cookies, "
                    "the request got stuck on their security check. Try running again in a minute; if "
                    "it keeps happening, tell me and I'll dig further."
                )
            else:
                if used_saved_cookies and COOKIES_FILE.exists():
                    COOKIES_FILE.unlink()
                message = (
                    "Saved cookies have expired. Export fresh cookies from a logged-in BigMint browser "
                    "session and paste them in — they'll be reused automatically next time."
                    if used_saved_cookies else
                    "These cookies didn't log in — they've likely expired. Export fresh cookies from a "
                    "logged-in BigMint browser session and paste them in again."
                )
            JOB_STATE.clear()
            JOB_STATE.update({
                "status": "cookies_expired",
                "message": message,
                "debug": {
                    "cookie_names_sent": sorted(set(c["name"] for c in cookies)),
                    "landed_url": str(resp.url),
                    "http_status": resp.status_code,
                    "body_snippet": body[:300],
                },
            })
            return

        save_cookies(cookies)

        for i, url in enumerate(urls):
            JOB_STATE["phase"] = f"scraping {slug_from_url(url)}"
            row = {"url": url, "current_price": "", "status": "ok", "notes": ""}
            parsed = parse_item_from_url(url)
            if not parsed:
                row["status"] = "error"
                row["notes"] = "Could not parse item ID/price type/currency from this URL"
            else:
                try:
                    price = await fetch_current_price(client, **parsed)
                    if price is None:
                        row["status"] = "check"
                    else:
                        row["current_price"] = clean_number(str(price))
                except Exception as e:
                    row["status"] = "error"
                    row["notes"] = str(e)[:200]
            results[url] = row
            JOB_STATE["progress"] = {"done": i + 1, "total": len(urls)}

    JOB_STATE["phase"] = "updating tracker"
    tracker_summary = update_tracker(results, today)
    write_csv(results, today)
    snapshot_path = generate_snapshot(today)

    ok_count = sum(1 for r in results.values() if r["status"] == "ok")

    JOB_STATE.clear()
    JOB_STATE.update({
        "status": "done",
        "date": today.strftime("%Y-%m-%d"),
        "scraped": len(results),
        "scraped_ok": ok_count,
        "results": list(results.values()),
        "tracker": tracker_summary,
        "snapshot_available": snapshot_path is not None,
        "downloads": {
            "xlsx": "/api/download/xlsx",
            "csv": "/api/download/csv",
            "png": "/api/download/png",
        },
    })


@app.get("/api/download/xlsx")
def download_xlsx():
    if not TRACKER_FILE.exists():
        return JSONResponse({"status": "error", "message": "No tracker file yet."}, status_code=404)
    return FileResponse(TRACKER_FILE, filename=TRACKER_FILE.name,
                         media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


@app.get("/api/download/csv")
def download_csv():
    if not OUTPUT_CSV.exists():
        return JSONResponse({"status": "error", "message": "No CSV output yet."}, status_code=404)
    return FileResponse(OUTPUT_CSV, filename=OUTPUT_CSV.name, media_type="text/csv")


@app.get("/api/download/png")
def download_png():
    snaps = sorted(SNAPSHOT_DIR.glob("snapshot_*.png"))
    if not snaps:
        return JSONResponse({"status": "error", "message": "No snapshot yet."}, status_code=404)
    latest = snaps[-1]
    return FileResponse(latest, filename=latest.name, media_type="image/png")
