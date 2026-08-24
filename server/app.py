"""
BigMint price tracker — web app backend.

Paste fresh bigmint.co cookies into the frontend textbox, hit Run, and this
scrapes the current price for every URL in data/urls.csv, updates
data/tracker.xlsx (the running two-year tracker), regenerates a snapshot
image, and writes a CSV of the run. If the cookies are missing/expired, it
reports that clearly instead of guessing.

Runs as a normal server process with its own outbound network access (no
sandbox network restrictions), so Playwright talks to bigmint.co directly.
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

from fastapi import FastAPI, Request
from fastapi.responses import JSONResponse, FileResponse
import openpyxl
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt

from playwright.async_api import async_playwright

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


async def is_logged_in(page):
    """The login modal's markup (password field etc.) is present in the DOM
    even when you ARE logged in, so we can't just look for 'password'. The
    reliable signal is the personalised greeting / nav that only render for
    an authenticated session."""
    try:
        text = (await page.locator("body").inner_text(timeout=5000))[:800].lower()
    except Exception:
        text = ""
    return ("hi, mr" in text or "hi, ms" in text or "my portfolio" in text) and "enter your phone number" not in text


TRACKER_HOSTS = (
    "posthog.com", "i.posthog.com",
    "google-analytics.com", "googletagmanager.com", "doubleclick.net",
    "facebook.net", "connect.facebook.net",
    "clarity.ms", "adroll.com", "sendinblue.com", "sibautomation.com",
    "hotjar.com",
)


async def _block_heavy_requests(route):
    """Drop images/media/fonts and known analytics/tracker requests so
    Chromium's memory footprint stays low enough to survive Render's
    free-tier RAM limit — we only need the page's text content, not its
    images or third-party trackers."""
    req = route.request
    if req.resource_type in ("image", "media", "font"):
        await route.abort()
        return
    if any(host in req.url for host in TRACKER_HOSTS):
        await route.abort()
        return
    await route.continue_()


async def wait_past_shield_check(page, max_wait_ms=20000, poll_ms=1000):
    """bigmint.co sits behind a Bunny Shield JS challenge ("Establishing a
    secure connection...") that briefly interstitials real navigations too.
    Poll until the title moves past it instead of assuming one fixed delay
    is enough."""
    waited = 0
    while waited < max_wait_ms:
        try:
            title = (await page.title()).lower()
        except Exception:
            title = ""
        if "establishing" not in title and "secure connection" not in title:
            break
        await page.wait_for_timeout(poll_ms)
        waited += poll_ms
    await page.wait_for_timeout(1500)


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
            if isinstance(val, float):
                return f"{val:,.0f}" if val.is_integer() else f"{val:,.2f}"
            return str(val)

        rows.append([fmt(ws.cell(row=r, column=c).value) for c in display_cols])

    n_rows = len(rows) + 1
    n_cols = len(display_cols)
    fig_width = max(10, n_cols * 1.3 + 2)
    fig_height = max(2, n_rows * 0.4)

    fig, ax = plt.subplots(figsize=(fig_width, fig_height))
    ax.axis("off")
    table = ax.table(cellText=rows, colLabels=headers, loc="center", cellLoc="center")
    table.auto_set_font_size(False)
    table.set_fontsize(9)
    table.scale(1, 1.6)
    table.auto_set_column_width(list(range(n_cols)))

    for row_idx in range(n_rows):
        cell = table[row_idx, 0]
        cell.set_text_props(ha="left")
        cell.PAD = 0.02

    for col_idx in range(n_cols):
        cell = table[0, col_idx]
        cell.set_facecolor("#2c3e50")
        cell.set_text_props(weight="bold", color="white")

    current_display_idx = display_cols.index(current_col)
    for row_idx in range(1, n_rows):
        table[row_idx, current_display_idx].set_facecolor("#eaf2f8")

    ax.set_title(f"BigMint Price Tracker — Snapshot {today.strftime('%Y-%m-%d')}",
                 fontsize=12, fontweight="bold", pad=20)

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


async def _do_scrape(cookies, used_saved_cookies):
    urls = read_urls()
    today = datetime.now(IST)
    JOB_STATE["progress"] = {"done": 0, "total": len(urls)}

    results = {}
    async with async_playwright() as p:
        browser = await p.chromium.launch(
            headless=True,
            args=[
                "--disable-blink-features=AutomationControlled",
                # Render's free tier has very little RAM; these cut Chromium's
                # footprint to reduce the odds of an OOM kill mid-run.
                # (Deliberately NOT --single-process — that flag is known to
                # be unstable for headless Chrome on Linux and can itself
                # cause the crashes it's meant to prevent.)
                "--disable-dev-shm-usage",
                "--disable-gpu",
                "--no-sandbox",
                "--disable-extensions",
            ],
        )
        context = await browser.new_context(
            viewport={"width": 1280, "height": 800},
            user_agent=(
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                "(KHTML, like Gecko) Chrome/128.0.0.0 Safari/537.36"
            ),
            locale="en-US",
        )
        await context.add_init_script(
            "Object.defineProperty(navigator, 'webdriver', { get: () => undefined });"
        )
        await context.route("**/*", _block_heavy_requests)
        await context.add_cookies(cookies)
        page = await context.new_page()

        JOB_STATE["phase"] = "loading bigmint.co"
        try:
            await page.goto(BIGMINT_HOME, wait_until="domcontentloaded", timeout=60000)
            await wait_past_shield_check(page)
        except Exception as e:
            await browser.close()
            JOB_STATE.clear()
            JOB_STATE.update({"status": "error", "message": f"Could not reach bigmint.co: {e}",
                               "traceback": traceback.format_exc()})
            return

        JOB_STATE["phase"] = "checking login"
        if not await is_logged_in(page):
            try:
                debug_url = page.url
                debug_title = await page.title()
                debug_snippet = (await page.locator("body").inner_text(timeout=3000))[:300]
            except Exception:
                debug_url, debug_title, debug_snippet = "", "", ""
            await browser.close()

            blocked_by_shield = "establishing" in debug_title.lower() or "secure connection" in debug_title.lower()
            if blocked_by_shield:
                message = (
                    "bigmint.co's bot-protection page didn't clear in time — this isn't your cookies, "
                    "the automated browser got stuck on their security check. Try running again in a "
                    "minute; if it keeps happening, tell me and I'll dig further."
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
                    "landed_url": debug_url,
                    "page_title": debug_title,
                    "body_snippet": debug_snippet,
                },
            })
            return

        save_cookies(cookies)

        for i, url in enumerate(urls):
            JOB_STATE["phase"] = f"scraping {slug_from_url(url)}"
            row = {"url": url, "current_price": "", "status": "ok", "notes": ""}
            try:
                await page.goto(url, wait_until="domcontentloaded", timeout=45000)
                await wait_past_shield_check(page, max_wait_ms=10000)
                await page.wait_for_timeout(1800)
                body_text = await page.locator("body").inner_text()
                m = re.search(r"₹\s*([0-9,]+(?:\.[0-9]+)?)", body_text)
                row["current_price"] = clean_number(m.group(1)) if m else ""
                if not row["current_price"]:
                    row["status"] = "check"
            except Exception as e:
                row["status"] = "error"
                row["notes"] = str(e)[:200]
            results[url] = row
            JOB_STATE["progress"] = {"done": i + 1, "total": len(urls)}

        await browser.close()

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
