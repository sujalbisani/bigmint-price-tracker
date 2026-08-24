# BigMint Price Tracker — Web App

A small web app version of the BigMint scraper: open a URL from any browser,
paste fresh cookies into a textbox, click Run. It scrapes today's prices for
the 24 tracked items, updates the tracker spreadsheet, generates the snapshot
image, and gives you download links. If the cookies have expired, it tells
you plainly instead of failing silently.

This runs as a real server with its own internet access — unlike a sandboxed
assistant session, it isn't blocked from reaching bigmint.co.

## What's in this folder

```
server/
  app.py            — FastAPI backend: scrape, update xlsx, generate snapshot
  static/index.html — the single-page frontend (cookie textbox + Run button)
  requirements.txt  — Python dependencies
  Dockerfile         — builds a container with Chromium + Python preinstalled
  data/
    urls.csv         — the 24 BigMint URLs being tracked (edit to add/remove items)
    tracker.xlsx      — the running two-year tracker (seeded with your latest data)
render.yaml           — one-file Render deployment config (optional but convenient)
```

## Deploying for free (Render)

Render's free web service tier costs nothing and does **not** require a
credit card. The tradeoff: the app "sleeps" after ~15 minutes with no
traffic and takes 30–60 seconds to wake up on the next request. For a tool
you run once a day, that's a fine tradeoff for $0.

**Steps:**

1. **Put this folder in a GitHub repo.**
   - Create a new repo on github.com (can be private).
   - Upload this whole `bigmint_webapp` folder to it (drag-and-drop on
     github.com works, or `git init && git add . && git commit -m "init" && git push`
     if you're comfortable with git).

2. **Create the Render service.**
   - Go to [render.com](https://render.com) and sign up (no card needed for
     the free tier).
   - Click **New +** → **Web Service**.
   - Connect your GitHub account and pick the repo you just created.
   - Render should auto-detect the `render.yaml` in the repo root and
     pre-fill the settings (Docker, free plan, pointed at `server/Dockerfile`).
     If it doesn't auto-detect, set these manually:
     - **Runtime:** Docker
     - **Dockerfile Path:** `server/Dockerfile`
     - **Docker Build Context Directory:** `server`
     - **Plan:** Free

3. **(Optional) Set a password.**
   - In the service's **Environment** tab, add an environment variable
     `APP_PASSWORD` with any value you choose. The web page will then ask
     for that password before running a scrape — useful since this touches
     your BigMint login cookies and internal pricing data.
   - Skip this if you don't need it; the app works fine without a password,
     it just means anyone with the URL can trigger a run.

4. **Deploy.**
   - Click **Create Web Service**. The first build takes a few minutes
     (it's installing Chromium). Render gives you a URL like
     `https://bigmint-price-tracker.onrender.com` — that's your web app.

5. **Use it.**
   - Open the URL, paste fresh bigmint.co cookies (see below), click
     **Run today's scrape**.
   - It'll show you the results and give you download links for the
     updated tracker `.xlsx`, a `.csv`, and the snapshot `.png`.

## Getting fresh cookies

When the app tells you cookies have expired:

1. Log into bigmint.co in your regular browser.
2. Use a cookie-export browser extension (e.g. "Cookie-Editor" for Chrome/Firefox) to export cookies for the bigmint.co domain as JSON.
3. Paste that JSON into the textbox on the web app and click Run.

The app only accepts `.bigmint.co` cookies — anything else in the export is
ignored.

## Important limitation: storage isn't guaranteed permanent

Render's free tier disk is tied to the running instance. It generally
survives the app sleeping and waking back up, but it is **not** a permanent
database — a redeploy (e.g. pushing new code) resets it back to whatever is
committed in the repo (i.e. back to `data/tracker.xlsx` as it exists in
GitHub). **Always download the updated `.xlsx` after a run** if you want to
keep it — don't rely on the server as your only copy. If you want the
tracker to always start from your latest data after a redeploy, replace
`server/data/tracker.xlsx` in the GitHub repo with your latest downloaded
copy before pushing changes.

## Editing the tracked items

To add or remove BigMint URLs, edit `server/data/urls.csv` (one URL per
line, header row `url`) and push the change to GitHub — Render will
redeploy automatically.

## Running locally (optional, to test before deploying)

If you have Docker installed on your own machine:

```bash
cd server
docker build -t bigmint-tracker .
docker run -p 8000:8000 bigmint-tracker
```

Then open http://localhost:8000 in your browser.
