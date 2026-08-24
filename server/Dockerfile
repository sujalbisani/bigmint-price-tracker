FROM python:3.11-slim

WORKDIR /app

# System deps Playwright's Chromium needs
RUN apt-get update && apt-get install -y --no-install-recommends \
    wget gnupg ca-certificates \
    && rm -rf /var/lib/apt/lists/*

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Installs Chromium + its OS-level dependencies in one shot
RUN playwright install --with-deps chromium

COPY . .

# Make sure the data dirs exist and survive container restarts as far as the
# platform's disk allows (see README for the persistence caveat)
RUN mkdir -p /app/data/backups /app/data/snapshots

ENV PORT=8000
EXPOSE 8000

CMD ["sh", "-c", "uvicorn app:app --host 0.0.0.0 --port ${PORT}"]
