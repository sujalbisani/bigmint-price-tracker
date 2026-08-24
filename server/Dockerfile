FROM python:3.11-slim-bookworm

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

# Make sure the data dirs exist and survive container restarts as far as the
# platform's disk allows (see README for the persistence caveat)
RUN mkdir -p /app/data/backups /app/data/snapshots

ENV PORT=8000
EXPOSE 8000

CMD ["sh", "-c", "uvicorn app:app --host 0.0.0.0 --port ${PORT}"]
