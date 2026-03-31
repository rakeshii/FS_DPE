# ── Base: official Python 3.11 slim (pip works out of the box) ──────────────
FROM python:3.11-slim

# ── System dependencies ──────────────────────────────────────────────────────
# libreoffice-calc + writer for formula recalculation step
RUN apt-get update \
 && apt-get install -y --no-install-recommends \
        libreoffice-calc \
        libreoffice-writer \
        libreoffice-headless \
        default-jre-headless \
 && apt-get clean \
 && rm -rf /var/lib/apt/lists/*

# ── Working directory ────────────────────────────────────────────────────────
WORKDIR /app

# ── Python dependencies (separate layer for Docker cache efficiency) ─────────
COPY requirements.txt .
RUN pip install --no-cache-dir --upgrade pip \
 && pip install --no-cache-dir -r requirements.txt

# ── Application code ─────────────────────────────────────────────────────────
COPY . .

# ── Runtime temp directories ─────────────────────────────────────────────────
RUN mkdir -p /tmp/fsproj_uploads /tmp/fsproj_outputs

# ── Port (Railway injects $PORT at runtime) ──────────────────────────────────
EXPOSE 8080

# ── Start ────────────────────────────────────────────────────────────────────
CMD gunicorn wsgi:app \
        --bind 0.0.0.0:${PORT:-8080} \
        --timeout 120 \
        --workers 2
