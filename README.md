# FinStatement Projector — Developer Version
### NSVR & Associates LLP · v8

A production-grade Flask application that adds a **March 2026 column** to
Indian financial statements, preserving all 1,230+ Excel formulas with
correct column reference shifting.

---

## Project Structure

```
finprojector_dev/
│
├── app.py                  ← App factory (create_app)
├── config.py               ← All config: sheet layout, defaults, paths
│
├── routes/
│   ├── main.py             ← Serves the HTML page  (GET /)
│   └── api.py              ← REST endpoints        (POST /api/generate)
│
├── core/
│   ├── projector.py        ← Pure business logic (no Flask)
│   └── validator.py        ← Upload sanity checks
│
├── templates/
│   └── index.html          ← Frontend UI (NSVR & Associates LLP)
│
├── tests/
│   └── test_projector.py   ← 17 unit tests (pytest)
│
├── uploads/                ← Temp upload dir (git-ignored)
├── outputs/                ← Temp output dir (git-ignored)
│
├── requirements.txt
├── Procfile                ← Railway / Heroku
├── nixpacks.toml           ← LibreOffice install on Railway
├── .env.example
└── .gitignore
```

---

## Quick Start (VSCode / Local)

### 1. Prerequisites
```bash
# Python 3.9+
python --version

# LibreOffice (needed to convert .xls → .xlsx)
# Windows: https://www.libreoffice.org/download
# Linux:
sudo apt install libreoffice
# Mac:
brew install --cask libreoffice
```

### 2. Setup
```bash
git clone <your-repo>
cd finprojector_dev

python -m venv venv
source venv/bin/activate        # Windows: venv\Scripts\activate

pip install -r requirements.txt

cp .env.example .env            # Edit SECRET_KEY
```

### 3. Run
```bash
python app.py
# → http://localhost:5000
```

### 4. Run Tests
```bash
python -m pytest tests/ -v
# → 17 passed
```

---

## API Reference

### `POST /api/generate`
Accepts multipart form data, returns `.xlsx` file.

| Field | Type | Default | Description |
|---|---|---|---|
| `file` | File | required | `.xls` or `.xlsx` upload |
| `new_header` | string | `As at 31 March, 2026` | 2026 column header text |
| `output_name` | string | `FS-FY_2025-26_Draft.xlsx` | Download filename |
| `col_content` | string | `copy2025` | `copy2025` or `blank` |
| `title_update` | string | `yes` | `yes` or `no` |

**Success:** Returns `.xlsx` binary (HTTP 200)
**Error:** Returns `{ "error": "..." }` (HTTP 400/422/500)

### `GET /api/health`
Returns `{ "status": "ok", "version": "v8" }`

---

## Deploy to Railway (Free)

1. Push this folder to a GitHub repo
2. Go to [railway.app](https://railway.app) → New Project → GitHub repo
3. Railway reads `nixpacks.toml` → installs LibreOffice automatically
4. Reads `Procfile` → starts gunicorn
5. Set `SECRET_KEY` in Railway environment variables
6. Done — live URL in ~2 minutes

## Deploy to Render.com (Free)

1. Push to GitHub
2. New Web Service → connect repo
3. Build: `pip install -r requirements.txt`
4. Add package `libreoffice` in Environment
5. Start: `gunicorn app:create_app() --bind 0.0.0.0:$PORT --timeout 120`

## Deploy with Docker

```dockerfile
FROM python:3.11-slim
RUN apt-get update && apt-get install -y libreoffice --no-install-recommends && rm -rf /var/lib/apt/lists/*
WORKDIR /app
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt
COPY . .
EXPOSE 8080
CMD ["gunicorn", "app:create_app()", "--bind", "0.0.0.0:8080", "--timeout", "120"]
```

```bash
docker build -t finprojector .
docker run -p 8080:8080 -e SECRET_KEY=mysecret finprojector
```

---

## How It Works

The core problem: `.xls` files store formula **results** (numbers), not formula
strings. To recover the actual formulas (e.g. `=SUM(D9:D12)`), we must convert
`.xls → .xlsx` using **LibreOffice** — this is the critical step that makes
everything else possible.

### Pipeline (`core/projector.py`)

```
1. _ensure_xlsx()         LibreOffice: .xls → .xlsx (recovers 1,230 formulas)
2. _process_financial()   For each of 11 financial sheets:
   a. Save merged cell ranges
   b. Unmerge all
   c. Shift cells right-to-left from insert column
   d. shift_formula() updates every ref: D→E, SUM(D9:D12)→SUM(E9:E12)
   e. Re-merge with shifted column numbers
   f. Write 2026 column (header + formulas + optional values)
3. _process_support()     Fix cross-sheet refs in WDV, RATIO WORKING, TB, COMPUTATION
4. _apply_chain_patches() Fix 8 merged header range formulas
5. wb.save()
```

### Key Formula Shifter Fix (v8)

The regex handles `Sheet!A1:B2` ranges as a single unit — both `A1` and `B2`
belong to `Sheet`, so neither leaks into same-sheet shifting. This fixed the
bug where `TB!C102:C106` became `TB!C102:D106`.

---

## Adding Support for New Sheets

Edit `config.py`:
```python
SHEET_CONFIG = {
    ...
    'My New Sheet': {'insert': 4},   # 1-indexed column to insert 2026 col
}
```

That's it — the projector handles the rest automatically.
