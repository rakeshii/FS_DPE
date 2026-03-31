"""
core/validator.py
Basic sanity checks on uploaded files before we run the heavy pipeline.
Returns an error string on failure, or None on success.
"""

import os


# Minimum expected sheet names — at least these must be present
REQUIRED_SHEETS = {'BS', 'P & L', 'CFS'}

# Max file size in bytes (20 MB)
MAX_SIZE = 20 * 1024 * 1024


def validate_upload(file_path: str) -> str | None:
    """
    Run quick checks on the uploaded file.
    Returns an error message string if invalid, None if OK.
    """
    # ── 1. File exists ─────────────────────────────────────────
    if not os.path.exists(file_path):
        return 'Uploaded file not found on server.'

    # ── 2. File size ───────────────────────────────────────────
    size = os.path.getsize(file_path)
    if size == 0:
        return 'Uploaded file is empty.'
    if size > MAX_SIZE:
        return f'File too large ({size // (1024*1024)} MB). Maximum is 20 MB.'

    # ── 3. Try reading it with openpyxl / xlrd ─────────────────
    ext = os.path.splitext(file_path)[1].lower()

    if ext == '.xlsx':
        try:
            import openpyxl
            wb = openpyxl.load_workbook(file_path, read_only=True, keep_links=False)
            sheet_names = set(wb.sheetnames)
            wb.close()
        except Exception as e:
            return f'Could not read .xlsx file: {e}'
    elif ext == '.xls':
        try:
            import xlrd
            wb = xlrd.open_workbook(file_path)
            sheet_names = set(wb.sheet_names())
        except Exception as e:
            return f'Could not read .xls file: {e}'
    else:
        return 'Unsupported file type. Please upload .xls or .xlsx.'

    # ── 4. Check required sheets ───────────────────────────────
    missing = REQUIRED_SHEETS - sheet_names
    if missing:
        return (
            f'File is missing expected sheets: {", ".join(sorted(missing))}. '
            f'Please check you are uploading the correct financial statements file.'
        )

    return None   # All checks passed
