#!/usr/bin/env bash
set -e  # exit on any error — makes failures visible

echo ">>> Installing LibreOffice..."
apt-get update -qq
apt-get install -y libreoffice

echo ">>> Verifying LibreOffice installation..."
which soffice && soffice --version || echo "WARNING: soffice not found on PATH"
which libreoffice && libreoffice --version || echo "WARNING: libreoffice not found on PATH"

echo ">>> Installing Python dependencies..."
pip install -r requirements.txt

echo ">>> Build complete."
