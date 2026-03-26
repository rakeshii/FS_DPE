FROM python:3.11-slim

# Install LibreOffice — works in Docker because we have full write access
RUN apt-get update && \
    apt-get install -y --no-install-recommends libreoffice && \
    apt-get clean && \
    rm -rf /var/lib/apt/lists/*

# Set working directory
WORKDIR /app

# Install Python dependencies
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copy application files
COPY . .

# Ensure tmp dirs exist
RUN mkdir -p /tmp/fsproj_uploads /tmp/fsproj_outputs

# Expose port
EXPOSE 10000

# Start gunicorn
CMD gunicorn wsgi:app --bind 0.0.0.0:${PORT:-10000} --timeout 120 --workers 2
