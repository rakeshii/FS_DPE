"""
WSGI entry point for gunicorn / Railway deployment.
"""
import os
import sys

# Must be before ANY local imports so that 'routes', 'app', etc. are found
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from app import create_app
# Ensure upload/output temp dirs exist at startup
for d in ['/tmp/fsproj_uploads', '/tmp/fsproj_outputs']:
    os.makedirs(d, exist_ok=True)

application = create_app()
app = application  # gunicorn compatible alias
