"""
routes/api.py
All API endpoints. Clean separation from business logic.
"""

import os
import uuid
import tempfile

from flask import Blueprint, request, jsonify, send_file, current_app
from werkzeug.utils import secure_filename

from core.projector import generate_projection
from core.validator import validate_upload

api_bp = Blueprint('api', __name__)


def allowed_file(filename):
    ext = filename.rsplit('.', 1)[-1].lower() if '.' in filename else ''
    return ext in current_app.config['ALLOWED_EXTENSIONS']


@api_bp.route('/generate', methods=['POST'])
def generate():
    """
    POST /api/generate
    Form data:
      file         — .xls or .xlsx upload
      new_header   — text for the 2026 column header
      output_name  — desired filename for download
      col_content  — 'copy2025' | 'blank'
      title_update — 'yes' | 'no'
    Returns:
      .xlsx file download on success
      JSON { error: str } on failure
    """
    # ── Validate file ──────────────────────────────────────────
    if 'file' not in request.files:
        return jsonify({'error': 'No file uploaded.'}), 400

    file = request.files['file']

    if not file or not file.filename:
        return jsonify({'error': 'No file selected.'}), 400

    if not allowed_file(file.filename):
        return jsonify({'error': 'Only .xls and .xlsx files are supported.'}), 400

    # ── Read form params ───────────────────────────────────────
    cfg          = current_app.config
    new_header   = request.form.get('new_header',   cfg['DEFAULT_NEW_HEADER']).strip()
    output_name  = request.form.get('output_name',  cfg['DEFAULT_OUTPUT_NAME']).strip()
    col_content  = request.form.get('col_content',  'copy2025')
    title_update = request.form.get('title_update', 'yes')

    if not output_name.endswith('.xlsx'):
        output_name += '.xlsx'

    # ── Save upload to temp dir ────────────────────────────────
    ext        = '.' + secure_filename(file.filename).rsplit('.', 1)[-1].lower()
    uid        = str(uuid.uuid4())[:10]
    tmpdir     = tempfile.mkdtemp()
    input_path = os.path.join(tmpdir, uid + ext)
    out_path   = os.path.join(tmpdir, uid + '_output.xlsx')

    try:
        file.save(input_path)

        # ── Validate content ───────────────────────────────────
        error = validate_upload(input_path)
        if error:
            return jsonify({'error': error}), 422

        # ── Run the projector ──────────────────────────────────
        generate_projection(
            input_path   = input_path,
            output_path  = out_path,
            sheet_config = cfg['SHEET_CONFIG'],
            chain_patches= cfg['CHAIN_PATCHES'],
            new_header   = new_header,
            copy_vals    = (col_content == 'copy2025'),
            update_titles= (title_update == 'yes'),
            lo_cmd       = cfg['LIBREOFFICE_CMD'],
            lo_timeout   = cfg['LIBREOFFICE_TIMEOUT'],
            named_ranges = cfg.get('TB_NAMED_RANGES'),
        )

        return send_file(
            out_path,
            as_attachment=True,
            download_name=output_name,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )

    except RuntimeError as e:
        # Catches LibreOffice not found, conversion failures, etc.
        msg = str(e)
        status = 503 if 'LibreOffice' in msg else 500
        return jsonify({'error': msg}), status

    except Exception as e:
        current_app.logger.exception('Unexpected error in /api/generate')
        return jsonify({'error': f'Unexpected error: {str(e)}'}), 500

    finally:
        try:
            if os.path.exists(input_path):
                os.remove(input_path)
        except OSError:
            pass


@api_bp.route('/check', methods=['GET'])
def check():
    """
    GET /api/check
    Checks whether LibreOffice is installed and returns its path.
    Called on page load so the user knows immediately if setup is complete.
    """
    from core.projector import _find_libreoffice
    try:
        lo_path = _find_libreoffice()
        return jsonify({'libreoffice': True, 'path': lo_path})
    except RuntimeError as e:
        return jsonify({'libreoffice': False, 'error': str(e)}), 200


@api_bp.route('/health', methods=['GET'])
def health():
    """GET /api/health — quick liveness check"""
    return jsonify({'status': 'ok', 'version': 'v8'})
