"""
Deck Converter — Browser-based tool.
Run: python app.py
Opens at http://localhost:5000
"""

import os
import json
import webbrowser
import threading
from flask import Flask, request, jsonify, send_file, send_from_directory
from werkzeug.utils import secure_filename

from detector import analyze_deck, SLIDE_TYPES, SLIDE_TYPE_LABELS, SLIDE_TYPE_DESCRIPTIONS
from mapper import map_slide
import template_slick
import template_colorful
import subprocess
import base64
import glob

app = Flask(__name__, static_folder='static')
app.config['UPLOAD_FOLDER'] = os.path.join(os.path.dirname(__file__), 'uploads')
app.config['OUTPUT_FOLDER'] = os.path.join(os.path.dirname(__file__), 'output')
app.config['THUMB_FOLDER'] = os.path.join(os.path.dirname(__file__), 'thumbs')
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024  # 50MB

os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
os.makedirs(app.config['OUTPUT_FOLDER'], exist_ok=True)
os.makedirs(app.config['THUMB_FOLDER'], exist_ok=True)

# Store analysis results in memory (single-user tool)
_current_analysis = None
_current_file = None
_thumbnails_available = False


def _check_libreoffice():
    """Check if LibreOffice is available."""
    for cmd in ['soffice', '/Applications/LibreOffice.app/Contents/MacOS/soffice',
                'libreoffice', '/usr/bin/libreoffice']:
        try:
            subprocess.run([cmd, '--version'], capture_output=True, timeout=5)
            return cmd
        except (FileNotFoundError, subprocess.TimeoutExpired):
            continue
    return None


def _generate_thumbnails(pptx_path):
    """Generate slide thumbnail images using LibreOffice + pdftoppm."""
    soffice = _check_libreoffice()
    if not soffice:
        return False

    thumb_dir = app.config['THUMB_FOLDER']
    # Clear old thumbnails
    for f in glob.glob(os.path.join(thumb_dir, 'slide-*')):
        os.remove(f)

    try:
        # Convert to PDF
        pdf_path = os.path.join(thumb_dir, 'slides.pdf')
        subprocess.run([
            soffice, '--headless', '--convert-to', 'pdf',
            '--outdir', thumb_dir, pptx_path
        ], capture_output=True, timeout=30)

        # Find the generated PDF (filename may vary)
        pdfs = glob.glob(os.path.join(thumb_dir, '*.pdf'))
        if not pdfs:
            return False
        pdf_path = pdfs[0]

        # Convert PDF pages to JPEG thumbnails
        subprocess.run([
            'pdftoppm', '-jpeg', '-r', '120',
            pdf_path, os.path.join(thumb_dir, 'slide')
        ], capture_output=True, timeout=30)

        # Check if thumbnails were created
        thumbs = sorted(glob.glob(os.path.join(thumb_dir, 'slide-*.jpg')))
        return len(thumbs) > 0

    except (subprocess.TimeoutExpired, FileNotFoundError):
        return False


def _get_thumbnail_b64(slide_num):
    """Get base64-encoded thumbnail for a slide number."""
    thumb_dir = app.config['THUMB_FOLDER']
    # pdftoppm names files like slide-01.jpg, slide-02.jpg, etc.
    patterns = [
        os.path.join(thumb_dir, f'slide-{slide_num:02d}.jpg'),
        os.path.join(thumb_dir, f'slide-{slide_num:01d}.jpg'),
        os.path.join(thumb_dir, f'slide-{slide_num:03d}.jpg'),
    ]
    for path in patterns:
        if os.path.exists(path):
            with open(path, 'rb') as f:
                return base64.b64encode(f.read()).decode('utf-8')
    return None


def _build_text_structure(slide_data):
    """Build a structured text breakdown of a slide's content."""
    boxes = slide_data.get("raw_boxes", [])
    if not boxes:
        return []

    structure = []
    for i, box in enumerate(boxes):
        font_size = box.get("max_font_size", 0)
        size_label = ""
        if font_size >= 36:
            size_label = "Title"
        elif font_size >= 24:
            size_label = "Heading"
        elif font_size >= 14:
            size_label = "Body"
        elif font_size > 0:
            size_label = "Small"
        else:
            size_label = "Text"

        paras = box.get("paragraphs", [])
        structure.append({
            "box_index": i + 1,
            "role": size_label,
            "font_size": font_size,
            "lines": paras[:8],  # Cap at 8 lines per box
            "truncated": len(paras) > 8,
        })

    return structure


@app.route('/')
def index():
    return send_from_directory('static', 'index.html')


@app.route('/api/upload', methods=['POST'])
def upload():
    global _current_analysis, _current_file

    if 'file' not in request.files:
        return jsonify({"error": "No file uploaded"}), 400

    f = request.files['file']
    if not f.filename.endswith('.pptx'):
        return jsonify({"error": "Please upload a .pptx file"}), 400

    filename = secure_filename(f.filename)
    filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
    f.save(filepath)

    try:
        analysis = analyze_deck(filepath)
        _current_analysis = analysis
        _current_file = filepath

        # Try to generate thumbnails
        has_thumbs = _generate_thumbnails(filepath)

        slides_out = []
        for s in analysis:
            sr = {
                "number": s["number"],
                "detected_type": s["detected_type"],
                "confidence": round(s["confidence"], 2),
                "reason": s["reason"],
                "candidates": s.get("candidates", []),
                "preview": s["preview"],
                "total_words": s["total_words"],
                "has_chart": s["has_chart"],
                "has_image": s["has_image"],
                "text_structure": _build_text_structure(s),
            }
            if has_thumbs:
                thumb = _get_thumbnail_b64(s["number"])
                if thumb:
                    sr["thumbnail"] = thumb
            slides_out.append(sr)

        return jsonify({
            "filename": filename,
            "slide_count": len(analysis),
            "has_thumbnails": has_thumbs,
            "slides": slides_out,
            "available_types": [{
                "value": t,
                "label": SLIDE_TYPE_LABELS.get(t, t),
                "description": SLIDE_TYPE_DESCRIPTIONS.get(t, ""),
            } for t in SLIDE_TYPES],
        })
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route('/api/build', methods=['POST'])
def build():
    global _current_analysis, _current_file

    if not _current_analysis or not _current_file:
        return jsonify({"error": "No file analyzed yet. Upload a .pptx first."}), 400

    data = request.json
    template = data.get("template", "slick")
    overrides = data.get("overrides", {})  # {"1": "section_divider", "3": "in_brief", ...}

    # Build slide configs
    slide_configs = []
    for slide_data in _current_analysis:
        num = str(slide_data["number"])
        slide_type = overrides.get(num, slide_data["detected_type"])

        if slide_type == "skip":
            continue

        mapped = map_slide(slide_data, slide_type)
        slide_configs.append((slide_type, mapped))

    # Build the deck
    output_name = os.path.splitext(os.path.basename(_current_file))[0]
    output_name = f"{output_name}_{template}.pptx"
    output_path = os.path.join(app.config['OUTPUT_FOLDER'], output_name)

    try:
        if template == "colorful":
            template_colorful.build_deck(slide_configs, output_path)
        else:
            template_slick.build_deck(slide_configs, output_path)

        return jsonify({"success": True, "filename": output_name})
    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({"error": str(e)}), 500


@app.route('/api/download/<filename>')
def download(filename):
    filepath = os.path.join(app.config['OUTPUT_FOLDER'], secure_filename(filename))
    if os.path.exists(filepath):
        return send_file(filepath, as_attachment=True)
    return jsonify({"error": "File not found"}), 404


def open_browser():
    """Open browser after a short delay to let Flask start."""
    import time
    time.sleep(1.5)
    webbrowser.open('http://localhost:5000')


if __name__ == '__main__':
    print("\n  Deck Converter")
    print("  ─────────────────────────────")
    print("  Opening http://localhost:5000")
    print("  Press Ctrl+C to stop\n")

    threading.Thread(target=open_browser, daemon=True).start()
    app.run(host='0.0.0.0', port=5000, debug=False)
