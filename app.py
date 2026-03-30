"""
Flask web app for Appendix A PPTX automation.
"""

import os
import re
import uuid
from pathlib import Path

from flask import (Flask, flash, redirect, render_template, request,
                   send_file, url_for)
from werkzeug.utils import secure_filename

from generator import detect_ga_count, generate_pptx, sort_design_plans

# ─── Config ───────────────────────────────────────────────────────────────────

BASE_DIR      = Path(__file__).parent
TEMPLATE_PPTX = BASE_DIR / "template_A.pptx"
UPLOAD_DIR    = BASE_DIR / "uploads"
OUTPUT_DIR    = BASE_DIR / "outputs"

UPLOAD_DIR.mkdir(exist_ok=True)
OUTPUT_DIR.mkdir(exist_ok=True)

ALLOWED_IMAGE_EXT = {".png", ".jpg", ".jpeg", ".gif", ".bmp", ".tiff", ".tif"}
ALLOWED_PPTX_EXT  = {".pptx"}

app = Flask(__name__)
app.secret_key = "appendix-a-secret"
app.config["MAX_CONTENT_LENGTH"] = 200 * 1024 * 1024   # 200 MB


# ─── Helpers ──────────────────────────────────────────────────────────────────

def save_upload(file_obj, subdir: str, filename: str = None) -> Path:
    target_dir = UPLOAD_DIR / subdir
    target_dir.mkdir(exist_ok=True)
    fname = filename or secure_filename(file_obj.filename)
    path = target_dir / fname
    file_obj.save(str(path))
    return path


def classify_images(paths: list):
    """Split uploaded images into design-plan list and building-image list."""
    design_plans = []
    building_imgs = []
    for p in paths:
        stem = Path(p).stem.lower()
        if "design plan" in stem or "design_plan" in stem:
            design_plans.append(str(p))
        else:
            building_imgs.append(str(p))
    return sort_design_plans(design_plans), building_imgs


# ─── Routes ───────────────────────────────────────────────────────────────────

@app.route("/", methods=["GET"])
def index():
    return render_template("index.html")


@app.route("/detect-ga", methods=["POST"])
def detect_ga():
    """AJAX: upload P01 PPTX, return detected GA slide count."""
    if "p01_file" not in request.files:
        return {"count": 0, "method": "none"}

    file = request.files["p01_file"]
    if not file.filename.lower().endswith(".pptx"):
        return {"count": 0, "method": "invalid"}

    tmp_id = uuid.uuid4().hex
    tmp_path = UPLOAD_DIR / f"detect_{tmp_id}.pptx"
    file.save(str(tmp_path))

    try:
        count = detect_ga_count(str(tmp_path))
        method = "xml" if count > 0 else "ocr"
        return {"count": count, "method": method}
    except Exception as e:
        return {"count": 0, "error": str(e)}
    finally:
        try:
            tmp_path.unlink()
        except Exception:
            pass


@app.route("/generate", methods=["POST"])
def generate():
    session_id = uuid.uuid4().hex

    # ── Validate required fields ──────────────────────────────────────────────
    errors = []

    def require_field(name, label):
        val = request.form.get(name, "").strip()
        if not val:
            errors.append(f"{label} is required.")
        return val

    p01_date        = require_field("p01_date",        "P01 date")
    p02_date        = require_field("p02_date",        "P02 date")
    author          = require_field("author",          "Author")
    checked         = require_field("checked",         "Checked")
    approved        = require_field("approved",        "Approved")
    station_address = require_field("station_address", "Station address")

    ga_override_raw = request.form.get("ga_count_override", "").strip()
    ga_override = int(ga_override_raw) if ga_override_raw.isdigit() else None

    # ── Required files ────────────────────────────────────────────────────────
    image_3d_file = request.files.get("image_3d")
    p01_file      = request.files.get("p01_file")

    if not image_3d_file or not image_3d_file.filename:
        errors.append("3D image is required.")
    elif Path(image_3d_file.filename).suffix.lower() not in ALLOWED_IMAGE_EXT:
        errors.append("3D image must be an image file (PNG, JPG, etc.).")

    if not p01_file or not p01_file.filename:
        errors.append("Input PPTX (P01) is required.")
    elif Path(p01_file.filename).suffix.lower() not in ALLOWED_PPTX_EXT:
        errors.append("Input PPTX must be a .pptx file.")

    image_files = request.files.getlist("images")
    if not image_files or all(not f.filename for f in image_files):
        errors.append("At least one design/building image is required.")

    if errors:
        for e in errors:
            flash(e, "error")
        return redirect(url_for("index"))

    # ── Save uploads ──────────────────────────────────────────────────────────
    subdir = session_id

    image_3d_path = save_upload(
        image_3d_file,
        subdir,
        secure_filename(image_3d_file.filename),
    )
    p01_path = save_upload(
        p01_file,
        subdir,
        "p01_input.pptx",
    )

    # Store (saved_path, original_filename) so we can extract titles from the
    # original name before secure_filename strips characters like '&'.
    img_entries = []
    for f in image_files:
        if f.filename and Path(f.filename).suffix.lower() in ALLOWED_IMAGE_EXT:
            p = save_upload(f, subdir, secure_filename(f.filename))
            img_entries.append((p, f.filename))

    if not img_entries:
        flash("No valid image files were uploaded.", "error")
        return redirect(url_for("index"))

    img_paths = [p for p, _ in img_entries]
    design_plans, building_imgs = classify_images(img_paths)

    # Extract building titles from original filenames (secure_filename strips '&' etc.)
    from generator import extract_building_title
    building_titles = [
        extract_building_title(orig)
        for p, orig in img_entries
        if str(p) in building_imgs
    ]

    # ── Generate ──────────────────────────────────────────────────────────────
    try:
        output_bytes = generate_pptx(
            template_path    = str(TEMPLATE_PPTX),
            input_pptx_path  = str(p01_path),
            image_3d_path    = str(image_3d_path),
            p01_date         = p01_date,
            p02_date         = p02_date,
            author           = author,
            checked          = checked,
            approved         = approved,
            station_address  = station_address,
            design_plan_images = design_plans,
            building_images    = building_imgs,
            building_titles    = building_titles,
            ga_count_override  = ga_override,
        )
    except Exception as exc:
        flash(f"Generation failed: {exc}", "error")
        return redirect(url_for("index"))

    # ── Save & serve output ───────────────────────────────────────────────────
    from generator import extract_station_info
    try:
        _, sname, combined = extract_station_info(str(image_3d_path))
        out_name = f"A – RF Schematics & GA Drawings - {combined}.pptx"
    except Exception:
        out_name = "Appendix_A_output.pptx"

    out_path = OUTPUT_DIR / f"{session_id}.pptx"
    out_path.write_bytes(output_bytes)

    return send_file(
        str(out_path),
        as_attachment=True,
        download_name=out_name,
        mimetype="application/vnd.openxmlformats-officedocument.presentationml.presentation",
    )


# ─── Entry point ──────────────────────────────────────────────────────────────

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5050))
    app.run(host="0.0.0.0", debug=False, port=port)
