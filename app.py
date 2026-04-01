"""
Flask web app for Appendix A PPTX automation – multi-provider edition.
"""

import io
import os
import re
import uuid
import zipfile as zipfile_mod
from pathlib import Path

from flask import (Flask, flash, redirect, render_template, request,
                   send_file, url_for)
from werkzeug.utils import secure_filename

from generator import (
    extract_building_title,
    extract_station_info,
    generate_pptx,
    sort_design_plans,
)
from generator_b import generate_pptx_b, PROVIDER_CODES as PROVIDER_CODES_B
from engine_b import process_pptx as extract_pptx_images

# ─── Config ───────────────────────────────────────────────────────────────────

BASE_DIR   = Path(__file__).parent
UPLOAD_DIR = BASE_DIR / "uploads"
OUTPUT_DIR = BASE_DIR / "outputs"

UPLOAD_DIR.mkdir(exist_ok=True)
OUTPUT_DIR.mkdir(exist_ok=True)

ALLOWED_IMAGE_EXT = {".png", ".jpg", ".jpeg", ".gif", ".bmp", ".tiff", ".tif"}

# Provider → template file + display name  (Appendix A)
PROVIDERS = {
    "EE":       {"template": BASE_DIR / "template_A_EE.pptx",      "label": "EE"},
    "THREE":    {"template": BASE_DIR / "Template_A_THREE.pptx",    "label": "THREE"},
    "VODAFONE": {"template": BASE_DIR / "Template_A_VODA.pptx",     "label": "VODAFONE"},
    "VMO2":     {"template": BASE_DIR / "Template_A_VMO2.pptx",     "label": "VMO2"},
}

# Appendix B uses ONE shared template for all providers
TEMPLATE_B = BASE_DIR / "templateB.pptx"

# Image filename prefix → provider key
PREFIX_MAP = {
    "EE_":       "EE",
    "THREE_":    "THREE",
    "VODAFONE_": "VODAFONE",   # matches VODAFONE_ / vodafone_ / Vodafone_ etc.
    "VODA_":     "VODAFONE",   # matches VODA_ / voda_ / Voda_ etc.
    "VMO2_":     "VMO2",
}

app = Flask(__name__)
app.secret_key = "appendix-a-secret"
app.config["MAX_CONTENT_LENGTH"] = 500 * 1024 * 1024   # 500 MB


# ─── Helpers ──────────────────────────────────────────────────────────────────

def save_upload(file_obj, subdir: str, filename: str) -> Path:
    target_dir = UPLOAD_DIR / subdir
    target_dir.mkdir(exist_ok=True)
    path = target_dir / filename
    file_obj.save(str(path))
    return path


def provider_for(original_name: str):
    """
    Return provider key for a filename, or None if unrecognised.

    Tries three patterns in order:
      1. PROVIDER_ prefix  (e.g. "EE_...", "THREE_...", "VODA_...")
      2. PROVIDER - prefix (e.g. "EE - ...", "Three - ...", "Vodafone - ...")
      3. Keyword anywhere  (e.g. "Lambeth North_VMO2NEW", "Lambeth North_voda1")
    """
    upper = original_name.upper()
    stem  = Path(original_name).stem.upper()

    # 1. Underscore-separated prefix  (existing behaviour)
    for prefix, key in PREFIX_MAP.items():
        if upper.startswith(prefix.upper()):
            return key

    # 2. "PROVIDER - " space-dash-space prefix
    space_dash = {
        "EE - ":       "EE",
        "THREE - ":    "THREE",
        "VODAFONE - ": "VODAFONE",
        "VODA - ":     "VODAFONE",
    }
    for prefix, key in space_dash.items():
        if upper.startswith(prefix.upper()):
            return key

    # 3. Keyword anywhere in the stem (handles embedded names like _VMO2NEW, _voda1)
    #    Order matters: check longer/more-specific keywords first.
    keyword_map = [
        ("VODAFONE", "VODAFONE"),
        ("VMO2",     "VMO2"),
        ("THREE",    "THREE"),
        ("VODA",     "VODAFONE"),
    ]
    for keyword, key in keyword_map:
        if keyword in stem:
            return key

    # EE is last — it's only 2 chars so check it carefully:
    # must appear as a standalone word (start of stem, or after space/dash/underscore)
    if re.search(r"(^|[\s\-_])EE([\s\-_]|$)", stem):
        return "EE"

    return None


def split_images_by_provider(img_entries: list) -> dict:
    """
    img_entries: list of (saved_path, original_filename)
    Returns: {provider_key: {"design_plans": [...], "building_imgs": [...], "building_titles": [...]}}
    """
    buckets = {
        key: {"design_plans": [], "building_imgs": [], "building_titles": []}
        for key in PROVIDERS
    }

    for saved_path, orig_name in img_entries:
        prov = provider_for(orig_name)
        if prov is None:
            continue   # skip unrecognised prefixes
        stem_lower = Path(orig_name).stem.lower()
        if "design plan" in stem_lower or "design_plan" in stem_lower:
            buckets[prov]["design_plans"].append(str(saved_path))
        else:
            buckets[prov]["building_imgs"].append(str(saved_path))
            buckets[prov]["building_titles"].append(extract_building_title(orig_name))

    # Sort design plans numerically per provider
    for prov in buckets:
        buckets[prov]["design_plans"] = sort_design_plans(buckets[prov]["design_plans"])

    return buckets



# ─── Routes ───────────────────────────────────────────────────────────────────

@app.route("/", methods=["GET"])
def home():
    return render_template("home.html")


@app.route("/appendix-a", methods=["GET"])
def appendix_a_page():
    return render_template("appendix_a.html")


# Legacy redirect so any bookmarks to "/" still work
@app.route("/index", methods=["GET"])
def index():
    return redirect(url_for("appendix_a_page"))


@app.route("/appendix-a/generate", methods=["POST"])
@app.route("/generate", methods=["POST"])   # ← keep old URL working
def generate():
    session_id = uuid.uuid4().hex
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

    # ── 3D image ──────────────────────────────────────────────────────────────
    image_3d_file = request.files.get("image_3d")
    if not image_3d_file or not image_3d_file.filename:
        errors.append("3D image is required.")
    elif Path(image_3d_file.filename).suffix.lower() not in ALLOWED_IMAGE_EXT:
        errors.append("3D image must be an image file (PNG, JPG, etc.).")

    # ── Images ────────────────────────────────────────────────────────────────
    image_files = request.files.getlist("images")
    if not image_files or all(not f.filename for f in image_files):
        errors.append("At least one image is required.")

    if errors:
        for e in errors:
            flash(e, "error")
        return redirect(url_for("appendix_a_page"))

    # ── Save uploads ──────────────────────────────────────────────────────────
    subdir = session_id

    image_3d_path = save_upload(
        image_3d_file,
        subdir,
        secure_filename(image_3d_file.filename),
    )

    # Store (saved_path, original_name) to preserve '&' etc. in titles
    img_entries = []
    for f in image_files:
        if f.filename and Path(f.filename).suffix.lower() in ALLOWED_IMAGE_EXT:
            saved = save_upload(f, subdir, secure_filename(f.filename))
            img_entries.append((saved, f.filename))

    if not img_entries:
        flash("No valid image files were uploaded.", "error")
        return redirect(url_for("appendix_a_page"))

    buckets = split_images_by_provider(img_entries)

    # Validate at least one provider has images
    active_providers = [k for k, v in buckets.items()
                        if v["design_plans"] or v["building_imgs"]]
    if not active_providers:
        flash(
            "No images matched a provider prefix (EE_, THREE_, VODA_, VMO2_). "
            "Please rename your images and try again.",
            "error",
        )
        return redirect(url_for("appendix_a_page"))

    # ── Generate one PPTX per active provider, zip them ──────────────────────
    try:
        _, _, combined = extract_station_info(str(image_3d_path))
    except Exception as exc:
        flash(f"Could not parse station from 3D image filename: {exc}", "error")
        return redirect(url_for("appendix_a_page"))

    zip_buf = io.BytesIO()
    generated = []

    with zipfile_mod.ZipFile(zip_buf, "w", zipfile_mod.ZIP_DEFLATED) as zf:
        for prov in ["EE", "THREE", "VODAFONE", "VMO2"]:
            if prov not in active_providers:
                continue
            info  = PROVIDERS[prov]
            bkt   = buckets[prov]
            try:
                pptx_bytes = generate_pptx(
                    template_path      = str(info["template"]),
                    image_3d_path      = str(image_3d_path),
                    p01_date           = p01_date,
                    p02_date           = p02_date,
                    author             = author,
                    checked            = checked,
                    approved           = approved,
                    station_address    = station_address,
                    design_plan_images = bkt["design_plans"],
                    building_images    = bkt["building_imgs"],
                    building_titles    = bkt["building_titles"],
                )
                fname = f"A – RF Schematics & GA Drawings - {prov}.pptx"
                zf.writestr(fname, pptx_bytes)
                generated.append(prov)
            except Exception as exc:
                flash(f"{prov} generation failed: {exc}", "error")

    if not generated:
        return redirect(url_for("appendix_a_page"))

    zip_buf.seek(0)

    # If only one provider was generated, return the PPTX directly
    if len(generated) == 1:
        prov = generated[0]
        fname = f"A – RF Schematics & GA Drawings - {prov}.pptx"
        with zipfile_mod.ZipFile(io.BytesIO(zip_buf.getvalue())) as zf:
            pptx_bytes = zf.read(fname)
        return send_file(
            io.BytesIO(pptx_bytes),
            as_attachment=True,
            download_name=fname,
            mimetype="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        )

    zip_name = f"Appendix_A_{combined}.zip"
    return send_file(
        zip_buf,
        as_attachment=True,
        download_name=zip_name,
        mimetype="application/zip",
    )


# ─── Appendix B routes ────────────────────────────────────────────────────────

@app.route("/appendix-b", methods=["GET"])
def appendix_b_page():
    return render_template("appendix_b.html")


@app.route("/appendix-b/generate", methods=["POST"])
def generate_b():
    session_id = uuid.uuid4().hex
    errors = []

    def require_field(name, label):
        val = request.form.get(name, "").strip()
        if not val:
            errors.append(f"{label} is required.")
        return val

    date            = require_field("date",            "Date")
    p01_date        = require_field("p01_date",        "P01 date")
    author          = require_field("author",          "Author")
    checked         = require_field("checked",         "Checked")
    approved        = require_field("approved",        "Approved")
    station_address = require_field("station_address", "Station address")

    # ── 3D station image ──────────────────────────────────────────────────────
    image_3d_file = request.files.get("image_3d")
    if not image_3d_file or not image_3d_file.filename:
        errors.append("3D station image is required.")
    elif Path(image_3d_file.filename).suffix.lower() not in ALLOWED_IMAGE_EXT:
        errors.append("3D image must be an image file (PNG, JPG, etc.).")

    # ── Input RF Predictions PPTX ─────────────────────────────────────────────
    input_pptx_file = request.files.get("input_pptx")
    if not input_pptx_file or not input_pptx_file.filename:
        errors.append("RF Predictions report (.pptx) is required.")
    elif Path(input_pptx_file.filename).suffix.lower() != ".pptx":
        errors.append("RF Predictions report must be a .pptx file.")

    if errors:
        for e in errors:
            flash(e, "error")
        return redirect(url_for("appendix_b_page"))

    # ── Save uploads ──────────────────────────────────────────────────────────
    subdir = session_id

    image_3d_path = save_upload(
        image_3d_file,
        subdir,
        secure_filename(image_3d_file.filename),
    )

    input_pptx_path = save_upload(
        input_pptx_file,
        subdir,
        secure_filename(input_pptx_file.filename),
    )

    # ── Dynamic providers & logos ─────────────────────────────────────────────
    # Read provider_name_0 / provider_logo_0 … provider_name_N / provider_logo_N
    # from the form.  Indices may have gaps (deleted rows); scan 0–29.
    # Falls back to the default four providers if the user submitted nothing.
    DEFAULT_B_PROVIDERS = ["THREE", "VODAFONE", "VMO2", "EE"]

    custom_providers = []   # list of {"name": str, "logo": str|None}
    for i in range(30):
        pname = request.form.get(f"provider_name_{i}", "").strip().upper()
        if not pname:
            continue
        logo_file = request.files.get(f"provider_logo_{i}")
        logo_path_val = None
        if logo_file and logo_file.filename:
            ext = Path(logo_file.filename).suffix.lower()
            if ext in ALLOWED_IMAGE_EXT:
                logo_saved = save_upload(logo_file, subdir, f"logo_{i}{ext}")
                logo_path_val = str(logo_saved)
        custom_providers.append({"name": pname, "logo": logo_path_val})

    if not custom_providers:
        custom_providers = [{"name": p, "logo": None} for p in DEFAULT_B_PROVIDERS]

    # Sort longer names first so substring matches go in the right direction
    # (e.g. "VODAFONE" before "VMO2", "THREE" before "EE")
    b_provider_names = sorted(
        [p["name"] for p in custom_providers],
        key=len, reverse=True,
    )
    logos = {p["name"]: p["logo"] for p in custom_providers}

    # ── Check shared template exists ──────────────────────────────────────────
    if not TEMPLATE_B.exists():
        flash(
            "Missing Appendix B template: templateB.pptx. "
            "Please place it in the server app folder.",
            "error",
        )
        return redirect(url_for("appendix_b_page"))

    # ── Extract RF map images from the input PPTX ─────────────────────────────
    extract_dir = UPLOAD_DIR / subdir / "extracted"
    try:
        extract_pptx_images(
            str(input_pptx_path),
            str(extract_dir),
            providers=b_provider_names,   # user-defined, sorted longest-first
        )
    except Exception as exc:
        flash(f"Failed to read RF Predictions PPTX: {exc}", "error")
        return redirect(url_for("appendix_b_page"))

    # ── Build image_entries per provider from extracted files ─────────────────
    buckets = {}
    for prov in b_provider_names:
        prov_dir = extract_dir / prov
        if prov_dir.exists():
            entries = [
                (prov_dir / fname, fname)
                for fname in sorted(os.listdir(str(prov_dir)))
                if fname.lower().endswith(".png")
            ]
            if entries:
                buckets[prov] = entries

    active_providers = [p for p in b_provider_names if p in buckets]
    if not active_providers:
        flash(
            "No RF map slides were recognised in the uploaded PPTX. "
            "Make sure the report contains slides for EE, Three, Vodafone or VMO2.",
            "error",
        )
        return redirect(url_for("appendix_b_page"))

    # ── Station name for output filename ─────────────────────────────────────
    try:
        _, _, combined = extract_station_info(str(image_3d_path))
    except Exception:
        combined = "UNKNOWN"

    # ── Generate one PPTX per active provider ────────────────────────────────
    zip_buf   = io.BytesIO()
    generated = []

    with zipfile_mod.ZipFile(zip_buf, "w", zipfile_mod.ZIP_DEFLATED) as zf:
        for prov in active_providers:
            entries = buckets[prov]
            try:
                pptx_bytes = generate_pptx_b(
                    template_path = str(TEMPLATE_B),
                    logo_path     = logos.get(prov),
                    image_3d_path = str(image_3d_path),
                    date          = date,
                    p01_date      = p01_date,
                    author        = author,
                    checked       = checked,
                    approved      = approved,
                    address       = station_address,
                    image_entries = entries,
                    provider      = prov,
                )
                fname = f"B – RF Predictions - {prov}.pptx"
                zf.writestr(fname, pptx_bytes)
                generated.append(prov)
            except Exception as exc:
                flash(f"{prov} generation failed: {exc}", "error")

    if not generated:
        return redirect(url_for("appendix_b_page"))

    zip_buf.seek(0)

    if len(generated) == 1:
        prov  = generated[0]
        fname = f"B – RF Predictions - {prov}.pptx"
        with zipfile_mod.ZipFile(io.BytesIO(zip_buf.getvalue())) as zf:
            pptx_bytes = zf.read(fname)
        return send_file(
            io.BytesIO(pptx_bytes),
            as_attachment=True,
            download_name=fname,
            mimetype="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        )

    zip_name = f"Appendix_B_{combined}.zip"
    return send_file(
        zip_buf,
        as_attachment=True,
        download_name=zip_name,
        mimetype="application/zip",
    )


# ─── Entry point ──────────────────────────────────────────────────────────────

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5050))
    app.run(host="0.0.0.0", debug=False, port=port)
