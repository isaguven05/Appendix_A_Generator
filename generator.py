"""
Core PPTX generation logic for Appendix A automation.
"""

import io
import os
import re
import uuid
import zipfile
import html as html_module
from pathlib import Path

from PIL import Image
from pptx import Presentation
from pptx.util import Emu, Pt
from pptx.dml.color import RGBColor

try:
    import pytesseract
    OCR_AVAILABLE = True
except ImportError:
    OCR_AVAILABLE = False


# ─── Shape bounds from template_A.pptx (in EMUs) ─────────────────────────────
# Slide 1 – Title: image rectangle
TITLE_IMAGE_BOUNDS = dict(left=349371, top=3255351, width=14414258, height=6462572)
# Slide 4 – RF Schematics: image rectangle
RF_IMAGE_BOUNDS    = dict(left=262515,  top=2511339, width=11871901, height=6783229)
# Slide 5 – GA: image rectangle
GA_IMAGE_BOUNDS    = dict(left=346087,  top=1890006, width=11818354, height=6989657)
# Slide 5 – GA: title text box
GA_TITLE_BOUNDS    = dict(left=435937,  top=616155,  width=3598287,  height=634118)


# ─── Helpers ──────────────────────────────────────────────────────────────────

def xml_escape(text: str) -> str:
    return (str(text)
            .replace("&", "&amp;")
            .replace("<", "&lt;")
            .replace(">", "&gt;")
            .replace('"', "&quot;"))


def extract_station_info(image_3d_path: str):
    """Parse '3D_DO93_TEMPLE.png' → (code='DO93', name='TEMPLE', combined='DO93_TEMPLE')."""
    stem = Path(image_3d_path).stem          # '3D_DO93_TEMPLE'
    parts = stem.split("_", 2)               # ['3D', 'DO93', 'TEMPLE']
    if len(parts) < 3:
        raise ValueError(f"Unexpected 3D image filename: {image_3d_path!r}")
    code = parts[1]
    name = parts[2]
    combined = f"{code}_{name}"
    return code, name, combined


def get_initials(full_name: str) -> str:
    """'Aleem Alvi' → 'AA'"""
    return "".join(p[0].upper() for p in full_name.strip().split() if p)


def sort_design_plans(paths: list) -> list:
    """Sort design-plan images by trailing number in filename."""
    def _key(p):
        m = re.search(r"(\d+)[^0-9]*$", Path(p).stem)
        return int(m.group(1)) if m else 0
    return sorted(paths, key=_key)


def extract_building_title(filename: str) -> str:
    """'EE_TEMPLE - Building 1 - Midway.jpeg' → 'Midway'
    Also handles secure_filename output: 'EE_TEMPLE_-_Building_1_-_Midway.jpeg'"""
    stem = Path(filename).stem
    # secure_filename() replaces spaces with underscores, so normalise first
    stem_normalised = stem.replace("_-_", " - ")
    parts = stem_normalised.split(" - ")
    title = parts[-1].strip() if len(parts) >= 2 else stem
    # secure_filename replaces spaces with underscores inside the title too
    title = title.replace("_", " ")
    return title


# ─── GA slide detection ───────────────────────────────────────────────────────

def detect_ga_count(pptx_path: str) -> int:
    """
    Count slides in pptx_path that contain the BoH Areas legend.

    Strategy:
    1. Text search in XML (fast).
    2. OCR on embedded slide images (fallback).
    Returns 0 if nothing found (caller can apply override).
    """
    # Method 1: XML text search
    count = _xml_text_count(pptx_path, "BoH")
    if count > 0:
        return count

    # Method 2: OCR on embedded images
    if OCR_AVAILABLE:
        count = _ocr_count(pptx_path, "BoH")
    return count


def _xml_text_count(pptx_path: str, keyword: str) -> int:
    count = 0
    try:
        with zipfile.ZipFile(pptx_path) as zf:
            slide_files = sorted(
                f for f in zf.namelist()
                if re.match(r"ppt/slides/slide\d+\.xml$", f)
            )
            for sf in slide_files:
                raw = zf.read(sf).decode("utf-8", errors="ignore")
                decoded = html_module.unescape(raw)
                if keyword in decoded:
                    count += 1
    except Exception:
        pass
    return count


def _ocr_count(pptx_path: str, keyword: str) -> int:
    count = 0
    try:
        prs = Presentation(pptx_path)
        for slide in prs.slides:
            found = False
            for shape in slide.shapes:
                if found:
                    break
                if shape.shape_type == 13:  # MSO_SHAPE_TYPE.PICTURE
                    try:
                        img = Image.open(io.BytesIO(shape.image.blob))
                        # Downsample for speed
                        max_w = 1200
                        if img.width > max_w:
                            ratio = max_w / img.width
                            img = img.resize(
                                (max_w, int(img.height * ratio)),
                                Image.LANCZOS
                            )
                        text = pytesseract.image_to_string(img, config="--psm 6")
                        if keyword in text:
                            count += 1
                            found = True
                    except Exception:
                        pass
    except Exception:
        pass
    return count


# ─── python-pptx slide helpers ────────────────────────────────────────────────

def delete_slide(prs: Presentation, index: int):
    """Delete slide at index from a Presentation (in-place)."""
    xml_slides = prs.slides._sldIdLst
    elem = xml_slides[index]
    r_ns = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    rId = elem.get(f"{{{r_ns}}}id")
    prs.part.drop_rel(rId)
    xml_slides.remove(elem)


def add_image_shape(slide, bounds: dict, image_path: str):
    """Add a picture to slide at given EMU bounds dict."""
    slide.shapes.add_picture(
        image_path,
        Emu(bounds["left"]),
        Emu(bounds["top"]),
        Emu(bounds["width"]),
        Emu(bounds["height"]),
    )


def add_text_box(slide, bounds: dict, text: str, font_size_pt: int = 20,
                 bold: bool = True):
    """Add a plain text box to a slide at given EMU bounds dict."""
    from pptx.util import Pt
    txBox = slide.shapes.add_textbox(
        Emu(bounds["left"]),
        Emu(bounds["top"]),
        Emu(bounds["width"]),
        Emu(bounds["height"]),
    )
    tf = txBox.text_frame
    tf.word_wrap = False
    p = tf.paragraphs[0]
    run = p.add_run()
    run.text = text
    run.font.size = Pt(font_size_pt)
    run.font.bold = bold
    run.font.color.rgb = RGBColor(0, 0, 0)
    return txBox


# ─── XML-level post-processing ────────────────────────────────────────────────

def _replace_in_bytes(content: bytes, replacements: dict) -> bytes:
    """
    Replace XML-escaped placeholder strings in raw XML bytes.
    Both the escaped form (&lt;&lt;FOO&gt;&gt;) and unescaped form (<<FOO>>)
    are searched; replacement values are XML-escaped automatically.
    """
    text = content.decode("utf-8")
    for raw_key, raw_value in replacements.items():
        if raw_key is None or raw_value is None:
            continue
        escaped_key = xml_escape(raw_key)
        escaped_value = xml_escape(str(raw_value))
        text = text.replace(escaped_key, escaped_value)
        # Also handle unescaped form (shouldn't happen in valid XML but be safe)
        if raw_key != escaped_key:
            text = text.replace(raw_key, escaped_value)
    return text.encode("utf-8")


def _install_slidenum_field(content: bytes) -> bytes:
    """
    Replace literal <NUM> text run in master XML with a PPTX slidenum field.
    This makes each slide show its own number automatically.
    """
    text = content.decode("utf-8")
    field_id = "{" + str(uuid.uuid4()).upper() + "}"
    # The escaped literal: <a:t>&lt;NUM&gt;</a:t>
    old_run = "<a:t>&lt;NUM&gt;</a:t>"
    new_field = (
        f'</a:r><a:fld id="{field_id}" type="slidenum">'
        f'<a:rPr lang="en-GB" smtClean="0" dirty="0"/>'
        f"<a:t>\u2039#\u203a</a:t>"
        f"</a:fld><a:r>"
    )
    # Replace inside <a:r>…</a:r> containing just <NUM>
    text = re.sub(
        r"<a:r>(<a:rPr[^/]*/>\s*)?<a:t>&lt;NUM&gt;</a:t></a:r>",
        f'<a:fld id="{field_id}" type="slidenum">'
        f'<a:rPr lang="en-GB" smtClean="0" dirty="0"/>'
        f"<a:t>\u2039#\u203a</a:t>"
        f"</a:fld>",
        text,
    )
    return text.encode("utf-8")


def _post_process_pptx(pptx_bytes: bytes, replacements: dict) -> bytes:
    """
    Open PPTX bytes as zip, apply text replacements across master, layouts,
    and all slides, then return modified bytes.
    """
    in_buf = io.BytesIO(pptx_bytes)
    out_buf = io.BytesIO()

    with zipfile.ZipFile(in_buf, "r") as zin, \
         zipfile.ZipFile(out_buf, "w", zipfile.ZIP_DEFLATED) as zout:

        for item in zin.infolist():
            data = zin.read(item.filename)

            if item.filename.endswith(".xml"):
                is_master = "slideMasters/" in item.filename
                is_layout = "slideLayouts/" in item.filename
                is_slide  = re.match(r".*ppt/slides/slide\d+\.xml$", item.filename)

                if is_master or is_layout or is_slide:
                    data = _replace_in_bytes(data, replacements)

                if is_master or is_layout:
                    data = _install_slidenum_field(data)

            zout.writestr(item, data)

    return out_buf.getvalue()


# ─── Main generation function ─────────────────────────────────────────────────

def generate_pptx(
    template_path: str,
    input_pptx_path: str,
    image_3d_path: str,
    p01_date: str,
    p02_date: str,
    author: str,
    checked: str,
    approved: str,
    station_address: str,
    design_plan_images: list,   # file paths sorted by plan number
    building_images: list,      # list of file paths
    building_titles: list = None,  # optional pre-extracted titles (preserves '&' etc.)
    ga_count_override: int = None,
) -> bytes:
    """
    Generate the Appendix A PPTX and return it as bytes.

    Output slide order:
      1  – Title slide
      2  – Design Change Summary
      3… – 2_GA slides  (ga_count of them)
      …  – RF Schematics slides (one per design plan image)
      …  – GA slides (one per building image)
    """

    # ── Station metadata ────────────────────────────────────────────────────
    code, station_name, combined = extract_station_info(image_3d_path)
    d_ini = get_initials(author)
    c_ini = get_initials(checked)
    a_ini = get_initials(approved)

    design_plan_images = sort_design_plans(design_plan_images)

    # ── GA count ────────────────────────────────────────────────────────────
    if ga_count_override is not None and ga_count_override >= 0:
        ga_count = int(ga_count_override)
    else:
        ga_count = detect_ga_count(input_pptx_path)

    # ── Load template ────────────────────────────────────────────────────────
    prs = Presentation(template_path)

    # Build layout map by name
    layout_map = {lay.name: lay for lay in prs.slide_layouts}
    lay_2ga = layout_map["2_GA"]
    lay_rf  = layout_map["RF Schematics"]
    lay_ga  = layout_map["GA"]

    # Template has 5 slides; we keep slides 1 & 2, delete 3, 4, 5
    for _ in range(3):
        delete_slide(prs, 2)   # always delete index 2 (shifts after each deletion)

    # ── 2_GA slides (GA drawings – no image per spec) ────────────────────────
    for _ in range(ga_count):
        prs.slides.add_slide(lay_2ga)

    # ── RF Schematics slides ─────────────────────────────────────────────────
    rf_slides = []
    for img_path in design_plan_images:
        slide = prs.slides.add_slide(lay_rf)
        rf_slides.append((slide, img_path))

    # ── GA (building) slides ─────────────────────────────────────────────────
    ga_slides = []
    for img_path in building_images:
        slide = prs.slides.add_slide(lay_ga)
        ga_slides.append((slide, img_path))

    # ── Title slide – replace 3D image and placeholders ──────────────────────
    title_slide = prs.slides[0]
    # Remove the <<3D_IMAGE>> text box (it's a labelling shape, not the frame)
    for shape in list(title_slide.shapes):
        if shape.has_text_frame and "<<3D_IMAGE>>" in shape.text_frame.text:
            sp = shape._element
            sp.getparent().remove(sp)
            break
    # Add the 3D image inside the existing Rectangle frame
    add_image_shape(title_slide, TITLE_IMAGE_BOUNDS, image_3d_path)

    # ── RF Schematics slides – insert cable images ────────────────────────────
    for slide, img_path in rf_slides:
        add_image_shape(slide, RF_IMAGE_BOUNDS, img_path)

    # ── GA slides – insert building images and titles ─────────────────────────
    for idx, (slide, img_path) in enumerate(ga_slides):
        if building_titles and idx < len(building_titles):
            bld_title = building_titles[idx]
        else:
            bld_title = extract_building_title(img_path)
        add_image_shape(slide, GA_IMAGE_BOUNDS, img_path)
        add_text_box(slide, GA_TITLE_BOUNDS, bld_title, font_size_pt=24, bold=True)

    # ── Save to bytes ─────────────────────────────────────────────────────────
    buf = io.BytesIO()
    prs.save(buf)
    pptx_bytes = buf.getvalue()

    # ── XML post-processing: text replacements + slidenum field ───────────────
    replacements = {
        # Master / layout level
        "<<STATION_NAME>>": combined,
        "<<ADDRESS>>":      station_address,
        "<P01>":            p01_date,
        "<P02>":            p02_date,
        "<D>":              d_ini,
        "<C>":              c_ini,
        "<A>":              a_ini,
        # Title slide level
        "<<STATION>>":      station_name,
        "<<CODE>>":         code,
        "<<Author>>":       author,
        "<<Checked>>":      checked,
        "<<Approved>>":     approved,
    }

    pptx_bytes = _post_process_pptx(pptx_bytes, replacements)
    return pptx_bytes
