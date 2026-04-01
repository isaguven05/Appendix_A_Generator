"""
Appendix B Automation Engine
============================
This is the brain of the tool. It does two main jobs:

  1. EXTRACTION — Opens the input PowerPoint (the RF Predictions report),
     reads each slide, figures out what's on it (which provider, technology,
     frequency, metric, and building), and saves the map image as a PNG file.

  2. BUILDING — Takes those saved PNG images and assembles a brand-new,
     properly ordered PowerPoint for each provider (EE, Vodafone, etc.),
     slotting each image into the right template slide.

All file paths and settings are passed in from the UI at runtime —
nothing is hardcoded that the user needs to change.
"""

import os, re, copy, shutil
from pptx import Presentation
from pptx.util import Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import MSO_AUTO_SIZE

# PIL (Pillow) is used to read image dimensions so logos scale correctly.
# We try to import it, but the rest of the code still works if it's missing.
try:
    from PIL import Image as PILImage
    HAS_PIL = True
except ImportError:
    HAS_PIL = False

# This is the XML namespace string that PowerPoint uses internally to track
# relationships between slides and their embedded images. We need it when
# duplicating slides so image links don't break.
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


# =============================================================================
# SLIDE SORT ORDER
# =============================================================================
# We want slides in the final output to appear in a consistent order.
# This list defines the "correct" order for technology/frequency/metric combos.
# e.g. GSM 1800 RXLEV comes first, then GSM 1800 SNIR, then UMTS, then LTE, etc.

_PRIMARY_ORDER = [
    ("GSM",  "1800", "RXLEV"),  ("GSM",  "1800", "SNIR"),
    ("UMTS", "2100", "RSCP"),   ("UMTS", "2100", "ECIO"),
    ("LTE",  "1800", "RSRP"),   ("LTE",  "1800", "SNIR"),
    ("LTE",  "2100", "RSRP"),   ("LTE",  "2100", "SNIR"),
    ("LTE",  "2300", "RSRP"),   ("LTE",  "2300", "SNIR"),
    ("LTE",  "2600", "RSRP"),   ("LTE",  "2600", "SNIR"),
    ("5GNR", "3500", "SSRSRP"), ("5GNR", "3500", "SNIR"),
]

# Convert the list above into a lookup dictionary so we can instantly find
# the rank (position number) of any tech/freq/metric combo.
# e.g. _PRIMARY_RANK[("GSM", "1800", "RXLEV")] → 0  (first)
#      _PRIMARY_RANK[("LTE", "1800", "RSRP")]   → 4  (fifth)
_PRIMARY_RANK = {key: i for i, key in enumerate(_PRIMARY_ORDER)}

# The default building order used if the user hasn't set one in the UI.
# Buildings will appear in slides in this sequence.
DEFAULT_BUILDING_ORDER = ["TICKET_HALL", "MIDWAY", "BASEMENT", "STAIRS", "CONCOURSE", "PLATFORM"]


def _make_sort_key(building_order):
    """
    Creates a sorting function that we can pass to Python's .sort() method.

    The sort function scores each slide entry so that when the list is sorted,
    slides end up in the right order:
      - First sorted by technology/frequency/metric (GSM before LTE, etc.)
      - Then by building type (e.g. Ticket Hall before Platform)
      - Then alphabetically by filename as a tiebreaker

    'building_order' is the list the user set in the UI, e.g.:
      ["Ticket Hall", "Midway", "Platform 1&2"]
    """
    # Build a rank lookup for buildings: {"TICKET_HALL": 0, "MIDWAY": 1, ...}
    # We normalise to uppercase + underscores so matching is consistent.
    building_rank = {b.upper().replace(" ", "_"): i for i, b in enumerate(building_order)}

    def key(entry):
        # Each 'entry' is a tuple: (provider, file, path, tech, mhz, metric, building)
        provider, file, path, tech, mhz, metric, building = entry
        return (
            # Score 1: position in the tech/freq/metric order (lower = earlier)
            _PRIMARY_RANK.get((tech, mhz, metric), len(_PRIMARY_ORDER)),
            # Score 2: position in the building order (lower = earlier)
            # If the building isn't in the list, it gets pushed to the end.
            building_rank.get(building.upper(), len(building_order)),
            # Score 3: filename alphabetically, just as a tiebreaker
            file,
        )
    return key


# =============================================================================
# TEMPLATE SLIDE MAP
# =============================================================================
# The template PowerPoint has one "blank" slide for each tech/freq/metric type.
# This dictionary maps each combo to the index (position) of its template slide.
#
# Index numbers (0-based):
#   [0] Title slide   [1] Exec summary   [2] Logo/operator slide
#   [3] GSM SNIR      [4] GSM RXLEV
#   [5] UMTS RSCP     [6] UMTS Ec/Io
#   [7–14] LTE bands  [15] NR SS-RSRP    [16] NR SS-SNIR
#
# So when we want to add a GSM 1800 RXLEV slide, we copy template slide #4.

TEMPLATE_MAP = {
    ("GSM",  "1800", "RXLEV"):   4,   ("GSM",  "1800", "SNIR"):   3,
    ("UMTS", "2100", "RSCP"):    5,   ("UMTS", "2100", "ECIO"):   6,
    ("LTE",  "1800", "RSRP"):    7,   ("LTE",  "1800", "SNIR"):   8,
    ("LTE",  "2100", "RSRP"):    9,   ("LTE",  "2100", "SNIR"):  10,
    ("LTE",  "2300", "RSRP"):   11,   ("LTE",  "2300", "SNIR"):  12,
    ("LTE",  "2600", "RSRP"):   13,   ("LTE",  "2600", "SNIR"):  14,
    ("5GNR", "3500", "SSRSRP"): 15,   ("5GNR", "3500", "SNIR"):  16,
}

# The template slide index that contains the big operator logo placeholder.
LOGO_SLIDE_IDX = 2

# PowerPoint measures positions and sizes in EMUs (English Metric Units).
# These are the full slide dimensions — used when centring logos.
# 1 inch = 914400 EMUs, so this is a 16.54" × 11.69" slide (A3 landscape).
_SLIDE_W = 15113000
_SLIDE_H = 10693400


# =============================================================================
# EXTRACTION HELPERS
# These functions all read the text from a slide and figure out what it means.
# =============================================================================

def get_slide_text(slide):
    """
    Grabs all visible text from every text box on a slide and joins it
    into one big uppercase string.

    We uppercase everything so our keyword searches don't have to worry
    about capitalisation — "Midway", "MIDWAY", and "midway" all become "MIDWAY".
    """
    return "".join(s.text + "\n" for s in slide.shapes if s.has_text_frame).upper()


def extract_provider(text, providers):
    """
    Looks at the slide text and works out which mobile provider it belongs to
    (e.g. EE, Vodafone, Three).

    It checks each provider name the user has configured — if the name appears
    anywhere in the slide text, that's our match.
    Returns "UNKNOWN" if none of the provider names are found.
    """
    for p in providers:
        if p.upper() in text:
            return p
    return "UNKNOWN"


def extract_technology(text):
    """
    Determines the radio technology used on this slide.
    Checks for keywords in order of specificity (5G first, then LTE, etc.)
    so we don't accidentally misclassify something.

    Returns one of: "5GNR", "LTE", "UMTS", "GSM"
    Defaults to "LTE" if nothing is found (most common type).
    """
    if "5G" in text or "NR" in text: return "5GNR"
    if "LTE" in text: return "LTE"
    if "UMTS" in text: return "UMTS"
    if "GSM" in text: return "GSM"
    return "LTE"


def extract_mhz(text, tech):
    """
    Extracts the frequency band in MHz from the slide text.

    5GNR is always 3500 MHz, so we skip the search for that.
    For everything else we look for patterns like "1800MHz" or "2.1GHz"
    using a regex (a search pattern).

    A regex like r"(\\d{3,4})\\s*MHZ" means:
      \\d{3,4}  → 3 or 4 digit number
      \\s*      → optional spaces
      MHZ       → followed by the text "MHZ"

    Returns None if no frequency is found (those slides get skipped).
    """
    if tech == "5GNR": return "3500"

    # Try to find a frequency written in MHz, e.g. "1800MHz" or "1800 MHz"
    mhz = re.search(r"(\d{3,4})\s*MHZ", text)
    if mhz: return mhz.group(1)

    # Try GHz format instead, e.g. "2.1GHz" → convert to "2100"
    ghz = re.findall(r"(\d+\.?\d*)\s*GHZ", text)
    if ghz: return str(int(float(ghz[0]) * 1000))

    return None  # No frequency found — this slide will be skipped


def extract_metric(text, tech):
    """
    Figures out which signal quality metric is shown on this slide.
    Different technologies use different metrics:
      - GSM uses RXLEV (signal level) and SNIR (signal quality)
      - UMTS uses RSCP and Ec/Io
      - LTE uses RSRP and SNIR
      - 5GNR uses SS-RSRP (shown as SSRSRP) and SNIR

    We check for keywords in the slide text, with some special cases
    (e.g. if the slide says "RSRP" but the tech is GSM, it's actually RXLEV).
    """
    if "SNIR" in text: return "SNIR"
    if "RSRP" in text:
        if tech == "GSM": return "RXLEV"           # GSM doesn't have RSRP — must be RXLEV
        return "SSRSRP" if tech == "5GNR" else "RSRP"
    if "RSCP" in text: return "RSCP"
    if "ECIO" in text or "EC/IO" in text: return "ECIO"
    if "RXLEV" in text: return "RXLEV"
    if tech == "GSM": return "RXLEV"               # Default for GSM
    return "RSRP"                                  # Default for everything else


def extract_building(text, building_order=None):
    """
    Extracts the building name from slide text.

    Source slides always label the building after a code and colon, e.g.:
      "BUILDING 1: TICKET HALL"
      "B113: PLATFORMS 1&2"
      "BUILDING 1: PLATFORM - ESCALATOR MACHINE CHAMBERS 1 & 3"

    We read the name directly from this pattern so it works for any building
    name regardless of what the user has configured — no pre-configuration or
    exact keyword match required.

    The building_order list (from the UI) is used ONLY for sorting slides into
    the correct order. Slides whose building name is not in the list are still
    included; they sort to the end.

    Returns the building name with spaces replaced by underscores (filename-safe),
    e.g. "TICKET_HALL", "PLATFORMS_1&2", "PLATFORM_-_ESCALATOR_MACHINE_CHAMBERS_1_&_3".
    Returns "UNKNOWN" if no building label pattern is found.
    """
    # Primary method: extract the building name that appears after a slide label.
    # Two observed patterns:
    #   "BUILDING 1: TICKET HALL"   — standard format
    #   "B113: PLATFORMS 1&2"       — station-code format (single letter + digits)
    m = re.search(r'(?:BUILDING\s*\d+|[A-Z]\d{2,})\s*:\s*(.+?)(?:\n|$)', text)
    if m:
        name = m.group(1).strip()
        if name:
            return name.replace(" ", "_")

    # Fallback: keyword search against the user-provided building order.
    # Only reached if the slide has no standard label pattern.
    if building_order:
        candidates = [(b.upper().replace("_", " "), b) for b in building_order]
        # Sort longest keyword first so "PLATFORM 1&2" matches before "PLATFORM"
        candidates.sort(key=lambda x: len(x[0]), reverse=True)
        for keyword, stored in candidates:
            if keyword in text:
                return stored

    return "UNKNOWN"  # No building label found on this slide


def process_pptx(pptx_path, output_folder, providers, building_order=None):
    """
    STEP 1 — The extraction phase.

    Opens the input PowerPoint (the RF Predictions report) and goes through
    every slide one by one. For each slide it:
      1. Reads all the text to identify: provider, tech, frequency, metric, building
      2. Finds the map image embedded in the slide (shape type 13 = picture)
      3. Saves that image as a PNG file into a subfolder named after the provider

    The saved filename encodes all the metadata, e.g.:
      EE_LTE_1800_RSRP_TICKET_HALL.png
      EE_GSM_1800_RXLEV_MIDWAY.png

    This naming system means we can read everything we need back from the
    filename later, without having to open the original PowerPoint again.

    Returns the total number of images extracted.
    """
    prs = Presentation(pptx_path)
    os.makedirs(output_folder, exist_ok=True)  # Create the output folder if it doesn't exist
    extracted = 0

    for i, slide in enumerate(prs.slides):
        # Pull all text from this slide into one uppercase string
        text = get_slide_text(slide)

        # Work out what this slide is about
        provider = extract_provider(text, providers)
        tech     = extract_technology(text)
        mhz      = extract_mhz(text, tech)
        metric   = extract_metric(text, tech)
        building = extract_building(text, building_order)

        # If we couldn't find a frequency, we can't reliably identify this slide — skip it
        if not mhz:
            continue

        # Create a subfolder for this provider (e.g. "extracted_images/EE/")
        folder = os.path.join(output_folder, provider)
        os.makedirs(folder, exist_ok=True)

        # Find any picture shapes on the slide (shape_type 13 = picture/image)
        for shape in slide.shapes:
            if shape.shape_type == 13:
                # Build the filename from all the extracted metadata
                name = f"{provider}_{tech}_{mhz}_{metric}_{building}.png"
                path = os.path.join(folder, name)

                # If a file with this name already exists (duplicate slide),
                # append the slide index to avoid overwriting it
                if os.path.exists(path):
                    name = f"{provider}_{tech}_{mhz}_{metric}_{building}_s{i}.png"
                    path = os.path.join(folder, name)

                # Write the raw image bytes to disk
                with open(path, "wb") as f:
                    f.write(shape.image.blob)
                extracted += 1

    return extracted


# =============================================================================
# SLIDE BUILDING HELPERS
# These functions handle the mechanics of copying and modifying slides.
# =============================================================================

def duplicate_slide(prs, tmpl_slide):
    """
    Copies a slide from the template into the output presentation.

    PowerPoint stores slides as XML files inside a zip archive (.pptx).
    Each image on a slide has an internal "relationship ID" (rId) that points
    to the actual image file in the zip. When we copy a slide, we have to:
      1. Register all the images from the old slide into the new slide
      2. Update all the rId references in the copied XML so they point to
         the newly registered images (not the old ones from the template)

    Without this step, images would be missing or broken in the output file.
    """
    # Add a new blank slide (using the same layout as the template slide)
    new_slide = prs.slides.add_slide(tmpl_slide.slide_layout)

    # Copy over image relationships and build a mapping from old rId → new rId
    rId_map = {}
    for old_rId, rel in tmpl_slide.part.rels.items():
        if "image" not in rel.reltype:
            continue  # Only care about image relationships
        new_rId = new_slide.part.relate_to(rel._target, rel.reltype)
        rId_map[old_rId] = new_rId

    # These are the XML attribute names that hold image relationship IDs
    embed_attr = f"{{{R_NS}}}embed"
    link_attr  = f"{{{R_NS}}}link"

    # Deep-copy every shape from the template slide into the new slide,
    # updating any image rId references along the way
    for shape in tmpl_slide.shapes:
        el = copy.deepcopy(shape.element)  # Deep copy = full independent copy of the XML
        for node in el.iter():             # Walk every XML node inside this shape
            for attr in (embed_attr, link_attr):
                if attr in node.attrib and node.attrib[attr] in rId_map:
                    # Swap the old rId for the new one
                    node.attrib[attr] = rId_map[node.attrib[attr]]
        # Insert the shape into the new slide's shape tree (before the extension list)
        new_slide.shapes._spTree.insert_element_before(el, 'p:extLst')

    return new_slide


def _fix_run_after_replace(run, old_val, new_val):
    """
    After we've swapped a placeholder like "STATION-NAME" for the real value,
    this function tidies up the formatting:
      - Forces the text colour to black (template may have it in red/white)
      - Converts the text to Title Case (e.g. "LAMBETH_NORTH" → "Lambeth North")

    A "run" in PowerPoint terms is a continuous piece of text with the same
    formatting — a paragraph can contain multiple runs.
    """
    if old_val in ("STATION-NAME", "OPERATOR-NAME",
                   "[STATIONCODE_STATIONNAME]", "[STATION CODE]"):
        run.font.color.rgb = RGBColor(0, 0, 0)  # Force black
        # Title-case station names (e.g. "LAMBETH_NORTH" → "Lambeth North")
        # For operator names: title-case but keep short all-caps acronyms
        # uppercase (e.g. "EE" stays "EE" not "Ee"; "THREE" → "Three")
        if old_val == "OPERATOR-NAME":
            titled = " ".join(
                w if (w.isupper() and len(w) <= 3) else w.capitalize()
                for w in new_val.split()
            )
            run.text = run.text.replace(new_val, titled)
        else:
            run.text = run.text.replace(
                new_val, new_val.replace("_", " ").title()
            )


def _initials(name):
    """
    Converts a person's name to their initials in uppercase.

    Examples:
      "John Smith"  → "JS"
      "MY"          → "MY"   (already initials — returned as-is in upper)
      "mark young"  → "MY"
      ""            → ""
    """
    parts = name.strip().split()
    if not parts:
        return ""
    if len(parts) == 1:
        return parts[0].upper()          # Single token — use as-is (e.g. "MY" → "MY")
    return "".join(p[0].upper() for p in parts if p)  # Multi-word → first letters


def replace_in_master(prs, replacements):
    """
    Applies text find-and-replace across every text shape in the slide master.

    This is how we fill in the per-run metadata that lives on the master
    (DRAWN: <<d>>, CHECKED: <<c>>, APPROVED: <<a>>, <<po1>>, <<po2>>)
    so that all content slides inherit the correct values automatically.

    'replacements' is a dict like:
      {"<<d>>": "MY", "<<c>>": "NA", "<<a>>": "CK",
       "<<po1>>": "03/10/2022", "<<po2>>": "01/07/2025"}
    """
    master = prs.slide_master
    for shape in master.shapes:
        if not shape.has_text_frame:
            continue
        for para in shape.text_frame.paragraphs:
            for run in para.runs:
                for old, new in replacements.items():
                    if old in run.text:
                        run.text = run.text.replace(old, new)


def replace_logo_in_master(prs, logo_path):
    """
    Replace the <<logo>> placeholder on the slide master with the provider
    logo image at a fixed position, so it appears on ALL slides automatically.

    Fixed position and size (as specified):
      Left  = 1050 pt   Top    = 62 pt
      Width =   43 pt   Height = 43 pt
    """
    master = prs.slide_master

    # Remove the <<logo>> placeholder shape from the master
    for shape in list(master.shapes):
        is_logo = False
        if shape.name and "logo" in shape.name.lower():
            is_logo = True
        elif shape.has_text_frame:
            try:
                if "<<logo>>" in shape.text_frame.text.lower():
                    is_logo = True
            except Exception:
                pass
        if is_logo:
            master.shapes._spTree.remove(shape.element)

    # If no logo supplied, just remove the placeholder text — no image added
    if not logo_path or not os.path.exists(logo_path):
        return

    # Insert the logo at the exact position and size requested
    master.shapes.add_picture(
        logo_path,
        left   = Pt(1050),
        top    = Pt(62),
        width  = Pt(43),
        height = Pt(43),
    )


def force_black_runs(slide):
    """
    Scans every text run on a slide and changes any red text to black.

    The template sometimes has red placeholder text (e.g. "GSM / LTE / 5GNR"
    or the station name) that should appear black in the final output.
    This function catches any that weren't fixed by the normal replacement logic.
    """
    RED = RGBColor(0xFF, 0x00, 0x00)  # Pure red in RGB
    for shape in slide.shapes:
        if not shape.has_text_frame:
            continue
        for para in shape.text_frame.paragraphs:
            for run in para.runs:
                try:
                    if run.font.color.rgb == RED:
                        run.font.color.rgb = RGBColor(0, 0, 0)  # Change to black
                except Exception:
                    pass  # Some runs don't have a colour set — just skip them


def replace_in_slide(slide, replacements):
    """
    Finds and replaces placeholder text across all text boxes and tables on a slide.

    'replacements' is a dictionary like:
      { "STATION-NAME": "Lambeth North", "OPERATOR-NAME": "EE" }

    We go through every text run in every paragraph in every shape and swap
    any matching placeholder text for the real value.
    Tables are handled separately because they have a different internal structure
    (rows → cells → paragraphs → runs).
    """
    for shape in slide.shapes:
        # Handle regular text boxes
        if shape.has_text_frame:
            for para in shape.text_frame.paragraphs:
                for run in para.runs:
                    for old, new in replacements.items():
                        if old in run.text:
                            run.text = run.text.replace(old, new)
                            _fix_run_after_replace(run, old, new)

        # Handle table cells (shape_type 19 = table)
        if shape.shape_type == 19:
            for row in shape.table.rows:
                for cell in row.cells:
                    for para in cell.text_frame.paragraphs:
                        for run in para.runs:
                            for old, new in replacements.items():
                                if old in run.text:
                                    run.text = run.text.replace(old, new)
                                    _fix_run_after_replace(run, old, new)


def _image_aspect(path):
    """
    Opens an image file and returns its width and height in pixels as a tuple.
    Returns None if Pillow (the image library) isn't installed.

    We need the original pixel dimensions to correctly scale logos and images
    while preserving their aspect ratio.
    """
    if HAS_PIL:
        try:
            with PILImage.open(path) as im:
                return im.size  # Returns (width_px, height_px)
        except Exception:
            pass
    return None


def _fit_image(img_w, img_h, box_w, box_h):
    """
    Calculates how to scale an image to fit inside a box without stretching it.

    Works by finding the largest scale factor that fits both dimensions:
      - If the image is wider than tall, width is the limiting factor
      - If the image is taller than wide, height is the limiting factor

    Also calculates offsets to centre the scaled image within the box.

    Returns: (scaled_width, scaled_height, left_offset, top_offset)
    All values are in EMUs (PowerPoint's internal unit).
    """
    scale = min(box_w / img_w, box_h / img_h)  # Scale to fit the tightest dimension
    w = int(img_w * scale)
    h = int(img_h * scale)
    # Centre the image: offset = half the leftover space
    return w, h, (box_w - w) // 2, (box_h - h) // 2


def replace_logo_placeholder(slide, logo_path, big=False):
    """
    Finds the <<logo>> placeholder text box on a slide, removes it, and
    inserts the actual provider logo image in its place.

    Two modes:
      - big=False  → small corner logo, scaled to fit the placeholder box exactly
      - big=True   → large centred logo (used on the dedicated operator slide),
                     constrained to 50% of slide width × 55% of slide height

    If the logo file doesn't exist or no <<logo>> placeholder is found,
    the function does nothing (safe to call even if there's no logo).
    """
    if not logo_path or not os.path.exists(logo_path):
        return  # No logo file — nothing to do

    for shape in list(slide.shapes):
        if not shape.has_text_frame:
            continue
        if "<<logo>>" not in shape.text_frame.text.lower():
            continue  # This isn't the logo placeholder — skip it

        # Record the placeholder's position and size
        box_left, box_top = shape.left, shape.top
        box_w,    box_h   = shape.width, shape.height

        # Get the logo's natural dimensions by adding it temporarily at 0,0.
        # This uses pptx's own DPI-aware sizing — no Pillow required.
        tmp = slide.shapes.add_picture(logo_path, 0, 0)
        nat_w, nat_h = tmp.width, tmp.height
        slide.shapes._spTree.remove(tmp.element)

        if big:
            # Full-page operator logo slide: remove the placeholder entirely since
            # there is no visible border frame to preserve here.
            slide.shapes._spTree.remove(shape.element)
            target_w = int(_SLIDE_W * 0.50)
            target_h = int(_SLIDE_H * 0.55)
            w, h, dx, dy = _fit_image(nat_w, nat_h, target_w, target_h)
            left = (_SLIDE_W - target_w) // 2 + dx
            top  = (_SLIDE_H - target_h) // 2 + dy
        else:
            # Corner logo: the <<logo>> text box IS the visible bordered frame —
            # clearing its text and keeping the shape preserves the blue border.
            # The logo picture is then appended to the shape tree so it sits
            # in front of (on top of) the now-empty frame.
            shape.text_frame.clear()
            w, h, dx, dy = _fit_image(nat_w, nat_h, box_w, box_h)
            left = box_left + dx
            top  = box_top  + dy

        # Insert the logo — appended last so it renders in front of the frame
        slide.shapes.add_picture(logo_path, left, top, width=w, height=h)
        break  # Only one <<logo>> placeholder per slide — stop after finding it


def _fill_metadata_table(slide, metadata):
    """
    Fills the metadata table on slide 1 with the user-supplied values.

    The table has a header row (row 0) and a data row (row 1).
    We find the right column for each field by matching the header text,
    then write the value into the corresponding data cell.

    This means the layout can change in the template (columns reordered,
    extra columns added) without breaking anything here.

    'metadata' is a dict like:
      {"date": "01/07/2025", "author": "JD", "checked": "SM", "approved": "KL"}
    Empty strings are left blank in the output — no placeholder text appears.
    """
    if not metadata:
        return

    # Map the user-facing key names to the header text used in the template table
    # Keys are lowercase; header text is matched case-insensitively
    field_map = {
        "date":     "date",
        "author":   "author",
        "checked":  "checked",
        "approved": "approved",
    }

    for shape in slide.shapes:
        if shape.shape_type != 19:  # 19 = table
            continue

        table = shape.table

        # Build a dict of {header_text_lowercase: column_index} from row 0
        header_cols = {}
        for c, cell in enumerate(table.rows[0].cells):
            header_text = cell.text_frame.text.strip().lower()
            header_cols[header_text] = c

        # Now fill row 1 for each metadata field
        for field, header in field_map.items():
            value = metadata.get(field, "").strip()
            if not value:
                continue  # Leave the cell empty if the user didn't fill it in

            col_idx = header_cols.get(header)
            if col_idx is None:
                continue  # Column not found in this table — skip safely

            # Get the data cell (row 1, matching column)
            cell = table.rows[1].cells[col_idx]
            tf = cell.text_frame
            tf.clear()

            # Add the value as a new run, matching the template's 11pt font
            run = tf.paragraphs[0].add_run()
            run.text = value
            run.font.size = Pt(11)  # Matches the existing table font size

        break  # Only one metadata table on this slide


def _set_building_label(slide, building):
    """
    Finds the <<BUILDING>> placeholder text box on a slide and replaces it
    with the actual building name (e.g. "Ticket Hall", "Midway").

    Underscores in the stored name are converted back to spaces for display.
    The font starts at 10pt and automatically shrinks if the text is too long
    to fit within the box, preventing overflow into adjacent elements.
    """
    for shape in slide.shapes:
        if shape.has_text_frame and "<<BUILDING>>" in shape.text:
            tf = shape.text_frame
            tf.clear()  # Wipe the placeholder text
            run = tf.paragraphs[0].add_run()
            run.text = building.replace("_", " ").title()
            run.font.size = Pt(10)
            # Shrink the font automatically if the text overflows the box
            tf.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE


def get_image_boundary(slide):
    """
    Returns the position and size of the image area on a slide, so we know
    exactly where to place the RF map image.

    Tries three things in order:
      1. Look for an existing picture shape on the slide (most reliable)
      2. Check the slide master layout for a named picture placeholder
      3. Fall back to hardcoded default coordinates (safe last resort)

    Returns: (left, top, width, height) all in EMUs
    """
    # Option 1: find an existing image on the slide and use its bounds
    for shape in slide.shapes:
        if shape.shape_type == 13:  # 13 = picture
            return shape.left, shape.top, shape.width, shape.height

    # Option 2: check the master slide layout for a placeholder named "Picture Placeholder"
    master = slide.slide_layout.slide_master
    for shape in master.shapes:
        if "Picture Placeholder" in shape.name:
            return shape.left, shape.top, shape.width, shape.height

    # Option 3: hardcoded fallback coordinates (nearly full-slide image area)
    return 599801, 508423, 11233727, 9364397


def _place_image_fitted(slide, img_path, box_left, box_top, box_w, box_h):
    """
    Adds img_path to the slide scaled to fit within the given bounding box
    while preserving the image's natural aspect ratio (no stretching).
    The image is centred inside the box — same approach as logo placement.
    """
    # Read natural pixel dimensions via a temporary add-then-remove
    tmp = slide.shapes.add_picture(img_path, 0, 0)
    nat_w, nat_h = tmp.width, tmp.height
    slide.shapes._spTree.remove(tmp.element)

    w, h, dx, dy = _fit_image(nat_w, nat_h, box_w, box_h)
    slide.shapes.add_picture(img_path,
                             box_left + dx, box_top + dy,
                             width=w, height=h)


def update_slide(slide, img_path, building, logo_path):
    """
    Populates a freshly duplicated template slide with:
      1. The RF map image (fitted inside the placeholder box, aspect ratio preserved)
      2. The building label (e.g. "Midway")
      3. The provider logo (replaces the <<logo>> text box)

    If there's already an image on the slide (the template placeholder image),
    we remove it first and insert our real map image fitted inside the same bounds.
    If there's no existing image, we use get_image_boundary() to find the bounds.
    """
    for shape in list(slide.shapes):
        if shape.shape_type == 13:  # Found a picture placeholder on the slide
            box_left, box_top = shape.left, shape.top
            box_w,    box_h   = shape.width, shape.height
            slide.shapes._spTree.remove(shape.element)
            _place_image_fitted(slide, img_path, box_left, box_top, box_w, box_h)
            _set_building_label(slide, building)
            replace_logo_placeholder(slide, logo_path, big=False)
            return

    # No placeholder image found — use boundary coordinates
    box_left, box_top, box_w, box_h = get_image_boundary(slide)
    _place_image_fitted(slide, img_path, box_left, box_top, box_w, box_h)
    _set_building_label(slide, building)
    replace_logo_placeholder(slide, logo_path, big=False)


def add_intro_slides(prs, tmpl, threed_image_path, provider, logo_path, metadata=None):
    """
    Adds the three introductory slides to the start of a provider's presentation:

      Slide 1 — Title slide with the 3D station model image and station name
      Slide 2 — Executive summary slide with station name and operator name
      Slide 3 — Full-page operator logo slide

    The station code and name are extracted from the 3D image filename.
    For example: "B137_LAMBETH_NORTH_3DMOD.png"
      → station_code = "B137"
      → station_name = "LAMBETH NORTH"

    After adding each slide, we:
      - Replace all placeholder text with real values
      - Remove any leftover empty title placeholders ("Double-click to edit")
      - Fix any text that's incorrectly coloured red in the template
    """
    # Extract station code and name from the 3D image filename
    # e.g. "B137_LAMBETH_NORTH_3DMOD.png" → parts = ["B137", "LAMBETH", "NORTH", "3DMOD"]
    parts = os.path.basename(threed_image_path).replace(".png", "").split("_")
    station_code = parts[0]                  # "B137"
    station_name = " ".join(parts[1:-1])     # "LAMBETH NORTH" (skip the last part "3DMOD")

    # --- Slide 1: Title slide ---
    slide1 = duplicate_slide(prs, tmpl.slides[0])
    replace_in_slide(slide1, {
        "[STATIONCODE_STATIONNAME]": f"{station_code}_{station_name}",
        "[STATION CODE]": station_code,
    })
    # Find the "3D MODEL IMAGE" placeholder text box and swap it for the actual image
    for shape in slide1.shapes:
        if shape.has_text_frame and "3D MODEL IMAGE" in shape.text:
            left, top, width, height = shape.left, shape.top, shape.width, shape.height
            slide1.shapes._spTree.remove(shape.element)
            slide1.shapes.add_picture(threed_image_path, left, top, width=width, height=height)
            break

    # Remove any "Project Name:" label inherited from the template —
    # it's a leftover green text box that should never appear in the output.
    # We normalise whitespace before comparing because templates sometimes use
    # non-breaking spaces (\xa0) which would cause a plain string match to fail.
    for shape in list(slide1.shapes):
        if not shape.has_text_frame:
            continue
        try:
            # Collapse all whitespace variants (spaces, \xa0, tabs, newlines)
            # into single regular spaces so the match is reliable.
            normalised = " ".join(shape.text_frame.text.split()).upper()
        except Exception:
            continue
        if "PROJECT" in normalised:
            slide1.shapes._spTree.remove(shape.element)

    # Fill in the metadata table (Date, Author, Checked, Approved) on slide 1 only
    _fill_metadata_table(slide1, metadata)

    # --- Slide 2: Executive summary slide ---
    slide2 = duplicate_slide(prs, tmpl.slides[1])
    replace_in_slide(slide2, {
        "STATION-NAME":  station_name,
        "OPERATOR-NAME": provider,
    })

    # Remove the empty "Title 1" placeholder from slides 1 and 2
    # (it shows as "Double-click to edit" if left in)
    for slide in (slide1, slide2):
        for shape in list(slide.shapes):
            if shape.shape_type == 14 and shape.name == "Title 1":
                slide.shapes._spTree.remove(shape.element)

    # --- Slide 3: Operator logo slide ---
    slide3 = duplicate_slide(prs, tmpl.slides[LOGO_SLIDE_IDX])
    replace_logo_placeholder(slide3, logo_path, big=True)  # Large centred logo

    # Remove the empty title placeholder from slide 3 too
    for shape in list(slide3.shapes):
        if shape.shape_type == 14 and shape.name == "Title 1":
            slide3.shapes._spTree.remove(shape.element)

    # Fix any red text on slides 1 and 2 (template artefacts — should be black)
    force_black_runs(slide1)
    force_black_runs(slide2)

    # Fix text alignment on slide 2: the template uses right-align on some paragraphs,
    # but we want justified (flush left and right) for a cleaner look.
    NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
    for shape in slide2.shapes:
        if shape.has_text_frame and "EXECUTIVE SUMMARY" not in shape.text_frame.text:
            for para in shape.text_frame.paragraphs:
                # Find or create the paragraph properties XML element
                pPr = para._p.find(f"{{{NS}}}pPr")
                if pPr is None:
                    from lxml import etree
                    pPr = etree.SubElement(para._p, f"{{{NS}}}pPr")
                    para._p.insert(0, pPr)
                pPr.set("algn", "just")  # "just" = justified alignment


def delete_template_slides(prs, n):
    """
    Removes the first 'n' slides from the presentation.

    When we copy the template file, it comes with its original placeholder slides
    already in it (the blank GSM, LTE, 5GNR slides etc.). After we've added all
    our real slides, we need to delete those original template slides.

    We always delete from position 0 (the first slide) 'n' times because
    once we delete slide 0, slide 1 becomes the new slide 0, and so on.
    """
    for _ in range(n):
        if not prs.slides._sldIdLst:
            break  # Nothing left to delete
        rId = prs.slides._sldIdLst[0].rId  # Get the relationship ID of the first slide
        prs.part.drop_rel(rId)             # Remove that slide's relationship from the package
        del prs.slides._sldIdLst[0]        # Remove it from the slide list


# =============================================================================
# MAIN BUILD FUNCTION
# =============================================================================

def build_provider_presentation(provider, logo_path, template_pptx, tmpl,
                                 output_folder, threed_image_path, output_path,
                                 building_order=None, metadata=None):
    """
    Builds the complete output PowerPoint for a single provider (e.g. EE).

    The process:
      1. Copy the template file to the output location (so we have a fresh start)
      2. Add the three intro slides (title, exec summary, logo)
      3. Find all the extracted PNG images for this provider
      4. Parse each filename to recover tech/frequency/metric/building info
      5. Sort them into the correct order
      6. For each image: duplicate the right template slide and insert the image
      7. Delete the original blank template slides
      8. Save the file

    Returns a tuple: (number_of_slides_placed, list_of_unrecognised_slides)
    """
    # Start from a fresh copy of the template
    shutil.copy(template_pptx, output_path)
    prs = Presentation(output_path)

    # Remember how many slides the template starts with (we'll delete these later)
    n_template_slides = len(prs.slides)

    # Add the title, exec summary, and logo slides at the top
    add_intro_slides(prs, tmpl, threed_image_path, provider, logo_path, metadata)

    # Find the folder containing this provider's extracted images
    # e.g. "work/extracted_images/EE/"
    folder = os.path.join(output_folder, provider)
    provider_slides = []  # Will hold info about each image to place
    missing = []          # Images whose tech/metric wasn't recognised

    if os.path.exists(folder):
        for file in os.listdir(folder):
            if not file.endswith(".png"):
                continue  # Skip any non-image files

            path = os.path.join(folder, file)

            # Decode the filename back into metadata
            # e.g. "EE_LTE_1800_RSRP_TICKET_HALL.png"
            #       [0]  [1] [2]  [3]   [4+]
            parts_f = file.replace(".png", "").split("_")
            if len(parts_f) < 5:
                continue  # Filename doesn't have enough parts — skip it

            tech    = parts_f[1]
            mhz     = parts_f[2]
            metric  = parts_f[3]
            building = "_".join(parts_f[4:])  # Everything from part 4 onwards is the building name

            key = (tech, mhz, metric)

            # Check if this combination has a corresponding template slide
            if key not in TEMPLATE_MAP:
                missing.append((key, file))  # Log it but don't crash
                continue

            provider_slides.append((provider, file, path, tech, mhz, metric, building))

    # Sort all slides into the correct order before adding them to the presentation
    sort_key = _make_sort_key(building_order or DEFAULT_BUILDING_ORDER)
    provider_slides.sort(key=sort_key)

    # Add a slide for each image, in the sorted order
    placed = 0
    for _, file, path, tech, mhz, metric, building in provider_slides:
        # Find the right blank template slide for this tech/freq/metric combo
        slide = duplicate_slide(prs, tmpl.slides[TEMPLATE_MAP[(tech, mhz, metric)]])
        # Fill in the image, building label, and logo
        update_slide(slide, path, building, logo_path)
        placed += 1

    # Remove the original blank template slides now that we've added all real ones
    delete_template_slides(prs, n_template_slides)

    # Fill in the slide-master metadata placeholders that appear on every content slide.
    # These are inherited from the master so we only need to replace them once here.
    #
    #   <<d>>   → initials of the author   (e.g. "Mark Young"  → "MY")
    #   <<c>>   → initials of the checker  (e.g. "Neil Adams"  → "NA")
    #   <<a>>   → initials of the approver (e.g. "Chris King"  → "CK")
    #   <<po1>> → original submission date (new dedicated input in the UI)
    #   <<po2>> → updated design date      (reuses the slide-1 date field)
    if metadata:
        master_replacements = {
            "<<d>>":   _initials(metadata.get("author",   "")),
            "<<c>>":   _initials(metadata.get("checked",  "")),
            "<<a>>":   _initials(metadata.get("approved", "")),
            "<<po1>>": metadata.get("p01_date", ""),
            "<<po2>>": metadata.get("date",     ""),
        }
        replace_in_master(prs, master_replacements)

    # Save the finished presentation to disk
    prs.save(output_path)

    return placed, missing


def run(config, progress_callback=None):
    """
    The main entry point — this is what the UI calls to kick everything off.

    'config' is a dictionary containing all the settings from the UI:
        pptx_path       — path to the input RF Predictions PowerPoint
        template_path   — path to the Appendix B template PowerPoint
        threed_path     — path to the 3D station model PNG image
        providers       — list of providers, each with a name and logo path
                          e.g. [{"name": "EE", "logo_path": "/path/to/EE.png"}, ...]
        output_dir      — folder where finished presentations will be saved
        work_dir        — temporary folder for intermediate files (extracted images)
        building_order  — ordered list of building types from the UI
                          e.g. ["Ticket Hall", "Midway", "Platform 1&2"]
        metadata        — dict of document metadata for slide 1's table (all optional)
                          e.g. {"date": "01/07/2025", "author": "JD",
                                "checked": "SM", "approved": "KL"}

    'progress_callback' is an optional function the UI provides so we can send
    status messages back (e.g. "Extracting images…", "Building EE presentation…").

    Returns a list of results, one per provider:
        [{"provider": "EE", "path": "/output/Appendix_B_EE.pptx",
          "placed": 42, "missing": []}, ...]
    """
    # Unpack all settings from the config dictionary
    pptx_path      = config["pptx_path"]
    template_path  = config["template_path"]
    threed_path    = config["threed_path"]
    providers      = config["providers"]
    output_dir     = config["output_dir"]
    work_dir       = config["work_dir"]
    # Use the user's building order from the UI, or fall back to the default
    building_order = config.get("building_order", DEFAULT_BUILDING_ORDER)

    # Metadata for the title slide table — all fields are optional
    # e.g. {"date": "01/07/2025", "author": "JD", "checked": "SM", "approved": "KL"}
    metadata = config.get("metadata", {})

    # Extract just the provider names (e.g. ["EE", "Vodafone", "Three"])
    provider_names = [p["name"] for p in providers]

    # Temporary folder where extracted PNG images will be stored
    extracted_folder = os.path.join(work_dir, "extracted_images")

    if progress_callback:
        progress_callback("Extracting RF images from input PPTX…")

    # Phase 1: Extract all map images from the input PowerPoint
    n = process_pptx(pptx_path, extracted_folder, provider_names, building_order)

    if progress_callback:
        progress_callback(f"Extracted {n} images. Loading template…")

    # Load the template PowerPoint once and keep it as a read-only reference.
    # We'll copy individual slides from this into each provider's output file.
    tmpl = Presentation(template_path)
    results = []

    # Phase 2: Build one output presentation per provider
    for p in providers:
        name      = p["name"]
        logo_path = p.get("logo_path", "")
        out_path  = os.path.join(output_dir, f"Appendix_B_{name}.pptx")

        if progress_callback:
            progress_callback(f"Building presentation for {name}…")

        # Check that we actually found some images for this provider before proceeding
        folder = os.path.join(extracted_folder, name)
        if not os.path.exists(folder) or not os.listdir(folder):
            if progress_callback:
                progress_callback(f"⚠ No images found for {name}, skipping.")
            continue

        # Build the full presentation for this provider
        placed, missing = build_provider_presentation(
            provider=name,
            logo_path=logo_path,
            template_pptx=template_path,
            tmpl=tmpl,
            output_folder=extracted_folder,
            threed_image_path=threed_path,
            output_path=out_path,
            building_order=building_order,
            metadata=metadata,
        )

        results.append({"provider": name, "path": out_path, "placed": placed, "missing": missing})

        if progress_callback:
            progress_callback(f"✓ {name}: {placed} slides placed.")

    return results
