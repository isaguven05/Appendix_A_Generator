"""
Core PPTX generation logic for Appendix B.

Uses the building logic from engine_b.py (the original engine.py) without
any modification.  The extraction phase (reading from an input PPTX) is
skipped — the user uploads RF prediction map images directly, named in the
format produced by engine_b's extraction phase:

    PROVIDER_TECH_MHZ_METRIC_BUILDING.png
    e.g.  EE_LTE_1800_RSRP_TICKET_HALL.png

Provider codes for the <P> placeholder:
    EE → 2   VODAFONE → 3   VMO2 → 4   THREE → 5
"""

import io
import os
import shutil
import tempfile

from pptx import Presentation

# ── Import building helpers from engine_b (unmodified engine.py) ──────────────
from engine_b import (
    TEMPLATE_MAP,
    DEFAULT_BUILDING_ORDER,
    _make_sort_key,
    duplicate_slide,
    update_slide,
    add_intro_slides,
    delete_template_slides,
    replace_in_master,
    _initials,
)

# ── Import XML post-processing from existing generator.py ────────────────────
from generator import _post_process_pptx, extract_station_info

# Provider code mapping — inserted into <P> placeholder on every slide
PROVIDER_CODES = {
    "EE":       "2",
    "VODAFONE": "3",
    "VMO2":     "4",
    "THREE":    "5",
}


def generate_pptx_b(
    template_path: str,
    image_3d_path: str,
    date: str,
    p01_date: str,
    author: str,
    checked: str,
    approved: str,
    address: str,
    image_entries: list,      # [(saved_path, original_name), …] for ONE provider
    provider: str,
    building_order: list = None,
    logo_path: str = None,
) -> bytes:
    """
    Generate one Appendix B PPTX for a single provider and return as bytes.

    Slide structure (driven by templateB.pptx + engine_b):
      Slide 1  – Title slide (3D station image)
      Slide 2  – Executive summary
      Slide 3  – Operator logo slide
      Slides 4+ – RF prediction maps, one per uploaded image, sorted by
                   tech / frequency / metric and then by building order

    Metadata replacements applied:
      Via engine_b replace_in_master:
        <<d>>   → author initials
        <<c>>   → checked initials
        <<a>>   → approved initials
        <<po1>> → P01 date
        <<po2>> → date  (same value as <<DATE>>)

      Via XML post-processing (_post_process_pptx):
        <<DATE>>         → date
        <<AUTHOR>>       → author
        <<CHECKED>>      → checked
        <<APPROVED>>     → approved
        <<ADDRESS>>      → address
        <<P01>>          → p01_date
        <<P02>>          → date
        <<STATION_NAME>> → station code_name (from 3D image filename)
        <D>              → author initials
        <C>              → checked initials
        <A>              → approved initials
        <P>              → provider code (EE→2, VODAFONE→3, VMO2→4, THREE→5)
        <NUM>            → auto slide-number field (via _install_slidenum_field)
    """

    # ── Station info from 3D image filename (same method as Appendix A) ───────
    try:
        code, name, combined = extract_station_info(image_3d_path)
        station_display = f"{code} {name}"   # used in engine_b slide replacements
    except Exception:
        combined        = "UNKNOWN"
        station_display = "UNKNOWN"

    # ── Metadata dict used by engine_b helpers ────────────────────────────────
    metadata = {
        "date":     date,
        "author":   author,
        "checked":  checked,
        "approved": approved,
        "p01_date": p01_date,
    }

    # ── Parse image entries into provider_slides list ─────────────────────────
    # Filename format: PROVIDER_TECH_MHZ_METRIC_BUILDING...png
    # e.g.  EE_LTE_1800_RSRP_TICKET_HALL.png
    #       [0]  [1]  [2]  [3]    [4+]
    provider_slides = []

    for saved_path, orig_name in image_entries:
        stem    = os.path.splitext(orig_name)[0]        # strip extension
        parts_f = stem.replace(" ", "_").split("_")     # normalise then split

        # Need at least: PROVIDER TECH MHZ METRIC BUILDING (5 parts minimum)
        if len(parts_f) < 5:
            continue

        tech     = parts_f[1].upper()
        mhz      = parts_f[2]
        metric   = parts_f[3].upper()
        building = "_".join(parts_f[4:])   # building may be multi-word
        key      = (tech, mhz, metric)

        if key not in TEMPLATE_MAP:
            continue   # unrecognised tech/metric combo — skip

        provider_slides.append(
            (provider, orig_name, str(saved_path), tech, mhz, metric, building)
        )

    # ── Sort by tech/freq/metric then building order ──────────────────────────
    sort_key = _make_sort_key(building_order or DEFAULT_BUILDING_ORDER)
    provider_slides.sort(key=sort_key)

    # ── Build the presentation in a temp dir (engine_b uses file paths) ───────
    with tempfile.TemporaryDirectory() as tmpdir:
        out_path = os.path.join(tmpdir, "output.pptx")

        # engine_b's approach: copy template → build on top → delete originals
        shutil.copy(template_path, out_path)
        prs  = Presentation(out_path)
        tmpl = Presentation(template_path)  # read-only reference for duplication

        n_template_slides = len(prs.slides)

        # ── Intro slides: title, exec summary, operator logo ──────────────────
        add_intro_slides(
            prs, tmpl,
            threed_image_path = image_3d_path,
            provider          = provider,
            logo_path         = logo_path or "",
            metadata          = metadata,
        )

        # ── Content slides: one per RF map image ──────────────────────────────
        for _, _fname, path, tech, mhz, metric, building in provider_slides:
            tmpl_slide = tmpl.slides[TEMPLATE_MAP[(tech, mhz, metric)]]
            slide      = duplicate_slide(prs, tmpl_slide)
            update_slide(slide, path, building, logo_path or "")

        # ── Remove original blank template slides ──────────────────────────────
        delete_template_slides(prs, n_template_slides)

        # ── Master metadata (engine_b placeholders) ────────────────────────────
        replace_in_master(prs, {
            "<<d>>":   _initials(author),
            "<<c>>":   _initials(checked),
            "<<a>>":   _initials(approved),
            "<<po1>>": p01_date,
            "<<po2>>": date,
        })

        # ── Save to bytes ──────────────────────────────────────────────────────
        buf = io.BytesIO()
        prs.save(buf)
        pptx_bytes = buf.getvalue()

    # ── XML post-processing: all remaining placeholder replacements ───────────
    d_ini = _initials(author)
    c_ini = _initials(checked)
    a_ini = _initials(approved)

    replacements = {
        # Appendix B specific
        "<<DATE>>":         date,
        "<<AUTHOR>>":       author,
        "<<CHECKED>>":      checked,
        "<<APPROVED>>":     approved,
        "<<ADDRESS>>":      address,
        "<<P01>>":          p01_date,
        "<<P02>>":          date,
        # Station name (same method as Appendix A)
        "<<STATION_NAME>>": combined,
        # Initials (both formats)
        "<D>":              d_ini,
        "<C>":              c_ini,
        "<A>":              a_ini,
        "<<d>>":            d_ini,
        "<<c>>":            c_ini,
        "<<a>>":            a_ini,
        "<<po1>>":          p01_date,
        "<<po2>>":          date,
        # Provider code
        "<P>":              PROVIDER_CODES.get(provider.upper(), ""),
        # engine_b slide-level placeholders
        "STATION-NAME":     station_display,
        "OPERATOR-NAME":    provider,
    }

    pptx_bytes = _post_process_pptx(pptx_bytes, replacements)
    return pptx_bytes
