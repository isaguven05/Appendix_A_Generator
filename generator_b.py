"""
Core PPTX generation logic for Appendix B automation.

Slide-building logic is identical to Appendix A (same layout names,
same image placement, always exactly 2 × 2_GA slides).

The ONLY difference is the metadata placeholder names and the
<P> provider-code replacement.
"""

import io

from pptx import Presentation

# Reuse all shared helpers from the existing A generator
from generator import (
    RF_IMAGE_BOUNDS,
    GA_IMAGE_BOUNDS,
    GA_TITLE_BOUNDS,
    sort_design_plans,
    extract_building_title,
    delete_slide,
    add_image_shape,
    add_text_box,
    get_initials,
    _post_process_pptx,
)

# Provider code mapping  (<P> placeholder in Appendix B templates)
PROVIDER_CODES = {
    "EE":       "2",
    "VODAFONE": "3",
    "VMO2":     "4",
    "THREE":    "5",
}


def generate_pptx_b(
    template_path: str,
    date: str,
    author: str,
    checked: str,
    approved: str,
    address: str,
    p01_date: str,
    design_plan_images: list,   # file paths sorted by plan number
    building_images: list,      # file paths
    building_titles: list = None,  # pre-extracted titles (preserves '&' etc.)
    provider: str = "",         # e.g. "EE", "THREE", "VODAFONE", "VMO2"
) -> bytes:
    """
    Generate one Appendix B PPTX for a single provider and return as bytes.

    Output slide order (identical to Appendix A):
      1  – Title slide
      2  – Design Change Summary
      3  – 2_GA slide  (always exactly 2)
      4  – 2_GA slide
      5… – RF Schematics slides (one per design plan image)
      …  – GA slides (one per building image)

    Metadata replacements (Appendix B–specific):
      <<DATE>>     → date
      <<AUTHOR>>   → author
      <<CHECKED>>  → checked
      <<APPROVED>> → approved
      <<ADDRESS>>  → address
      <<P01>>      → p01_date
      <<P02>>      → date  (same value as <<DATE>>)
      <D>          → initials of author
      <C>          → initials of checked
      <A>          → initials of approved
      <P>          → provider code (EE→2, VODAFONE→3, VMO2→4, THREE→5)
    """
    d_ini = get_initials(author)
    c_ini = get_initials(checked)
    a_ini = get_initials(approved)
    provider_code = PROVIDER_CODES.get(provider.upper(), "")

    design_plan_images = sort_design_plans(design_plan_images)

    # ── Load template ─────────────────────────────────────────────────────────
    prs = Presentation(template_path)

    # Build layout map – first occurrence wins (avoids duplicate-name overwrite)
    layout_map = {}
    for lay in prs.slide_layouts:
        if lay.name not in layout_map:
            layout_map[lay.name] = lay

    lay_2ga = layout_map["2_GA"]
    lay_rf  = layout_map["RF Schematics"]
    lay_ga  = layout_map["GA"]

    # Keep slides 0 & 1 (Title + Design Change Summary); delete the rest
    while len(prs.slides) > 2:
        delete_slide(prs, 2)

    # ── Always exactly 2 × 2_GA slides ───────────────────────────────────────
    for _ in range(2):
        prs.slides.add_slide(lay_2ga)

    # ── RF Schematics slides ──────────────────────────────────────────────────
    rf_slides = []
    for img_path in design_plan_images:
        slide = prs.slides.add_slide(lay_rf)
        rf_slides.append((slide, img_path))

    # ── GA (building) slides ──────────────────────────────────────────────────
    ga_slides = []
    for img_path in building_images:
        slide = prs.slides.add_slide(lay_ga)
        ga_slides.append((slide, img_path))

    # ── RF Schematics – insert design plan images ─────────────────────────────
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

    # ── XML post-processing: Appendix B metadata replacements ────────────────
    replacements = {
        "<<DATE>>":     date,
        "<<AUTHOR>>":   author,
        "<<CHECKED>>":  checked,
        "<<APPROVED>>": approved,
        "<<ADDRESS>>":  address,
        "<<P01>>":      p01_date,
        "<<P02>>":      date,      # <<P02>> mirrors <<DATE>> for Appendix B
        "<D>":          d_ini,
        "<C>":          c_ini,
        "<A>":          a_ini,
        "<P>":          provider_code,
    }

    pptx_bytes = _post_process_pptx(pptx_bytes, replacements)
    return pptx_bytes
