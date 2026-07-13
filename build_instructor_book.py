#!/usr/bin/env python3
"""Assembles Vibe_Coding_Instructor_Book.docx from the source-material docs."""
import os

from docx import Document
from docxcompose.composer import Composer

from docx_merge import load_brand, new_base_document, render_cover, render_about_page, load_excerpt, apply_header_footer

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
SOURCE_DIR = os.path.join(BASE_DIR, "source-material")
FULL_COURSE_GUIDE = os.path.join(SOURCE_DIR, "Vibe_Coding_Claude_Code_Full_Course_Guide.docx")
MASTER_GUIDE = os.path.join(SOURCE_DIR, "Vibe_Coding_Master_Teaching_Guide_v2.docx")
DEEP_WORKSHOPS = os.path.join(SOURCE_DIR, "Vibe_Coding_Deep_Workshops_v2.docx")
OUT = os.path.join(BASE_DIR, "Vibe_Coding_Instructor_Book.docx")


def add_part_a(doc, composer, brand):
    h = doc.add_heading("Part A — Course Design & Teaching Philosophy", level=1)
    for run in h.runs:
        from docx_merge import _color
        run.font.color.rgb = _color(brand["palette"]["primary"])
    composer.append(load_excerpt(
        FULL_COURSE_GUIDE,
        "1. Course Snapshot",
        "Part B — Pre-Course Onboarding (Week 0)",
    ))


def add_part_b(doc, composer, brand):
    h = doc.add_heading("Part B — Technical Encyclopedia", level=1)
    for run in h.runs:
        from docx_merge import _color
        run.font.color.rgb = _color(brand["palette"]["primary"])
    composer.append(load_excerpt(
        MASTER_GUIDE,
        "Part 1 — Mindset & What You Are Actually Learning",
    ))


def add_part_c(doc, composer, brand):
    pass  # implemented in Task 8


def add_part_d(doc, composer, brand):
    pass  # implemented in Task 9


def main():
    brand = load_brand()
    doc = new_base_document()
    render_cover(doc, brand, light=False)
    render_about_page(doc, brand, light=False)
    composer = Composer(doc)

    add_part_a(doc, composer, brand)
    add_part_b(doc, composer, brand)
    add_part_c(doc, composer, brand)
    add_part_d(doc, composer, brand)

    apply_header_footer(doc, brand, "Instructor Book")
    doc.save(OUT)
    print("Wrote", OUT)


if __name__ == "__main__":
    main()
