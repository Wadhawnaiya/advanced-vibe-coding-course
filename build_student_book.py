#!/usr/bin/env python3
"""Assembles Vibe_Coding_Student_Book.docx from the source-material docs."""
import os

from docxcompose.composer import Composer

from docx_merge import load_brand, new_base_document, render_cover, render_about_page, load_excerpt, apply_header_footer, _color

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
SOURCE_DIR = os.path.join(BASE_DIR, "source-material")
STUDY_BOOK = os.path.join(SOURCE_DIR, "Vibe_Coding_Student_Study_Book_v2.docx")
WORKBOOK = os.path.join(SOURCE_DIR, "Vibe_Coding_Student_Workbook.docx")
OUT = os.path.join(BASE_DIR, "Vibe_Coding_Student_Book.docx")

MAIN_SEQUENCE = [
    (WORKBOOK, "Week 0 — Setup Checklist", "Module 1: Digital Ground Zero — Files, Terminal, First Page"),
    (STUDY_BOOK, "Chapter 1 — What You Are Learning", "Chapter 3 — Claude Code Features (Student Encyclopedia)"),
    (WORKBOOK, "Module 1: Digital Ground Zero — Files, Terminal, First Page", "Module 2: Directing Claude — Prompt Craft & Plan Mode"),
    (STUDY_BOOK, "Chapter 3 — Claude Code Features (Student Encyclopedia)", "Chapter 6 — HTML Study Guide"),
    (WORKBOOK, "Module 2: Directing Claude — Prompt Craft & Plan Mode", "Module 3: How Web Apps Work (No Coding Torture)"),
    (STUDY_BOOK, "Chapter 6 — HTML Study Guide", "Chapter 9 — Python Study Guide"),
    (WORKBOOK, "Module 3: How Web Apps Work (No Coding Torture)", "Module 4: Interactive Frontend — Pages That Do Things"),
    (WORKBOOK, "Module 4: Interactive Frontend — Pages That Do Things", "Module 5: Think Before You Build — PRD & Scope"),
    (WORKBOOK, "Module 5: Think Before You Build — PRD & Scope", "Module 6: Data & Backend Basics"),
    (STUDY_BOOK, "Chapter 9 — Python Study Guide", "Chapter 10 — Git Study Guide"),
]


def build_main_sequence(doc, composer):
    for path, start, end in MAIN_SEQUENCE:
        composer.append(load_excerpt(path, start, end))


def add_tail_sections(doc, composer, brand):
    pass  # implemented in Task 11


def main():
    brand = load_brand()
    doc = new_base_document()
    render_cover(doc, brand, light=True)
    render_about_page(doc, brand, light=True)
    composer = Composer(doc)

    build_main_sequence(doc, composer)
    add_tail_sections(doc, composer, brand)

    apply_header_footer(doc, brand, "Student Book")
    doc.save(OUT)
    print("Wrote", OUT)


if __name__ == "__main__":
    main()
