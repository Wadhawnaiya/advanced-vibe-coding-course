#!/usr/bin/env python3
"""Assembles Vibe_Coding_Instructor_Book.docx from the source-material docs."""
import os

from docxcompose.composer import Composer

from docx_merge import (
    load_brand,
    new_base_document,
    render_cover,
    render_about_page,
    load_excerpt,
    strip_closing_trailer,
    recolor_headings,
    apply_header_footer,
)

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
    excerpt = strip_closing_trailer(load_excerpt(
        MASTER_GUIDE,
        "Part 1 — Mindset & What You Are Actually Learning",
    ))
    composer.append(excerpt)


PART_C_SEQUENCE = [
    (FULL_COURSE_GUIDE, "Part B — Pre-Course Onboarding (Week 0)", "Module 1 — Digital Ground Zero: Computer, Files, Terminal"),
    (FULL_COURSE_GUIDE, "Module 1 — Digital Ground Zero: Computer, Files, Terminal", "Module 2 — The Vibe Coding Craft: Directing Claude Code"),
    (DEEP_WORKSHOPS, "Workshop 1 — Install Claude Code and Understand Every Step", "Workshop 2 — Slash Commands Lab (Do / Observe / Explain)"),
    (FULL_COURSE_GUIDE, "Module 2 — The Vibe Coding Craft: Directing Claude Code", "Module 3 — How the Internet & Web Apps Work (No Syntax Torture)"),
    (DEEP_WORKSHOPS, "Workshop 2 — Slash Commands Lab (Do / Observe / Explain)", "Workshop 3 — HTML Deep Build (Portfolio Page)"),
    (FULL_COURSE_GUIDE, "Module 3 — How the Internet & Web Apps Work (No Syntax Torture)", "Module 4 — Interactive Frontend: Making Pages Do Things"),
    (DEEP_WORKSHOPS, "Workshop 3 — HTML Deep Build (Portfolio Page)", "Workshop 4 — CSS Deep Build (Make It Professional)"),
    (FULL_COURSE_GUIDE, "Module 4 — Interactive Frontend: Making Pages Do Things", "Module 5 — Think Before You Build: PRD, Scope, Wireframes"),
    (DEEP_WORKSHOPS, "Workshop 4 — CSS Deep Build (Make It Professional)", "Workshop 5 — JavaScript Deep Build (Quiz App)"),
    (DEEP_WORKSHOPS, "Workshop 5 — JavaScript Deep Build (Quiz App)", "Workshop 6 — Python Backend Mini API"),
    (FULL_COURSE_GUIDE, "Module 5 — Think Before You Build: PRD, Scope, Wireframes", "Module 6 — Data & Backend Basics (Still Human Language)"),
    (FULL_COURSE_GUIDE, "Module 6 — Data & Backend Basics (Still Human Language)", "Module 7 — Full-Stack MVP with Claude Code"),
    (DEEP_WORKSHOPS, "Workshop 6 — Python Backend Mini API", "Workshop 7 — Persistence & CRUD Properly"),
    (DEEP_WORKSHOPS, "Workshop 7 — Persistence & CRUD Properly", "Workshop 8 — Git From Zero to Daily Habit"),
    (FULL_COURSE_GUIDE, "Module 7 — Full-Stack MVP with Claude Code", "Module 8 — Users, Auth, and App Polish"),
    (DEEP_WORKSHOPS, "Workshop 10 — Full-Stack Vertical Slice", "Workshop 11 — Auth & Ownership"),
    (FULL_COURSE_GUIDE, "Module 8 — Users, Auth, and App Polish", "Module 9 — Git & GitHub: Save Points for Your Work"),
    (DEEP_WORKSHOPS, "Workshop 11 — Auth & Ownership", "Workshop 12 — Deploy & Public Demo"),
    (FULL_COURSE_GUIDE, "Module 9 — Git & GitHub: Save Points for Your Work", "Module 10 — Deploy, Test, and Demo Day"),
    (DEEP_WORKSHOPS, "Workshop 8 — Git From Zero to Daily Habit", "Workshop 9 — GitHub Complete Hands-On"),
    (DEEP_WORKSHOPS, "Workshop 9 — GitHub Complete Hands-On", "Workshop 10 — Full-Stack Vertical Slice"),
    (FULL_COURSE_GUIDE, "Module 10 — Deploy, Test, and Demo Day", "Module 11 (Optional Advanced) — Skills, MCP, and Power Features"),
    (DEEP_WORKSHOPS, "Workshop 12 — Deploy & Public Demo", "Workshop 13 — Capstone Integration Week"),
    (FULL_COURSE_GUIDE, "Module 11 (Optional Advanced) — Skills, MCP, and Power Features", "Module 12 (Optional) — From Vibe Coding to Responsible Building"),
    (FULL_COURSE_GUIDE, "Module 12 (Optional) — From Vibe Coding to Responsible Building", "Part C — Instructor Facilitation Playbook"),
    (DEEP_WORKSHOPS, "Workshop 13 — Capstone Integration Week", "Appendix — Facilitator Notes for Workshops"),
]


def add_part_c(doc, composer, brand):
    from docx_merge import _color, demote_part_headings
    h = doc.add_heading("Part C — Module-by-Module Teaching Flow", level=1)
    for run in h.runs:
        run.font.color.rgb = _color(brand["palette"]["primary"])
    for path, start, end in PART_C_SEQUENCE:
        excerpt = load_excerpt(path, start, end)
        demote_part_headings(excerpt)  # only affects the Week-0 "Part B — ..." heading; no-op on Module/Workshop headings
        composer.append(excerpt)


def add_part_d(doc, composer, brand):
    from docx_merge import _color, demote_part_headings
    h = doc.add_heading("Part D — Facilitation Playbook & Resources", level=1)
    for run in h.runs:
        run.font.color.rgb = _color(brand["palette"]["primary"])

    workshop_appendix = strip_closing_trailer(
        load_excerpt(DEEP_WORKSHOPS, "Appendix — Facilitator Notes for Workshops")
    )
    composer.append(workshop_appendix)

    resources = strip_closing_trailer(
        load_excerpt(FULL_COURSE_GUIDE, "Part C — Instructor Facilitation Playbook")
    )
    demote_part_headings(resources)
    composer.append(resources)


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

    recolor_headings(doc, brand)
    apply_header_footer(doc, brand, "Instructor Book")
    doc.save(OUT)
    print("Wrote", OUT)


if __name__ == "__main__":
    main()
