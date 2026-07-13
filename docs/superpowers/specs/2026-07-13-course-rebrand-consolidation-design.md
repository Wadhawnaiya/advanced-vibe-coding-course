# Vibe Coding Course — Rebrand & Consolidation (Phase 1 Design)

## Context

The `vibe-coding` folder contains a "Vibe Coding with Claude Code" training course built by junior staff: 7 generator scripts (4 JS using the `docx`/`pptxgenjs` npm packages, 3 Python using `python-docx`) producing 6 deliverables (5 `.docx`, 1 `.pptx`). None of the current output carries any instructor branding, and content/color/versioning conventions are defined independently in each script, causing drift.

The owner (CA Shailesh S Wadhawaniya) wants the course rebranded as his own and consolidated down to the 3 documents he actually uses: one deep reference for himself to learn from and teach with, one self-contained book for students, and one slide deck to run in class. Content-quality deepening is explicitly deferred to a Phase 2, done module-by-module after this phase ships.

Branding facts (fetched from shaileshai.qzz.io, owner-confirmed context):
- Name: CA Shailesh S Wadhawaniya
- Credentials line: Chartered Accountant · AI Trainer · ESG Consultant · CPA(US) Educator
- Tagline: "Transforming businesses through the power of Generative AI, Data Analytics, and Sustainable ESG strategies."
- Roles: National ICAI Faculty (AICA Committee, DITS & WTO Committee), Principal AI & ESG Consultant, Chief Educator — StrideX CPA(US)
- Orgs: Wadhawaniya & Co. (CA firm), StrideX Institute / StrideX CPA(US)
- Contact: +91 79900 46540 · wadhawaniya@gmail.com · linkedin.com/in/ca-shailesh-wadhawaniya · Ahmedabad, Gujarat, India
- Credentials detail: CA (ICAI, Jul 2021), M.Com, B.Com, FAFD, AICA Level 1 & 2, Certified Social Auditor, AIF (IFSCA GIFT City), CPA(US) in progress, EA(US)

## Goal

Reduce 7 scripts / 6 outputs to 3 generator scripts / 3 outputs, fully rebranded under one consistent, print-friendly design system, with no content lost in the merge (only genuine duplicates collapsed).

## Final architecture

- `brand.json` — single source of truth read by both Python and Node: name, credentials line, tagline, org names, contact block, course title, version, date, and the print-first color palette. Any future detail change (phone, tagline, etc.) happens once here.
- `build_instructor_book.py` → **Vibe_Coding_Instructor_Book.docx**
  Merges: `generate-course.js` (Full Course Guide) + `build_master_guide.py` (Master Teaching Guide) + `build_volume2_workshops.py` (Deep Workshops)
- `build_student_book.py` → **Vibe_Coding_Student_Book.docx**
  Merges: `build_student_studybook.py` (Student Study Book) + `generate-workbook.js` (Student Workbook)
- `generate-slides.js` → **Vibe_Coding_Module_Slides.pptx**
  Rebrand of the existing `generate-slides-v2.js`; the original v1 `generate-slides.js` is retired.

All superseded scripts and their `.docx`/`.pptx` outputs move to `archive/` (not deleted).

## Branding placement (all 3 docs)

- Cover page: name + credentials line + tagline + "A StrideX Institute Program, in partnership with Wadhawaniya & Co."
- Header/footer every page: course title + name + page number (print-friendly palette, no color change from current per-doc accents beyond making them consistent)
- Dedicated "About the Instructor" page after the cover: bio paragraph, credentials list, roles, contact block
- Consistent course title, version, and date string across all 3 documents
- Student Book uses a lighter version: name + credentials line + "A StrideX Institute Program" only — no personal phone/email, since it's handed to a class

## Instructor Book — merge plan

1. Cover + About the Instructor
2. Part A — Course Design & Teaching Philosophy (from Full Course Guide: snapshot, schedule, gap analysis, learning outcomes, classroom setup)
3. Part B — Technical Encyclopedia (from Master Teaching Guide: terminal, Claude Code features, slash commands, CLAUDE.md/Skills/MCP, HTML/CSS/JS/Python/Git/GitHub) — the canonical lookup reference
4. Part C — Module-by-Module Teaching Flow: the 12 modules from the Full Course Guide, each immediately followed by its matching Deep Workshop (e.g., Module 3 + Workshop 3 HTML build) so the teaching script and the hands-on walkthrough sit together
5. Part D — Facilitation Playbook & Resources: session templates, assessment rubrics, prompt library, glossary, troubleshooting, FAQ, checklists, alternate course formats

Dedup rule: where content repeats verbatim in purpose (e.g., slash-command lists), Part B keeps the full encyclopedia entry as the canonical reference; Part C keeps only the short curated "essential set for this lesson" — these are reference vs. in-the-moment teaching list, not true duplicates, so both stay.

## Student Book — merge plan

Interleaved by week/chapter: each concept chapter from the Study Book is immediately followed by that week's fillable pages from the Workbook (lab, copy-paste prompts, homework, "my notes"). Ends with: Full-Stack Build Path, Deploy Study Guide, Quick Reference Cards, Capstone Tracker, Homework Log, Glossary (merged single copy), and the closing "You Can Build Software" section.

## Slides

`generate-slides-v2.js` becomes `generate-slides.js`, reads `brand.json` for the title slide, footer, and course name; output renamed to drop "v2". Original v1 script archived.

## Testing / verification

- Each new/edited script must run cleanly (`node --check` / `python -m py_compile`, then execute) and produce a valid, openable `.docx`/`.pptx`.
- Heading-outline diff: the merged document's heading list must account for every section in its source documents (minus only the explicitly-agreed dedup collapses) — no silent content loss.
- Grep generated document XML for the brand fields (name, phone, tagline) to confirm they render, not just that the script didn't crash.

## Explicitly out of scope for this phase

- Deepening/rewriting actual teaching content, examples, or exercises (Phase 2, module-by-module, starting with the Instructor Book, after this phase ships).
- Any visual theme beyond the existing print-friendly, per-document accent-color approach made consistent (no dark/gold digital-brand redesign).
