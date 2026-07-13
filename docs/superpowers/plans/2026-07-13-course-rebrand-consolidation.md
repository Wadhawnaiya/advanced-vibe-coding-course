# Vibe Coding Course — Rebrand & Consolidation Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Reduce the "Vibe Coding with Claude Code" course from 7 generator scripts / 6 outputs to 3 branded, consolidated deliverables (Instructor Book, Student Book, Slides), fully attributed to CA Shailesh S Wadhawaniya.

**Architecture:** A single `brand.json` is the source of truth for all branding facts and the print-friendly color palette. Two new Python scripts (`build_instructor_book.py`, `build_student_book.py`) assemble their target `.docx` by excerpting heading-bounded ranges out of the *already-generated* source `.docx` files (moved to `source-material/`) and stitching them together with the `docxcompose` library — this reuses all existing, already-correct formatting instead of re-deriving ~250KB of paragraph-building code from the original JS/Python generators. The slide deck is rebranded in place. Superseded scripts and outputs move to `archive/legacy-scripts/`.

**Tech Stack:** Python 3 + `python-docx` 1.2.0 + `docxcompose` (in a local `.venv`, since the system Python is externally-managed) for the two `.docx` builders; Node.js + `pptxgenjs` (already in `node_modules`) for the slide deck, unchanged.

## Global Constraints

- Branding facts (exact values, from `docs/superpowers/specs/2026-07-13-course-rebrand-consolidation-design.md`):
  - Name: `CA Shailesh S Wadhawaniya`
  - Credentials line: `Chartered Accountant · AI Trainer · ESG Consultant · CPA(US) Educator`
  - Tagline: `Transforming businesses through the power of Generative AI, Data Analytics, and Sustainable ESG strategies.`
  - Provider line: `A StrideX Institute Program, in partnership with Wadhawaniya & Co.`
  - Contact: `+91 79900 46540` · `wadhawaniya@gmail.com` · `linkedin.com/in/ca-shailesh-wadhawaniya` · `Ahmedabad, Gujarat, India`
- Unified palette is the **green** family already used by 4 of the 5 source docs (`build_master_guide.py`, `build_student_studybook.py`, `build_volume2_workshops.py`, `generate-workbook.js`): primary `0B3D2E`, accent `1A7A5C`, dark `1C2833`, muted `5D6D7E`. The Full Course Guide's lone blue palette is the one being brought into line, not the other way around.
- Student Book branding is the "light" variant: name + credentials line + provider line only — no personal phone/email on student-facing material.
- No content is dropped except the explicitly-named dedup collapses in Task 8 (encyclopedia vs. curated list) — every other section from every source document must appear somewhere in the merged output. Verification steps in each task check this.
- Deepening/rewriting actual teaching content is explicitly out of scope for this plan (Phase 2, later).

---

### Task 1: Local build environment (venv, git, dependencies)

**Files:**
- Create: `.gitignore`
- Create: `requirements.txt`

**Interfaces:**
- Produces: a `.venv/` with `python-docx` and `docxcompose` importable, used by every later Python task via `.venv/bin/python`.

- [ ] **Step 1: Create `.gitignore` so the environment/build folders never get committed**

```
.venv/
node_modules/
__pycache__/
*.pyc
```

- [ ] **Step 2: Initialize git and commit the current state as a baseline, before any changes**

This is a one-time bootstrap of an untracked project — `git add -A` is intentional here (with `.gitignore` already in place to exclude `node_modules/`), not a pattern to repeat in later tasks, where files should be added by name.

Run:
```bash
git init
git add -A
git commit -m "chore: baseline snapshot before rebrand and consolidation"
```
Expected: a first commit succeeds; `git status` shows a clean tree.

- [ ] **Step 3: Create the venv and install dependencies**

Run:
```bash
python3 -m venv .venv
.venv/bin/pip install --quiet --upgrade pip
.venv/bin/pip install --quiet python-docx==1.2.0 docxcompose
```

- [ ] **Step 4: Freeze dependencies**

```bash
.venv/bin/pip freeze | grep -iE "python-docx|docxcompose|lxml" > requirements.txt
```

- [ ] **Step 5: Verify imports work**

Run: `.venv/bin/python -c "import docx, docxcompose; print('ok', docx.__version__)"`
Expected: `ok 1.2.0`

- [ ] **Step 6: Commit**

```bash
git add .gitignore requirements.txt
git commit -m "chore: add local venv setup for docx build tooling"
```

---

### Task 2: Reorganize source files (source-material/ and archive/legacy-scripts/)

The two new builder scripts read pre-rendered content out of the *existing* generated `.docx` files rather than regenerating it, so those five files must be relocated to a permanent, clearly-labeled location **before** the builder scripts are written to reference them — not deleted, since they remain a live build dependency.

**Files:**
- Create directories: `source-material/`, `archive/legacy-scripts/`
- Move: 5 generated `.docx` files → `source-material/`
- Move: 6 superseded scripts + old slide pptx → `archive/legacy-scripts/`

**Interfaces:**
- Produces: `source-material/Vibe_Coding_Claude_Code_Full_Course_Guide.docx`, `source-material/Vibe_Coding_Master_Teaching_Guide_v2.docx`, `source-material/Vibe_Coding_Deep_Workshops_v2.docx`, `source-material/Vibe_Coding_Student_Study_Book_v2.docx`, `source-material/Vibe_Coding_Student_Workbook.docx` — these five paths are consumed by Tasks 6–11.

- [ ] **Step 1: Create the two directories**

```bash
mkdir -p source-material archive/legacy-scripts
```

- [ ] **Step 2: Move the five source docx files (kept as active build input)**

```bash
git mv Vibe_Coding_Claude_Code_Full_Course_Guide.docx source-material/
git mv Vibe_Coding_Master_Teaching_Guide_v2.docx source-material/
git mv Vibe_Coding_Deep_Workshops_v2.docx source-material/
git mv Vibe_Coding_Student_Study_Book_v2.docx source-material/
git mv Vibe_Coding_Student_Workbook.docx source-material/
```

- [ ] **Step 3: Move the superseded generator scripts and the old v1 slide script (dead code, no live output)**

```bash
git mv generate-course.js archive/legacy-scripts/
git mv generate-workbook.js archive/legacy-scripts/
git mv build_master_guide.py archive/legacy-scripts/
git mv build_student_studybook.py archive/legacy-scripts/
git mv build_volume2_workshops.py archive/legacy-scripts/
```

- [ ] **Step 4: Verify nothing else in the root references the old paths**

Run: `grep -rn "Vibe_Coding_Claude_Code_Full_Course_Guide\|Vibe_Coding_Master_Teaching_Guide\|Vibe_Coding_Deep_Workshops\|Vibe_Coding_Student_Study_Book\|Vibe_Coding_Student_Workbook" --include="*.js" --include="*.py" .`
Expected: no matches (the only remaining references will be added fresh, with `source-material/` paths, in Tasks 6 and 10).

- [ ] **Step 5: Commit**

```bash
git add -A
git commit -m "refactor: move source docs to source-material/, retire superseded scripts to archive/legacy-scripts/"
```

---

### Task 3: `brand.json` — single source of truth

**Files:**
- Create: `brand.json`

**Interfaces:**
- Produces: a JSON file with top-level keys `name`, `credentials_line`, `tagline`, `bio`, `roles` (list), `credentials_detail` (list), `orgs` (dict with `primary`, `secondary`, `provider_line`), `contact` (dict with `phone`, `email`, `linkedin`, `location`), `course` (dict with `title`, `subtitle`, `version`, `date`), `palette` (dict with `primary`, `accent`, `dark`, `muted`, `border`, `light_bg`). Consumed by `docx_merge.py` (Task 4) and `generate-slides.js` (Task 12).

- [ ] **Step 1: Write `brand.json`**

```json
{
  "name": "CA Shailesh S Wadhawaniya",
  "credentials_line": "Chartered Accountant · AI Trainer · ESG Consultant · CPA(US) Educator",
  "tagline": "Transforming businesses through the power of Generative AI, Data Analytics, and Sustainable ESG strategies.",
  "bio": "I empower modern finance professionals by bridging the gap between Generative AI, Data Analytics, and Industry 4.0 technologies. Having trained 10,000+ professionals and business leaders in AI, I help enterprises automate operations and scale intelligently. As an ESG & Sustainability Leader, I guide organizations on environmental responsibilities globally, and as a National Level Faculty for ICAI, I train Chartered Accountants and business leaders across India in Generative AI, ESG, and Industry 4.0 practices.",
  "roles": [
    "National ICAI Faculty — AICA Committee, DITS & WTO Committee",
    "Principal AI & ESG Consultant",
    "Chief Educator — StrideX CPA(US)"
  ],
  "credentials_detail": [
    "CA (ICAI, July 2021)",
    "M.Com, Gujarat University",
    "B.Com, Gujarat University",
    "FAFD (Forensic Accounting & Fraud Detection), ICAI",
    "AICA Level 1 & 2 Certified, ICAI AI Committee",
    "Certified Social Auditor, ISAI",
    "AIF in IFSCA GIFT City",
    "CPA(US) — Regulation & TCP (in progress), AICPA",
    "EA(US) — All Three Parts, IRS Enrolled Agent"
  ],
  "orgs": {
    "primary": "StrideX Institute",
    "secondary": "Wadhawaniya & Co.",
    "provider_line": "A StrideX Institute Program, in partnership with Wadhawaniya & Co."
  },
  "contact": {
    "phone": "+91 79900 46540",
    "email": "wadhawaniya@gmail.com",
    "linkedin": "linkedin.com/in/ca-shailesh-wadhawaniya",
    "location": "Ahmedabad, Gujarat, India"
  },
  "course": {
    "title": "Vibe Coding with Claude Code",
    "subtitle": "From Zero Tech Knowledge to Building Full-Stack Applications",
    "version": "2.0",
    "date": "2026"
  },
  "palette": {
    "primary": "0B3D2E",
    "accent": "1A7A5C",
    "dark": "1C2833",
    "muted": "5D6D7E",
    "border": "D5D8DC",
    "light_bg": "E8F8F5"
  }
}
```

- [ ] **Step 2: Verify it parses and has the required keys**

Run:
```bash
.venv/bin/python -c "
import json
b = json.load(open('brand.json'))
required = ['name','credentials_line','tagline','bio','roles','credentials_detail','orgs','contact','course','palette']
missing = [k for k in required if k not in b]
assert not missing, missing
print('ok')
"
```
Expected: `ok`

- [ ] **Step 3: Commit**

```bash
git add brand.json
git commit -m "feat: add brand.json as single source of truth for course branding"
```

---

### Task 4: `docx_merge.py` — shared assembly toolkit

**Files:**
- Create: `docx_merge.py`
- Test: ad-hoc verification script run inline (no pytest in this project; see Step 6)

**Interfaces:**
- Consumes: `brand.json` (Task 3)
- Produces (used by Tasks 6, 7, 8, 9, 10, 11):
  - `load_brand() -> dict`
  - `load_excerpt(path: str, start_heading: str, end_heading: str | None = None) -> docx.Document` — opens `path` fresh, trims its body to the paragraph range `[start_heading, end_heading)` matched on exact stripped paragraph text, returns the trimmed `Document`. Raises `AssertionError` naming the missing heading if `start_heading` isn't found.
  - `demote_part_headings(doc) -> docx.Document` — turns every `Heading 1` paragraph matching `"Part X — Title"` into a `Heading 2` reading just `"Title"` (drops the `Part X — ` prefix), in place. Returns the same `doc`.
  - `new_base_document() -> docx.Document` — a fresh `Document()` with the shared 0.75"/0.85" margins.
  - `render_cover(doc, brand: dict, light: bool = False) -> docx.Document`
  - `render_about_page(doc, brand: dict, light: bool = False) -> docx.Document`
  - `apply_header_footer(doc, brand: dict, subtitle: str) -> docx.Document`

- [ ] **Step 1: Write `docx_merge.py`**

```python
"""Shared toolkit for assembling the branded, consolidated Vibe Coding documents."""
import json
import os
import re

from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
BRAND_PATH = os.path.join(BASE_DIR, "brand.json")

PART_LETTER_RE = re.compile(r"^Part [A-Z] — (.+)$")


def load_brand():
    with open(BRAND_PATH, encoding="utf-8") as f:
        return json.load(f)


def _color(hexstr):
    return RGBColor(int(hexstr[0:2], 16), int(hexstr[2:4], 16), int(hexstr[4:6], 16))


def _set_run(run, size=11, bold=False, italic=False, color=None, font="Calibri"):
    run.font.name = font
    run._element.rPr.rFonts.set(qn("w:eastAsia"), font)
    run.font.size = Pt(size)
    run.bold = bold
    run.italic = italic
    if color is not None:
        run.font.color.rgb = _color(color)


def load_excerpt(path, start_heading, end_heading=None):
    """Open `path` fresh and trim its body to the paragraph range
    [start_heading, end_heading). Matching is exact on stripped paragraph
    text. Raises AssertionError if start_heading is not found."""
    doc = Document(path)
    body = doc.element.body
    start_el = None
    end_el = None
    for para in doc.paragraphs:
        text = para.text.strip()
        if start_el is None and text == start_heading:
            start_el = para._p
            continue
        if start_el is not None and end_heading and text == end_heading:
            end_el = para._p
            break
    assert start_el is not None, f"heading not found in {path}: {start_heading!r}"
    children = list(body)
    start_idx = children.index(start_el)
    end_idx = children.index(end_el) if end_el is not None else len(children) - 1
    keep_ids = {id(c) for c in children[start_idx:end_idx]}
    for child in children:
        if id(child) not in keep_ids:
            body.remove(child)
    return doc


def demote_part_headings(doc):
    """Turn source 'Part X — Title' H1s into H2s with the 'Part X — ' prefix
    stripped, so they nest under one umbrella H1 in the merged document."""
    for para in doc.paragraphs:
        if para.style.name != "Heading 1":
            continue
        m = PART_LETTER_RE.match(para.text.strip())
        if not m:
            continue
        new_text = m.group(1)
        for i, run in enumerate(para.runs):
            run.text = new_text if i == 0 else ""
        para.style = doc.styles["Heading 2"]
    return doc


def new_base_document():
    doc = Document()
    section = doc.sections[0]
    section.top_margin = Inches(0.75)
    section.bottom_margin = Inches(0.75)
    section.left_margin = Inches(0.85)
    section.right_margin = Inches(0.85)
    return doc


def render_cover(doc, brand, light=False):
    palette = brand["palette"]
    for _ in range(3):
        doc.add_paragraph()
    t = doc.add_paragraph(); t.alignment = WD_ALIGN_PARAGRAPH.CENTER
    _set_run(t.add_run(brand["course"]["title"].upper()), size=26, bold=True, color=palette["primary"])
    t = doc.add_paragraph(); t.alignment = WD_ALIGN_PARAGRAPH.CENTER
    _set_run(t.add_run(brand["course"]["subtitle"]), size=13, italic=True, color=palette["muted"])
    doc.add_paragraph()
    t = doc.add_paragraph(); t.alignment = WD_ALIGN_PARAGRAPH.CENTER
    _set_run(t.add_run(brand["name"]), size=18, bold=True, color=palette["dark"])
    t = doc.add_paragraph(); t.alignment = WD_ALIGN_PARAGRAPH.CENTER
    _set_run(t.add_run(brand["credentials_line"]), size=11, color=palette["muted"])
    doc.add_paragraph()
    t = doc.add_paragraph(); t.alignment = WD_ALIGN_PARAGRAPH.CENTER
    _set_run(t.add_run(brand["tagline"]), size=11, italic=True, color=palette["accent"])
    doc.add_paragraph()
    t = doc.add_paragraph(); t.alignment = WD_ALIGN_PARAGRAPH.CENTER
    _set_run(t.add_run(brand["orgs"]["provider_line"]), size=10, color=palette["muted"])
    if not light:
        t = doc.add_paragraph(); t.alignment = WD_ALIGN_PARAGRAPH.CENTER
        contact = brand["contact"]
        _set_run(
            t.add_run(f"{contact['phone']}  ·  {contact['email']}  ·  {contact['linkedin']}"),
            size=9, color=palette["muted"],
        )
    doc.add_paragraph()
    t = doc.add_paragraph(); t.alignment = WD_ALIGN_PARAGRAPH.CENTER
    _set_run(t.add_run(f"Version {brand['course']['version']} · {brand['course']['date']}"), size=9, color=palette["muted"])
    doc.add_page_break()
    return doc


def render_about_page(doc, brand, light=False):
    palette = brand["palette"]
    h = doc.add_heading("About the Instructor" if not light else "Course Provider", level=1)
    for run in h.runs:
        run.font.color.rgb = _color(palette["primary"])
    p1 = doc.add_paragraph()
    _set_run(p1.add_run(brand["name"]), size=14, bold=True, color=palette["dark"])
    p2 = doc.add_paragraph()
    _set_run(p2.add_run(brand["credentials_line"]), size=11, color=palette["muted"])
    doc.add_paragraph()
    body = doc.add_paragraph()
    _set_run(body.add_run(brand["bio"] if not light else brand["tagline"]), size=11)
    if not light:
        h2 = doc.add_heading("Roles", level=2)
        for run in h2.runs:
            run.font.color.rgb = _color(palette["accent"])
        for role in brand["roles"]:
            b = doc.add_paragraph(style="List Bullet")
            _set_run(b.add_run(role), size=11)
        h3 = doc.add_heading("Credentials", level=2)
        for run in h3.runs:
            run.font.color.rgb = _color(palette["accent"])
        for cred in brand["credentials_detail"]:
            b = doc.add_paragraph(style="List Bullet")
            _set_run(b.add_run(cred), size=11)
        h4 = doc.add_heading("Contact", level=2)
        for run in h4.runs:
            run.font.color.rgb = _color(palette["accent"])
        contact = brand["contact"]
        for label, value in [
            ("Phone", contact["phone"]),
            ("Email", contact["email"]),
            ("LinkedIn", contact["linkedin"]),
            ("Location", contact["location"]),
        ]:
            b = doc.add_paragraph()
            _set_run(b.add_run(f"{label}: "), size=11, bold=True)
            _set_run(b.add_run(value), size=11)
    doc.add_page_break()
    return doc


def apply_header_footer(doc, brand, subtitle):
    palette = brand["palette"]
    section = doc.sections[0]
    header_p = section.header.paragraphs[0]
    header_p.text = ""
    _set_run(header_p.add_run(f"{brand['course']['title']} — {subtitle}"), size=9, color=palette["muted"])
    footer_p = section.footer.paragraphs[0]
    footer_p.text = ""
    _set_run(footer_p.add_run(f"{brand['name']} · {brand['orgs']['primary']}"), size=9, color=palette["muted"])
    return doc
```

- [ ] **Step 2: Write a throwaway self-check script to verify `load_excerpt` and `demote_part_headings` against a real source file**

Create `/tmp/check_docx_merge.py` (not committed — scratch verification only):

```python
import sys
sys.path.insert(0, "/home/shailesh/Downloads/Training/vibe-coding")
from docx_merge import load_excerpt, demote_part_headings, load_brand

brand = load_brand()
assert brand["name"] == "CA Shailesh S Wadhawaniya"

excerpt = load_excerpt(
    "/home/shailesh/Downloads/Training/vibe-coding/source-material/Vibe_Coding_Master_Teaching_Guide_v2.docx",
    "Part 5 — Claude Code Features Encyclopedia",
    "Part 6 — Slash Commands Encyclopedia",
)
texts = [p.text.strip() for p in excerpt.paragraphs]
assert "5.1 Interactive session (REPL)" in texts, "expected content missing from excerpt"
assert "1.1 What “vibe coding” means (precisely)" not in texts, "excerpt leaked earlier content"
assert "Part 6 — Slash Commands Encyclopedia" not in texts, "excerpt leaked later content"

demoted = load_excerpt(
    "/home/shailesh/Downloads/Training/vibe-coding/source-material/Vibe_Coding_Claude_Code_Full_Course_Guide.docx",
    "Part C — Instructor Facilitation Playbook",
)
demote_part_headings(demoted)
demoted_texts = [(p.style.name, p.text.strip()) for p in demoted.paragraphs]
assert ("Heading 2", "Instructor Facilitation Playbook") in demoted_texts, demoted_texts[:3]
assert ("Heading 1", "Part C — Instructor Facilitation Playbook") not in demoted_texts

print("ok")
```

- [ ] **Step 3: Run it and confirm it fails before `docx_merge.py` exists incorrectly (sanity: run once now that the module is written — this validates the module itself, standing in for a unit test since this project has no pytest suite)**

Run: `.venv/bin/python /tmp/check_docx_merge.py`
Expected: `ok`

If it raises an `AssertionError`, the heading text in the script doesn't exactly match the source `.docx` — open the named source file's headings with `verify_docx.py` (Task 5) to get the exact text and fix the mismatch.

- [ ] **Step 4: Commit**

```bash
git add docx_merge.py
git commit -m "feat: add docx_merge toolkit for excerpting and branding merged documents"
```

---

### Task 5: `verify_docx.py` — output verification helper

**Files:**
- Create: `verify_docx.py`

**Interfaces:**
- Produces: `heading_list(path, levels=(1,2)) -> list[str]`, `all_text(path) -> str`, `assert_contains(path, needles: list[str]) -> None` (raises `AssertionError` listing every missing needle). Used by Tasks 9, 11, 12, 14 to confirm each build's output is complete and branded.

- [ ] **Step 1: Write `verify_docx.py`**

```python
"""Verification helpers for the generated docx/pptx deliverables."""
import sys
from docx import Document


def heading_list(path, levels=(1, 2)):
    doc = Document(path)
    names = {1: "Heading 1", 2: "Heading 2", 3: "Heading 3"}
    wanted = {names[level] for level in levels}
    return [p.text.strip() for p in doc.paragraphs if p.style.name in wanted and p.text.strip()]


def all_text(path):
    doc = Document(path)
    return "\n".join(p.text for p in doc.paragraphs)


def assert_contains(path, needles):
    text = all_text(path)
    missing = [n for n in needles if n not in text]
    assert not missing, f"{path} is missing expected text: {missing}"


if __name__ == "__main__":
    for h in heading_list(sys.argv[1]):
        print(h)
```

- [ ] **Step 2: Run it against an existing source doc to confirm it works**

Run: `.venv/bin/python verify_docx.py source-material/Vibe_Coding_Deep_Workshops_v2.docx | head -5`
Expected: prints heading lines starting with `Workshop 1 — Install Claude Code and Understand Every Step`

- [ ] **Step 3: Commit**

```bash
git add verify_docx.py
git commit -m "feat: add verify_docx heading/text-presence checks for generated deliverables"
```

---

### Task 6: `build_instructor_book.py` — skeleton, cover, about page, Part A

**Files:**
- Create: `build_instructor_book.py`

**Interfaces:**
- Consumes: `docx_merge.load_brand`, `.new_base_document`, `.render_cover`, `.render_about_page`, `.load_excerpt` (Task 4)
- Produces: `Vibe_Coding_Instructor_Book.docx` (partial — Part A only, at this stage). Functions `add_part_a(doc, composer, brand)`, `add_part_b(doc, composer, brand)` (stub), `add_part_c(doc, composer, brand)` (stub), `add_part_d(doc, composer, brand)` (stub) — later tasks fill in the stubs by exact string replacement.

- [ ] **Step 1: Write `build_instructor_book.py`**

```python
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
    pass  # implemented in Task 7


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
```

- [ ] **Step 2: Run it**

Run: `.venv/bin/python build_instructor_book.py`
Expected: `Wrote /home/shailesh/Downloads/Training/vibe-coding/Vibe_Coding_Instructor_Book.docx`

- [ ] **Step 3: Verify the cover, about page, and Part A content are present**

Run:
```bash
.venv/bin/python -c "
from verify_docx import assert_contains, heading_list
OUT = 'Vibe_Coding_Instructor_Book.docx'
assert_contains(OUT, [
    'CA Shailesh S Wadhawaniya',
    'Chartered Accountant',
    'wadhawaniya@gmail.com',
    'Part A — Course Design & Teaching Philosophy',
    'Recommended Schedule',
])
print('ok:', heading_list(OUT)[:6])
"
```
Expected: `ok: [...]` including `'About the Instructor'` and `'Part A — Course Design & Teaching Philosophy'`

- [ ] **Step 4: Commit**

```bash
git add build_instructor_book.py
git commit -m "feat: assemble Instructor Book cover, about page, and Part A"
```

---

### Task 7: Instructor Book — Part B (Technical Encyclopedia)

**Files:**
- Modify: `build_instructor_book.py`

**Interfaces:**
- Consumes: `MASTER_GUIDE`, `load_excerpt` (already imported)
- Produces: `add_part_b` fully implemented; no new interface for later tasks.

- [ ] **Step 1: Replace the `add_part_b` stub**

```python
def add_part_b(doc, composer, brand):
    h = doc.add_heading("Part B — Technical Encyclopedia", level=1)
    for run in h.runs:
        from docx_merge import _color
        run.font.color.rgb = _color(brand["palette"]["primary"])
    composer.append(load_excerpt(
        MASTER_GUIDE,
        "Part 1 — Mindset & What You Are Actually Learning",
    ))
```

- [ ] **Step 2: Run it**

Run: `.venv/bin/python build_instructor_book.py`
Expected: `Wrote .../Vibe_Coding_Instructor_Book.docx`

- [ ] **Step 3: Verify Part B content landed, in addition to Part A**

Run:
```bash
.venv/bin/python -c "
from verify_docx import assert_contains
OUT = 'Vibe_Coding_Instructor_Book.docx'
assert_contains(OUT, [
    'Part B — Technical Encyclopedia',
    'Part 1 — Mindset & What You Are Actually Learning',
    'Part 9 — How the Web Works + Full-Stack Mental Model',
    'Appendix H — Capstone PRD Template (Full Text)',
    'Part A — Course Design & Teaching Philosophy',
])
print('ok')
"
```
Expected: `ok`

- [ ] **Step 4: Commit**

```bash
git add build_instructor_book.py
git commit -m "feat: append Instructor Book Part B (Technical Encyclopedia)"
```

---

### Task 8: Instructor Book — Part C (Module-by-Module Teaching Flow, interleaved with Workshops)

**Files:**
- Modify: `build_instructor_book.py`

**Interfaces:**
- Consumes: `FULL_COURSE_GUIDE`, `DEEP_WORKSHOPS`, `load_excerpt`
- Produces: `add_part_c` fully implemented.

Dedup note: the slash-command lists inside Modules 1–2 stay as-is (they're the short curated "essential set for this lesson," distinct from Part B's full encyclopedia) — nothing removed here.

- [ ] **Step 1: Replace the `add_part_c` stub**

```python
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
```

- [ ] **Step 2: Run it**

Run: `.venv/bin/python build_instructor_book.py`
Expected: `Wrote .../Vibe_Coding_Instructor_Book.docx`

- [ ] **Step 3: Verify Part C's full module/workshop sequence and correct ordering**

Run:
```bash
.venv/bin/python -c "
from verify_docx import heading_list
h = heading_list('Vibe_Coding_Instructor_Book.docx')
i_m7 = h.index('Module 7 — Full-Stack MVP with Claude Code')
i_w10 = h.index('Workshop 10 — Full-Stack Vertical Slice')
i_m8 = h.index('Module 8 — Users, Auth, and App Polish')
assert i_m7 < i_w10 < i_m8, (i_m7, i_w10, i_m8)
assert 'Workshop 13 — Capstone Integration Week' in h
assert 'Module 12 (Optional) — From Vibe Coding to Responsible Building' in h
assert 'Pre-Course Onboarding (Week 0)' in h  # demoted heading, prefix stripped
print('ok, total headings:', len(h))
"
```
Expected: `ok, total headings: <some number greater than 90>`

- [ ] **Step 4: Commit**

```bash
git add build_instructor_book.py
git commit -m "feat: append Instructor Book Part C interleaving modules with matching workshops"
```

---

### Task 9: Instructor Book — Part D (Facilitation Playbook & Resources), header/footer, final save + verify

**Files:**
- Modify: `build_instructor_book.py`

**Interfaces:**
- Consumes: `FULL_COURSE_GUIDE`, `DEEP_WORKSHOPS`, `demote_part_headings`, `apply_header_footer`
- Produces: final `Vibe_Coding_Instructor_Book.docx`, fully assembled.

- [ ] **Step 1: Replace the `add_part_d` stub**

```python
def add_part_d(doc, composer, brand):
    from docx_merge import _color, demote_part_headings
    h = doc.add_heading("Part D — Facilitation Playbook & Resources", level=1)
    for run in h.runs:
        run.font.color.rgb = _color(brand["palette"]["primary"])

    workshop_appendix = load_excerpt(DEEP_WORKSHOPS, "Appendix — Facilitator Notes for Workshops")
    composer.append(workshop_appendix)

    resources = load_excerpt(FULL_COURSE_GUIDE, "Part C — Instructor Facilitation Playbook")
    demote_part_headings(resources)
    composer.append(resources)
```

- [ ] **Step 2: Run it**

Run: `.venv/bin/python build_instructor_book.py`
Expected: `Wrote .../Vibe_Coding_Instructor_Book.docx`

- [ ] **Step 3: Full verification pass — headings, branding, and no accidental content loss**

Run:
```bash
.venv/bin/python -c "
from verify_docx import assert_contains, heading_list
OUT = 'Vibe_Coding_Instructor_Book.docx'
assert_contains(OUT, [
    'CA Shailesh S Wadhawaniya',
    'A StrideX Institute Program',
    'Part A — Course Design & Teaching Philosophy',
    'Part B — Technical Encyclopedia',
    'Part C — Module-by-Module Teaching Flow',
    'Part D — Facilitation Playbook & Resources',
    'Facilitator Notes for Workshops',
    'Instructor Facilitation Playbook',
    'Ready-to-Teach Prompt Library',
    'Capstone Project Brief',
    'Beginner Glossary',
    'Troubleshooting Field Guide',
    'Weekly Homework Sheets',
    'Closing Note to the Instructor',
])
h = heading_list(OUT)
assert len(h) > 150, f'suspiciously few headings: {len(h)}'
print('ok, total headings:', len(h))
"
```
Expected: `ok, total headings: <a number > 150>`

- [ ] **Step 4: Spot-check the header/footer text (python-docx section object, not plain text search)**

Run:
```bash
.venv/bin/python -c "
from docx import Document
d = Document('Vibe_Coding_Instructor_Book.docx')
sec = d.sections[0]
print('header:', sec.header.paragraphs[0].text)
print('footer:', sec.footer.paragraphs[0].text)
assert 'Instructor Book' in sec.header.paragraphs[0].text
assert 'CA Shailesh S Wadhawaniya' in sec.footer.paragraphs[0].text
print('ok')
"
```
Expected: `ok`

- [ ] **Step 5: Commit**

```bash
git add build_instructor_book.py Vibe_Coding_Instructor_Book.docx
git commit -m "feat: complete Instructor Book Part D, header/footer, and final assembly"
```

---

### Task 10: `build_student_book.py` — skeleton, light cover/about, Week 0, interleaved chapters

**Files:**
- Create: `build_student_book.py`

**Interfaces:**
- Consumes: `docx_merge` module (Task 4)
- Produces: partial `Vibe_Coding_Student_Book.docx` (through Chapter 9/Module 6); `add_tail_sections(doc, composer, brand)` stub for Task 11.

- [ ] **Step 1: Write `build_student_book.py`**

```python
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
```

- [ ] **Step 2: Run it**

Run: `.venv/bin/python build_student_book.py`
Expected: `Wrote .../Vibe_Coding_Student_Book.docx`

- [ ] **Step 3: Verify the light branding and the interleaved sequence so far**

Run:
```bash
.venv/bin/python -c "
from verify_docx import assert_contains, heading_list
OUT = 'Vibe_Coding_Student_Book.docx'
assert_contains(OUT, [
    'CA Shailesh S Wadhawaniya',
    'A StrideX Institute Program',
    'Week 0 — Setup Checklist',
    'Chapter 1 — What You Are Learning',
    'Module 1: Digital Ground Zero',
    'Chapter 6 — HTML Study Guide',
    'Module 5: Think Before You Build',
])
assert 'wadhawaniya@gmail.com' not in open(OUT, 'rb').read().decode('latin1'), 'student book must not leak personal contact info'
h = heading_list(OUT)
assert 'Chapter 10 — Git Study Guide' not in h, 'Chapter 10 is the exclusive end-boundary of the last MAIN_SEQUENCE entry and must not appear until Task 11 adds the entry that starts there'
i_ch9 = h.index('Chapter 9 — Python Study Guide')
assert i_ch9 == len(h) - 5
print('ok, headings so far:', len(h))
"
```
Expected: `ok, headings so far: <some number>` (note: the personal-info check on raw bytes is a coarse guard, not a full text-decoding check — Step 3 of Task 11 does the authoritative `all_text` check once the document is complete. Also note: `load_excerpt`'s end boundary is exclusive, so "Chapter 10 — Git Study Guide" — the last entry's end heading — is correctly absent here; it only appears once Task 11's `MAIN_SEQUENCE` extension adds the entry that starts there.)

- [ ] **Step 4: Commit**

```bash
git add build_student_book.py
git commit -m "feat: assemble Student Book light cover/about, Week 0, and first half of interleaved chapters"
```

---

### Task 11: Student Book — remaining chapters, tail sections, header/footer, final verify

**Files:**
- Modify: `build_student_book.py`

**Interfaces:**
- Consumes: `STUDY_BOOK`, `WORKBOOK`, `load_excerpt`, `apply_header_footer`
- Produces: final `Vibe_Coding_Student_Book.docx`, fully assembled.

- [ ] **Step 1: Extend `MAIN_SEQUENCE` with the remaining modules/chapters**

Replace the `MAIN_SEQUENCE = [ ... ]` list literal written in Task 10 with this complete 19-entry version (every tuple is `(path, start_heading, end_heading)`; `end_heading` is `None` for the one entry that runs to the end of its source):

```python
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
    (WORKBOOK, "Module 6: Data & Backend Basics", "Module 7: Full-Stack MVP with Claude Code"),
    (STUDY_BOOK, "Chapter 12 — Full-Stack Build Path (Your Labs)", "Chapter 13 — Deploy Study Guide"),
    (WORKBOOK, "Module 7: Full-Stack MVP with Claude Code", "Module 8: Users, Login & Polish"),
    (WORKBOOK, "Module 8: Users, Login & Polish", "Module 9: Git & GitHub — Save Points"),
    (STUDY_BOOK, "Chapter 10 — Git Study Guide", "Chapter 12 — Full-Stack Build Path (Your Labs)"),
    (WORKBOOK, "Module 9: Git & GitHub — Save Points", "Module 10: Deploy, Test & Demo Day"),
    (STUDY_BOOK, "Chapter 13 — Deploy Study Guide", "Chapter 14 — Quick Reference Cards"),
    (WORKBOOK, "Module 10: Deploy, Test & Demo Day", "Tools You Will Use Every Day"),
    (STUDY_BOOK, "Chapter 14 — Quick Reference Cards", None),
]
```

Edit the original list in place (don't append at runtime) so the file has one clean 19-entry literal.

- [ ] **Step 2: Replace the `add_tail_sections` stub**

```python
def add_tail_sections(doc, composer, brand):
    for start, end in [
        ("Tools You Will Use Every Day", "Prompt Pocket Library"),
        ("Prompt Pocket Library", "Beginner Glossary"),
        ("Beginner Glossary", "Capstone Tracker"),
        ("Capstone Tracker", "Homework Log"),
        ("Homework Log", "You Can Build Software"),
        ("You Can Build Software", None),
    ]:
        composer.append(load_excerpt(WORKBOOK, start, end))
```

- [ ] **Step 3: Run it**

Run: `.venv/bin/python build_student_book.py`
Expected: `Wrote .../Vibe_Coding_Student_Book.docx`

- [ ] **Step 4: Full verification pass**

Run:
```bash
.venv/bin/python -c "
from verify_docx import assert_contains, all_text, heading_list
OUT = 'Vibe_Coding_Student_Book.docx'
assert_contains(OUT, [
    'CA Shailesh S Wadhawaniya',
    'A StrideX Institute Program',
    'Chapter 14 — Quick Reference Cards',
    'Module 10: Deploy, Test & Demo Day',
    'Tools You Will Use Every Day',
    'Prompt Pocket Library',
    'Beginner Glossary',
    'Capstone Tracker',
    'Homework Log',
    'You Can Build Software',
])
text = all_text(OUT)
assert 'wadhawaniya@gmail.com' not in text, 'student book leaked the instructor personal email'
assert '+91 79900 46540' not in text, 'student book leaked the instructor phone number'
h = heading_list(OUT)
assert len(h) > 60, f'suspiciously few headings: {len(h)}'
print('ok, total headings:', len(h))
"
```
Expected: `ok, total headings: <a number > 60>`

- [ ] **Step 5: Spot-check header/footer**

Run:
```bash
.venv/bin/python -c "
from docx import Document
d = Document('Vibe_Coding_Student_Book.docx')
sec = d.sections[0]
assert 'Student Book' in sec.header.paragraphs[0].text
assert 'CA Shailesh S Wadhawaniya' in sec.footer.paragraphs[0].text
print('ok')
"
```
Expected: `ok`

- [ ] **Step 6: Commit**

```bash
git add build_student_book.py Vibe_Coding_Student_Book.docx
git commit -m "feat: complete Student Book remaining chapters, tail sections, and header/footer"
```

---

### Task 12: Rebrand the slide deck

**Files:**
- Create: `generate-slides.js` (from `generate-slides-v2.js`, edited) — the root no longer has an old `generate-slides.js`; it was archived as `archive/legacy-scripts/generate-slides-v1.js` in a follow-up fix after Task 2 (Task 2's own git mv list missed it)
- Move: `generate-slides-v2.js` → `archive/legacy-scripts/` (after the new file is created and works)

**Interfaces:**
- Consumes: `brand.json` (Task 3)
- Produces: `Vibe_Coding_Module_Slides.pptx`

- [ ] **Step 1: Copy `generate-slides-v2.js` to `generate-slides.js`**

```bash
cp generate-slides-v2.js generate-slides.js
```

- [ ] **Step 2: Edit `generate-slides.js` to load `brand.json` and use it**

Modify the top of the file (after the existing `const path = require("path");` line):

```javascript
const fs = require("fs");
const brand = JSON.parse(fs.readFileSync(path.join(__dirname, "brand.json"), "utf8"));
```

- [ ] **Step 3: Update the presentation metadata**

Find:
```javascript
  pres.author = "Vibe Coding Course";
  pres.title = "Vibe Coding with Claude Code — Edition 2 Slides";
```
Replace with:
```javascript
  pres.author = brand.name;
  pres.title = `${brand.course.title} — Module Slides`;
```

- [ ] **Step 4: Update the shared footer to carry the brand name**

Find:
```javascript
  const footer = (s) => {
    s.addText("Vibe Coding · Claude Code · Edition 2", {
      x: 0.45, y: 5.28, w: 9.1, h: 0.22,
      fontSize: 10, fontFace: "Arial", color: P.muted, margin: 0,
    });
  };
```
Replace with:
```javascript
  const footer = (s) => {
    s.addText(`${brand.course.title} · ${brand.name}`, {
      x: 0.45, y: 5.28, w: 9.1, h: 0.22,
      fontSize: 10, fontFace: "Arial", color: P.muted, margin: 0,
    });
  };
```

- [ ] **Step 5: Add instructor attribution to the title slide**

Find the closing of the `// ===== TITLE =====` block:
```javascript
    s.addText("Features · Slash commands · HTML/CSS/JS/Python · Git · GitHub · Full-stack", {
      x: 0.55, y: 5.12, w: 9, h: 0.35,
      fontSize: 13, color: P.cream, fontFace: "Arial", margin: 0,
    });
  }
```
Replace with:
```javascript
    s.addText("Features · Slash commands · HTML/CSS/JS/Python · Git · GitHub · Full-stack", {
      x: 0.55, y: 5.12, w: 9, h: 0.35,
      fontSize: 13, color: P.cream, fontFace: "Arial", margin: 0,
    });
    s.addText(`${brand.name}  ·  ${brand.orgs.provider_line}`, {
      x: 0.55, y: 4.75, w: 9, h: 0.3,
      fontSize: 11, color: P.mint, fontFace: "Arial", margin: 0,
    });
  }
```

- [ ] **Step 6: Update the output filename**

Find:
```javascript
  const out = path.join(
    "/home/shailesh/Downloads/Training/vibe-coding",
    "Vibe_Coding_Module_Slides_v2.pptx"
  );
```
Replace with:
```javascript
  const out = path.join(
    "/home/shailesh/Downloads/Training/vibe-coding",
    "Vibe_Coding_Module_Slides.pptx"
  );
```

- [ ] **Step 7: Run it**

Run: `node generate-slides.js`
Expected: `Wrote /home/shailesh/Downloads/Training/vibe-coding/Vibe_Coding_Module_Slides.pptx`

- [ ] **Step 8: Verify the pptx contains the brand strings (pptx is a zip of XML — grep the extracted slide XML)**

Run:
```bash
mkdir -p /tmp/slide_check && cd /tmp/slide_check && unzip -o -q "/home/shailesh/Downloads/Training/vibe-coding/Vibe_Coding_Module_Slides.pptx" -d extracted
grep -rl "CA Shailesh S Wadhawaniya" extracted/ppt/slides/ | head -3
grep -rl "StrideX Institute" extracted/ppt/slides/ | head -3
cd /home/shailesh/Downloads/Training/vibe-coding
```
Expected: both `grep -rl` calls print at least one matching slide XML file path.

- [ ] **Step 9: Archive the old v2 script (now superseded by the renamed, rebranded `generate-slides.js`) and the stale v2-named pptx**

```bash
git mv generate-slides-v2.js archive/legacy-scripts/
git mv Vibe_Coding_Module_Slides_v2.pptx archive/legacy-scripts/
```

- [ ] **Step 10: Add an npm script for convenience**

Modify `package.json`'s `"scripts"` block from:
```json
  "scripts": {
    "test": "echo \"Error: no test specified\" && exit 1"
  },
```
to:
```json
  "scripts": {
    "test": "echo \"Error: no test specified\" && exit 1",
    "build:slides": "node generate-slides.js"
  },
```

- [ ] **Step 11: Commit**

```bash
git add generate-slides.js Vibe_Coding_Module_Slides.pptx package.json
git add archive/legacy-scripts/generate-slides-v2.js archive/legacy-scripts/Vibe_Coding_Module_Slides_v2.pptx
git commit -m "feat: rebrand slide deck via brand.json, retire v2 script and stale output"
```

---

### Task 13: Root README describing the 3 deliverables

**Files:**
- Create: `README.md`

**Interfaces:** none (documentation only).

- [ ] **Step 1: Write `README.md`**

```markdown
# Vibe Coding with Claude Code

A course by CA Shailesh S Wadhawaniya — A StrideX Institute Program, in partnership with Wadhawaniya & Co.

## The 3 deliverables

| File | Purpose | Rebuild command |
|---|---|---|
| `Vibe_Coding_Instructor_Book.docx` | Everything the instructor needs: course design, the full technical encyclopedia, module-by-module teaching flow paired with hands-on workshops, and the facilitation playbook. | `.venv/bin/python build_instructor_book.py` |
| `Vibe_Coding_Student_Book.docx` | Self-contained student book: concepts, weekly labs, prompts, homework, and trackers in one place. | `.venv/bin/python build_student_book.py` |
| `Vibe_Coding_Module_Slides.pptx` | Slide deck to run live in class. | `node generate-slides.js` |

## Editing branding

All names, credentials, contact details, and colors live in `brand.json`. Change it once, then rebuild all 3 deliverables with the commands above.

## Folder layout

- `source-material/` — the original generated documents the two `.docx` builders excerpt content from. Required for rebuilding; do not delete.
- `archive/legacy-scripts/` — superseded generator scripts and outputs, kept for history only.
- `docx_merge.py` / `verify_docx.py` — shared Python helpers used by the two `.docx` builders.
```

- [ ] **Step 2: Commit**

```bash
git add README.md
git commit -m "docs: add README describing the 3 deliverables and how to rebuild them"
```

---

### Task 14: Final end-to-end regeneration and folder sanity check

**Files:** none created; this is a full-repo verification pass.

- [ ] **Step 1: Clean-rebuild all 3 deliverables from scratch**

```bash
rm -f Vibe_Coding_Instructor_Book.docx Vibe_Coding_Student_Book.docx Vibe_Coding_Module_Slides.pptx
.venv/bin/python build_instructor_book.py
.venv/bin/python build_student_book.py
node generate-slides.js
```
Expected: three `Wrote ...` lines, no tracebacks.

- [ ] **Step 2: Re-run every verification assertion from Tasks 9, 11, and 12 Step 8 against the freshly-rebuilt files**

Run:
```bash
.venv/bin/python -c "
from verify_docx import assert_contains, heading_list
assert_contains('Vibe_Coding_Instructor_Book.docx', [
    'CA Shailesh S Wadhawaniya', 'Part A — Course Design & Teaching Philosophy',
    'Part B — Technical Encyclopedia', 'Part C — Module-by-Module Teaching Flow',
    'Part D — Facilitation Playbook & Resources', 'Closing Note to the Instructor',
])
assert_contains('Vibe_Coding_Student_Book.docx', [
    'CA Shailesh S Wadhawaniya', 'Week 0 — Setup Checklist',
    'Chapter 14 — Quick Reference Cards', 'You Can Build Software',
])
print('instructor headings:', len(heading_list('Vibe_Coding_Instructor_Book.docx')))
print('student headings:', len(heading_list('Vibe_Coding_Student_Book.docx')))
print('ok')
"
```
Expected: `ok`, with both heading counts printed for a manual sanity glance.

- [ ] **Step 3: Confirm the root folder shows exactly the 3 deliverables plus build tooling — no stray old files**

Run: `ls *.docx *.pptx 2>&1`
Expected: exactly `Vibe_Coding_Instructor_Book.docx`, `Vibe_Coding_Student_Book.docx`, `Vibe_Coding_Module_Slides.pptx`

- [ ] **Step 4: Commit the freshly-rebuilt binaries (confirms the build is reproducible)**

```bash
git add Vibe_Coding_Instructor_Book.docx Vibe_Coding_Student_Book.docx Vibe_Coding_Module_Slides.pptx
git commit -m "chore: verify clean end-to-end rebuild of all 3 deliverables" --allow-empty
```

---

## What's explicitly deferred (Phase 2, not this plan)

- Deepening/rewriting the actual teaching content module-by-module, starting with the Instructor Book, once this phase ships.
- Any visual theme change beyond making the existing print-friendly green palette consistent (no dark/gold digital-brand redesign — that option was declined during design).
