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

## License

Licensed under the [PolyForm Noncommercial License 1.0.0](LICENSE) — free to use, copy, modify, and share for any noncommercial purpose. Commercial use requires separate written permission from CA Shailesh S Wadhawaniya.
