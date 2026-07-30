---
name: rdl-windows-word-fidelity
description: Build, inspect, and certify page-locked native editable DOCX output against the RDL Converter Service's canonical PDF in Microsoft Word for Windows. Use for DOCX_EDITABLE renderer work, Word page-count or table-geometry defects, font embedding, OOXML package audits, and final Windows Word fidelity certification.
---

# RDL Windows Word Fidelity

## Workflow

1. Read `AGENTS.md` and `references/windows-word-contract.md`.
2. Use the canonical PDF and its internal layout trace as the only layout authority.
3. Keep report text in native `w:t` runs and report tables in native `w:tbl` grids.
4. Run `scripts/audit_docx_ooxml.mjs <docx> --out tmp/<name>-docx-audit.json`.
5. Copy the DOCX and canonical PDF to a licensed Windows QA host with Microsoft Word installed.
6. Run `scripts/export_word_pdf.ps1 -InputDocx <docx> -OutputPdf <word.pdf> -ResultJson <word.json>`.
7. Run the PDF certification skill against the canonical PDF and Word-exported PDF.
8. Do not certify unless Word opens without repair or external-field prompts and all gates pass.

## Hard Rules

- Microsoft Word for Windows is the sole pagination authority for DOCX certification.
- Require exact page count, page dimensions, explicit line wrapping, table fragments, cell widths, row heights, borders, fonts, headers, and footers at generation time.
- Editing may repaginate the document; that does not weaken the initial-generation gate.
- Never use full-page screenshots, rasterized tables, or shape-per-line text.
- RDL images and chart renders may remain pictures; report text and tablixes must remain native.
- Embed every consumed font variant only when OpenType `OS/2.fsType` permits editable embedding.
- Fail closed for unrepresentable overlaps, unsupported rotations, Word tables over 63 columns, and pages over 22 by 22 inches.
- Never introduce a report-specific branch or layout constant.

## Resources

- `scripts/audit_docx_ooxml.mjs` checks native content, sections, tables, drawings, fonts, and external relationships.
- `scripts/export_word_pdf.ps1` performs authoritative Word repagination and PDF export.
- `references/windows-word-contract.md` records the OOXML and Windows-only certification contract.
