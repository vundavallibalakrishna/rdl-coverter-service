---
name: rdl-pdf-layout-certification
description: Capture and certify the canonical selectable PDF produced by the RDL Converter Service. Use for PDF renderer changes, pagination or geometry investigations, font and border RCA, canonical PDF manifests, or PDF-to-DOCX fidelity work that treats the PDF layout as authoritative.
---

# RDL PDF Layout Certification

## Workflow

1. Read `AGENTS.md` and keep every generated artifact directly under repository `tmp/`.
2. Generate the PDF from the exact RDL, parameters, datasets, subreports, and licensed font versions being certified.
3. Run `scripts/capture_pdf_manifest.mjs <pdf> --out tmp/<name>-pdf-layout.json`.
4. Inspect page size, page count, embedded fonts, selectable text, and text bounding boxes before visual comparison.
5. For a renderer change, generate before/after PDFs from identical inputs and run
   `scripts/compare_pdf_layout.mjs <reference.pdf> <candidate.pdf> --out tmp/<name>-pdf-diff.json`.
6. Treat page-count, page-size, missing-text, missing-font, or geometry differences as blocking. Never explain them away with a pixel-only comparison.
7. Inspect the rendered 144-DPI pages when a comparison reports a failure.

## Certification Contract

- Require identical page count and dimensions.
- Require selectable text with the same normalized content and order.
- Require all consumed font families and variants to be present without substitution.
- Require geometry within 0.5 point.
- At 144 DPI, allow at most 0.5% of pixels per page to differ by more than 16 RGB levels.
- Never mark reconstructed data as SSRS-certified.
- Never add a report-name, filename, page-number, visible-text, or dataset-value production predicate.

Read `references/layout-contract.md` before changing PDF measurement, pagination, borders, or the internal layout trace.

## Resources

- `scripts/capture_pdf_manifest.mjs` creates a deterministic PDF geometry/font/text manifest.
- `scripts/compare_pdf_layout.mjs` performs the complete page and raster gate.
- `references/layout-contract.md` records the trace and certification invariants.
