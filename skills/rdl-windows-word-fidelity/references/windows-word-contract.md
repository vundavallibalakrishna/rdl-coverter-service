# Windows Word Fidelity Contract

## Authority

Microsoft Word for Windows is the only authoritative DOCX layout viewer. Production Linux creates OOXML
but never invokes Word. Final release certification therefore runs on a licensed Windows QA workstation.

## Native representation

- Use one next-page section per canonical PDF page.
- Use fixed Word table grids with explicit table, grid-column, and cell widths.
- Use exact row heights from the PDF trace.
- Materialize the PDF's line breaks as native Word line breaks.
- Materialize repeated tablix headers and page-dependent page bands per traced page.
- Keep every report textbox and tablix value in `w:t`.
- Permit drawings only for declared RDL images and rendered charts.

## Platform limits

- Word supports a maximum physical page size of 22 by 22 inches.
- A Word table cannot safely exceed 63 columns.
- Unrepresentable overlapping native regions and unsupported writing directions fail with
  `UNSUPPORTED_FEATURE`; they are never scaled or rasterized.

## Fonts

Embed regular, bold, italic, and bold-italic faces using the matching Word font-table relationships.
Inspect the OpenType `OS/2.fsType` flags first. Allow installable embedding (`0`) and editable embedding
(`8`). Restricted, preview/print-only, and bitmap-only fonts cannot satisfy editable certification.
Honor the no-subsetting flag.

## Windows release gate

Open with link updates and macros disabled, repaginate, export to PDF, and compare every exported page with
the canonical PDF. The gate requires no repair prompt, exact page count and dimensions, matching text and
ordering, geometry within 0.5 point, and the repository's 144-DPI pixel tolerance.
