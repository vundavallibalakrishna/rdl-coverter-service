# Repository Agent Instructions

Guidance for coding agents working on the standalone RDL Converter Service.

## Instruction File Synchronization

- `AGENTS.md` and `CLAUDE.md` must remain byte-for-byte identical.
- Every instruction change must update both files in the same change. Never modify, add, remove, or reorder
  guidance in only one of them.
- Before finishing an instruction-file change, verify synchronization with `cmp AGENTS.md CLAUDE.md`.

## Required Renderer Skills

- Before changing PDF measurement, pagination, drawing, borders, fonts, or the internal canonical layout
  trace, read and follow `skills/rdl-pdf-layout-certification/SKILL.md`.
- Before changing or certifying `DOCX_EDITABLE`, Word page geometry, native tables, font embedding, or OOXML
  packaging, read and follow `skills/rdl-windows-word-fidelity/SKILL.md`.
- For PDF-locked editable Word work, use the PDF skill first to establish the canonical layout and the
  Windows Word skill second to build and certify the native DOCX.
- Microsoft Word for Windows is the sole authoritative viewer for `DOCX_EDITABLE` pagination and visual
  certification. Do not use Word for Mac, LibreOffice, Google Docs, or browser previews as release evidence.

## Service Boundary

This repository is an independent Node.js service. It is deliberately separate from Lumina.

- Do not import code, configuration, database models, authentication, or runtime services from Lumina.
- Do not modify the Lumina repository while working here unless the user explicitly requests a separate Lumina integration task.
- Keep this service stateless: no database, Redis, queue, persistent report storage, or external data-source execution.
- The caller supplies every parameter and dataset row required to render the report.
- RDL data sources, queries, and stored procedures are metadata only. Never execute, return, or log them.

## Runtime and API

The service requires Node.js 22 and uses Fastify with ESM modules.

Endpoints:

- `GET /healthz`: process liveness.
- `GET /readyz`: temporary storage, fonts, PDFKit, and Poppler readiness.
- `POST /v1/analyze`: safe RDL analysis and compatibility response.
- `POST /v1/render`: synchronous binary rendering of one requested output.

Both POST endpoints accept JSON/base64 and multipart requests. Render outputs are:

- `PDF`: selectable fixed-layout PDF generated directly from the normalized RDL model.
- `DOCX_EDITABLE`: Windows-only page-locked native editable Word content constructed from the canonical
  PDF renderer's resolved layout trace. The PDF is an internal layout authority, not a page screenshot or an
  external PDF-to-Word conversion.
- `DOCX_VISUAL`: PDF rendered internally, rasterized at 300 DPI, then packaged as exactly one full-page image per Word page.
- `XLSX`: native Excel workbook (exceljs); tablixes become styled cell blocks with live typed numbers/dates.

There is no `DOCX_FIXED_EDITABLE`. Do not reintroduce the former continuous/reflowable renderer,
shape-per-line text, or full-page screenshots in `DOCX_EDITABLE`. It must use native Word tables and text,
with declared RDL images and charts as drawings. `DOCX_VISUAL` remains the separate raster page-image mode.

`DOCX_EDITABLE` reports the numeric canonical PDF page count, `X-Docx-Layout-Mode:
windows-paged-editable`, and `X-Docx-Editable-Text-Ratio`. PDF and visual DOCX also return numeric canonical
page counts.

Microsoft Word is not installed or invoked on the production server. The Linux service writes OOXML directly; Word is an authoritative release-certification viewer on a developer/QA workstation only.

## Repository Map

```text
src/server.js                 Process entrypoint and graceful shutdown
src/app.js                    Fastify routes, headers, logging, and error mapping
src/config.js                 Runtime limits and environment configuration
src/request.js                JSON/multipart input and filename sanitization
src/readiness.js              Storage, font, PDFKit, and Poppler probes
src/rdl/parser.js             Secure XML parsing and normalized intermediate model
src/rdl/capabilities.js       Schema/path and expression capability policies
src/rdl/rdl2016SchemaNames.js Published 2016 XSD element/attribute inventory
src/rdl/expression.js         Safe SSRS expression evaluator; never use eval
src/rdl/validation.js         Parameters, exact DataField rows, filters, and sorting
src/render/pdf.js             PDF layout, pagination, drawing, and canonical trace recording
src/render/layoutTrace.js     Deterministic PDF layout trace model and validation
src/render/docx.js            Windows-paged editable DOCX public renderer
src/render/pagedDocx.js       Trace-driven native Word page/table construction
src/render/windowsWordCompatibility.js Windows Word geometry and request policy
src/render/excel.js           Native XLSX generation (exceljs)
src/render/visualDocx.js      PDF-to-300-DPI-page-image Word generation
src/render/fonts.js           Licensed font discovery and fail-closed checks
src/worker/runner.js          Concurrency, worker lifetime, timeout, and cleanup
src/worker/renderWorker.js    Isolated render process
test/                         Unit, API, artifact, isolation, and cleanup tests
scripts/render-sample.js      Host smoke test for all output modes
scripts/docker-smoke.js       Build/start/readiness/render/cleanup Docker smoke test
Dockerfile                    Production Node 22 and Poppler image
compose.yaml                  Hardened local/private-network container profile
```

## Security Invariants

These are non-negotiable:

- XML DTDs and entity declarations remain rejected.
- External-entity processing remains disabled.
- Enforce RDL byte, complete request, XML node, XML depth, row, concurrency, timeout, and worker-memory limits.
- Unknown required namespaces and unsupported constructs must fail closed.
- Every encountered element and attribute path must be classified. Unclassified paths default to `REJECTED`; never widen the default.
- Never add JavaScript `eval`, `Function`, VM execution, or dynamic module loading for SSRS expressions.
- Never execute RDL query text, data-source references, report custom code, or external images.
- Never log RDL XML, queries, parameters, dataset values, temporary-file contents, or output bytes.
- Temporary request directories must be unique and `0700`; contained files must be `0600`.
- Generate and validate the complete artifact before sending successful response headers.
- Clean temporary data after success, validation failure, renderer failure, timeout, disconnect, and shutdown.
- Preserve stable JSON error codes and sanitized messages.
- Do not silently substitute fonts in strict/production mode.
- Keep CORS and application authentication absent unless the user explicitly expands the architecture. This service must remain private-network-only behind TLS and network/IP allowlisting until authentication is added.

## RDL Data Rules

RDL distinguishes the internal field `Name` from `DataField`. Incoming JSON row keys must match `DataField` exactly, including spaces, punctuation, and case. The normalization layer may map those values to internal field names only after validation.

Parameter dropdown datasets are optional when concrete parameter values are supplied, unless the same dataset is also required by rendered content. Missing rendered datasets, fields, and required parameters must return stable validation errors.

Preserve these model concepts when extending the parser:

- Page dimensions, margins, headers, footers, and page breaks.
- Textboxes, paragraphs, text runs, styles, visibility, and formats.
- Embedded images, rectangles, and lines.
- Tablix columns, rows, cells, column spans, hierarchies, groups, filters, sorts, repeated headers, and visibility.
- Dataset field-name/DataField mappings and parameter metadata without query contents.

Supported namespaces are adapted into the same intermediate model. Add broader RDL coverage through namespace adapters and certified fixtures, not by scattering version checks through renderers.

## Rendering Rules

- Use RDL units converted to points as the PDF layout source of truth.
- Keep page measurement and drawing deterministic for the same inputs, installed font files, and runtime version.
- PDF text must remain selectable.
- PDF trace capture must be non-invasive: recording layout may observe resolved drawing operations but must
  not change PDF bytes, geometry, pagination, text, borders, or fonts.
- Construct `DOCX_EDITABLE` only from the validated canonical PDF layout trace. Create one next-page Word
  section per PDF page and preserve exact page dimensions, page-specific headers/footers, native text,
  explicit PDF wrap points, fixed table grids, cell widths, exact row heights, merges, repeated headers,
  margins, fills, shared borders, and fragment-closing borders.
- Compatible free-form textboxes and rectangles become native cells. Lines become cell borders. Only
  declared RDL images and charts may be drawings; never emit a page screenshot, rasterized table, or
  shape-per-line text.
- Reject obsolete DOCX fragmentation/profile settings with `RDL_INVALID`. Fail with
  `UNSUPPORTED_FEATURE` when Windows Word cannot safely represent the traced geometry, including incompatible
  editable overlaps, unsupported rotations, required tables wider than 63 columns, and pages exceeding
  22 by 22 inches.
- Any expression-capable style property must be resolved through the style helpers at render time (see the `EXPRESSION_PROPERTIES` catalogue and the renderer raw-consumption guard test); never regex/compare/measure a raw `=expression` string.
- Never implement visual DOCX through any path other than the internally generated PDF and 300-DPI page rasterization.
- Correctly account for column spans by summing the widths of every covered grid column.
- Cap every PDF tablix segment, including row-span geometry, at the printable body boundary (`page height - bottom margin - footer height`). A split row must never draw backgrounds, text, or borders inside the footer band.
- Repeated headers and explicit page breaks must not add blank trailing pages.
- Unsupported behavior must produce `UNSUPPORTED_FEATURE`, not a visually misleading approximation.
- Keep the schema capability catalogue separate from the SSRS/VB expression catalogue. XSD coverage does not imply expression support.

## Fonts

Production requires legally licensed regular, bold, italic, and bold-italic variants of Arial and Times New Roman mounted into `/app/fonts`. Reports that declare another consumed family require its licensed variants too; the Combined Assurance client RDL declares Segoe UI.

- Leave `RDL_STRICT_FONTS=true` in production.
- Readiness must fail when a required variant is absent.
- Tests may use `RDL_STRICT_FONTS=false` when licensed fonts are unavailable.
- Do not copy or commit licensed font files into this repository or Docker image.
- `DOCX_EDITABLE` must embed every consumed regular, bold, italic, and bold-italic face. Inspect
  `OS/2.fsType`, permit only installable or editable embedding, honor the no-subsetting flag, obfuscate font
  parts per WordprocessingML, and fail with `FONT_MISSING` or `FONT_EMBEDDING_FORBIDDEN` instead of
  substituting.

## Development Commands

Run from this repository root:

```bash
npm install
npm start
npm test
npm run audit:schema
npm run test:coverage
npm run smoke:sample -- "/path/to/report.rdl"
RDL_FONT_HOST_DIR="/path/to/licensed-font-directory" npm run smoke:docker -- "/path/to/report.rdl"
RDL_FONT_HOST_DIR="/path/to/licensed-font-directory" docker compose up --build
```

Do not hand-edit `package-lock.json`. Use npm for dependency changes and avoid unnecessary dependency churn.

## Generated Files and Test Artifacts

- Use the repository-root `tmp/` directory as the only location for files generated while developing,
  rendering, debugging, smoke-testing, comparing, or certifying this service.
- Write those files directly into `tmp/`. Do not create ad hoc artifact directories elsewhere in the
  repository, and do not create task-, report-, run-, or format-specific subdirectories under `tmp/`.
- Use clear, collision-resistant filenames when multiple artifacts must coexist, and remove obsolete files
  instead of introducing another folder hierarchy.
- The service's security-controlled, unique per-request temporary directories are the sole exception. They
  are runtime implementation details under the configured temporary root and must retain their existing
  permissions and complete cleanup guarantees.

## Testing and Quality Bar

Every parser, evaluator, validator, renderer, API, or isolation change needs the smallest focused regression test plus the relevant broader verification.

- Parser changes: test namespaces, security limits, unsupported constructs, and the normalized model.
- Expression changes: test the exact expression and nested combinations without weakening fail-closed behavior.
- Dataset/parameter changes: test defaults, multi-value parameters, exact `DataField` names, missing fields, filters, sorts, and row limits.
- PDF changes: verify page count, page size, extracted text, raster dimensions, and visual output.
- Editable DOCX changes: inspect ZIP/OpenXML for native `w:t` and `w:tbl`, exact twip geometry, explicit
  line breaks, merges, embedded-font relationships, no external links, no full-page image, and no
  shape-per-line content.
- Certify `DOCX_EDITABLE` only in Microsoft Word for Windows: force pagination through PowerShell/COM,
  export to PDF without updating fields or prompting, then require exact page count/dimensions and displayed
  text/order, geometry within 0.5 pt, and the specified 144-DPI pixel threshold. Other viewers are not
  certification evidence.
- Visual DOCX changes: verify one PNG relationship and media file per PDF page.
- Worker/API changes: test concurrency rejection, `Retry-After`, timeouts, disconnect/abort behavior, response headers, and an empty temporary root afterward.
- Docker/runtime changes: run the Docker smoke test with licensed fonts mounted read-only.

Never delete, skip, or loosen meaningful tests to make a change pass. Fix the root cause.

## Supplied Report and Certification

The original development sample is `Combined Assurance Reports (2).rdl`, supplied outside this repository. Smoke scripts accept its path and generate synthetic rows from exact `DataField` definitions. Do not copy the report or its generated artifacts into the repository unless the user explicitly requests a sanitized fixture.

Supplied client reports are verification oracles only. Do not encode their names, hashes, item names,
visible strings, dimensions, column counts, row counts, or data values in production behavior.

Dynamic column groups (matrix / cross-tab) with column-hierarchy `TablixHeader` and `TablixCorner`, group
header/footer subtotal rows, and recursive/parent (`Group/Parent`) row groups are implemented. Bundled,
invocation-scoped subreports are supported in PDF, `DOCX_EDITABLE`, `DOCX_VISUAL`, and XLSX `REPORT`; XLSX
`DATA` subreports remain fail-closed. `Group/Variables` are resolved as `Variables!Name.Value` in the current row scope via
`globals.variables`; `DomainScope` and `NaturalGroup` are acknowledged metadata. The `Aggregate()` function,
charts beyond the supported set, maps, gauges, custom code, external resources, and other uncatalogued
variants remain fail-closed.

Synthetic or reconstructed hydration success proves parser, renderer, packaging, worker isolation, and cleanup behavior only for the supplied rows. It does not prove SSRS pixel equivalence. `DOCX_VISUAL` must use floating full-page images plus explicit page breaks so LibreOffice and Word produce exactly the PDF page count; inline page-height images create blank overflow pages and are forbidden.

Do not claim formal pixel-perfect certification until the user supplies all of the following from the same SSRS run:

- Reference PDF.
- Exact parameters.
- Exact rows for every rendering dataset.
- Exact licensed font files or recorded font versions.

Certification must then verify exact page count and dimensions, displayed text and ordering, geometry within 0.5 pt, and the specified 144-DPI pixel-difference tolerance. The repository reference PDF and reconstructed hydration are useful regression inputs, but they are not from the same SSRS data run and therefore cannot satisfy certification.

## Change Workflow

**Fixes must be generic, never report-specific.** This overrides everything else in this section. Fix the
root cause for the RDL *construct*, driven by the normalized model (coordinates, sizes, styles, alignment,
expressions, group structure) — not tuned to Combined Assurance, KRI, Incident, or the sample in front of
you. Concretely:

- **Every root-cause fix requires a cross-renderer impact audit.** Before changing code, classify whether
  the corrected behavior is an RDL/model semantic shared by PDF, `DOCX_EDITABLE`, `DOCX_VISUAL`, and XLSX,
  or a genuinely format-specific behavior. For a shared semantic, fix it at the lowest common layer when
  possible and implement any necessary renderer adaptations in **every applicable output in the same
  change**. Never land a PDF-only, Word-only, or Excel-only workaround when the same RDL construct is also
  consumed by another renderer.
- A shared fix is not complete until focused tests cover the construct in PDF, `DOCX_EDITABLE`, and XLSX;
  also verify `DOCX_VISUAL` when the change can affect canonical PDF bytes, page geometry, pagination, or
  rasterization. Tests must assert the format-appropriate equivalent—not assume that PDF drawing commands,
  Word OOXML, and Excel cells have identical mechanics.
- If an output is intentionally unaffected or cannot represent the construct, record the concrete
  format-semantic reason in the code/test or change report and add or preserve fail-closed capability
  detection where appropriate. “The issue was reported only in one format” is not a valid reason to skip
  the other renderers.
- During RCA and handoff, report a renderer-impact matrix covering PDF, `DOCX_EDITABLE`, `DOCX_VISUAL`, and
  XLSX with one of: fixed and tested, inherited from canonical PDF, not applicable with reason, or
  unsupported and fail-closed. Do not describe a generic issue as fixed while an applicable renderer still
  contains the same defect.
- **PDF and DOCX fixes are global root-cause fixes.** Never identify or branch on an output filename,
  report name, definition hash, item name, visible text, dataset value, page number, known row count, or
  client-specific layout signature. An `if`/`else`, lookup, profile, tolerance, or geometry adjustment is
  permitted only when its predicate is a general RDL/model/format semantic that applies to every report
  using that construct.
- Fix the lowest shared layer that owns the defect: parsing, normalization, expression resolution,
  materialization, measurement, layout, pagination, border resolution, or font selection. Renderer-specific
  code is acceptable only where PDF and Word have genuinely different format semantics, and it must still
  handle the construct generically rather than recognize a document.
- Cover the construct's meaningful variants, including literal and expression-backed properties,
  visibility, nesting, groups, spans, page boundaries, style inheritance, fonts, and empty/large data.
  Do not add a narrow branch that handles only the observed fixture while leaving equivalent cases broken.
- Every regression must include a minimal synthetic trigger plus counterexamples or variants that prove the
  rule is general. Validate at least one unrelated real report when practical. Real client artifacts are
  verification oracles only and must never supply production predicates or constants.
- If the generic semantics cannot be established safely, fail closed with `UNSUPPORTED_FEATURE` and
  document the unsupported variant instead of introducing a report-specific approximation.
- No hardcoded values, element `Name`s, text strings, counts, thresholds, pixel offsets, or magic numbers
  that only make one report look right. If a change would break on a different RDL that uses the same
  construct, it is wrong.
- Derive behavior from what the RDL declares. Read more of the model to raise fidelity; never special-case a
  document.
- Expression-capable properties resolve through the style helpers (`styleValue`/`styleColor`/`styleSize`/
  `isHidden`); the `EXPRESSION_PROPERTIES` catalogue and the renderer raw-consumption guard test enforce this.
- Write regression tests against a **synthetic minimal RDL** isolating the construct. Use the real client
  reports only as a verification oracle (render them, diff against the PDF) — never fix *to* them.

1. Identify the RDL construct or runtime boundary being changed.
2. Update the normalized model before adding renderer-specific workarounds.
3. Keep parsing, validation, layout, rendering, and HTTP concerns separated.
4. Add fail-closed compatibility detection for unsupported variants.
5. Add focused tests (synthetic RDL) and inspect representative artifacts visually.
6. Run `npm test`.
7. For rendering or deployment changes, run both supplied-sample and Docker smoke tests.
8. Report clearly whether the result is implemented, smoke-tested, or formally SSRS-certified.
