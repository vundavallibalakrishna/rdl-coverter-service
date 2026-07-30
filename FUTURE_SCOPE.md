# Future Scope: RDL Fidelity and SSRS Certification

## Status

Deferred. The current report defects and deployment issues take priority. Work in this plan should begin
only after the active issues have been resolved, regression-tested, and accepted.

This is a forward-looking engineering plan, not a claim that the listed capabilities are implemented or
formally SSRS-certified.

## Goal

Establish defensible SSRS PDF certification, then close generic renderer gaps without report-specific
fixes or regressions to DOCX or XLSX.

Scope:

- PDF fidelity is the primary target.
- `DOCX_VISUAL` inherits certified PDF output.
- `DOCX_EDITABLE` and XLSX remain native/editable and are regression-tested separately.
- Bundled subreport rendering remains supported for PDF and `DOCX_VISUAL`.
- Data-source execution and custom-code execution remain out of scope.
- Every implementation must be driven by RDL constructs and the normalized model, never by report names,
  client text, or document-specific geometry.

## Phase 0: Freeze the Baseline

Before changing rendering:

1. Record the current commit, Node.js version, dependency versions, Poppler version, and font versions.
2. Run the complete corpus with the current RDL and JSON pairs.
3. Record per output:
   - Success or failure.
   - Page count and dimensions.
   - Layout mode.
   - Extracted text count.
   - Font inventory.
   - Package integrity.
   - Worker-memory estimate.
4. Record bundled-subreport parent results separately from definition-only analysis.
5. Write generated evidence directly under `tmp/`, with no artifact subdirectories.
6. Preserve hashes for every RDL, JSON payload, reference PDF, and font file.

Acceptance:

- Existing compatible corpus outputs are reproducible.
- No unexplained failures remain.
- Baseline reporting clearly separates definition-only compatibility from bundle-aware rendering.

## Phase 1: Build the SSRS Certification Harness

Create `scripts/certify-pixel.js`.

### Inputs

Use an external certification manifest:

```json
{
  "rdl": "/secure/path/report.rdl",
  "request": "/secure/path/request.json",
  "referencePdf": "/secure/path/ssrs-reference.pdf",
  "fontDirectory": "/secure/path/fonts",
  "exactInputs": true,
  "expectedHashes": {}
}
```

Client reports and data remain outside Git unless they are explicitly sanitized and approved as fixtures.

### Certification Checks

For every page:

- Exact page count.
- Page dimensions within `0.01 pt`.
- Item geometry within `0.5 pt`.
- Extracted text presence and ordering.
- Important text bounding-box anchors.
- Font-family and embedded-font verification.
- Border/vector-position checks.
- Canonical 144-DPI pixel comparison.
- Optional 300-DPI diagnostic comparison for hairline borders.

The 144-DPI result remains the formal gate. The 300-DPI result is diagnostic.

### Diagnostics

For failing pages, generate directly under `tmp/`:

- Generated page PNG.
- Reference page PNG.
- Difference heatmap.
- Per-page JSON report.

Process one page at a time to avoid excessive memory use on long reports.

### Certification Rule

Set `certified: true` only when:

- Inputs are confirmed to come from the same SSRS run.
- All hashes match.
- Licensed font versions match.
- No required checks were skipped.
- Every required threshold passes.

A skipped CI job must never count as certification.

## Phase 2: Small, Proven Core Fixes

### 2.1 Embedded-Image Validation

External and Database images are already rejected. Add validation for Embedded images that:

- Reference a nonexistent embedded-image key.
- Have missing data.
- Contain invalid Base64.
- Declare an unsupported MIME type or image format.

The service must return a stable `UNSUPPORTED_FEATURE` or `RDL_INVALID` error rather than silently dropping
the image.

### 2.2 Line-Style Rendering

Make RDL `Line` items use shared stroke semantics:

- Solid.
- Dashed.
- Dotted.
- Double, when valid for the construct.
- Width expressions.
- Color expressions.
- Minimum PDF stroke width.
- Correct diagonal direction.

Add minimal synthetic RDL tests plus raster and vector assertions.

### 2.3 Schema-Audit Enforcement

Change the schema audit from inventory-only reporting to a real gate:

- Every known schema path is classified.
- No encountered path remains unclassified.
- Catalogue totals are internally consistent.
- Unknown required constructs remain rejected.

### 2.4 Bundle-Aware Analysis

Do not rebuild subreport rendering; it already exists. Improve `/v1/analyze`:

- Without bundles, report missing bundled definitions clearly.
- With complete bundles, resolve children safely and recalculate compatibility.
- Report compatibility separately for PDF, `DOCX_VISUAL`, `DOCX_EDITABLE`, and XLSX.
- Preserve fail-closed behavior for missing invocation data, cycles, excessive nesting, and unsupported
  child constructs.

## Phase 3: Introduce a PDF Layout Scene

Before replacing the current body flow, separate layout from drawing.

Create an internal scene containing:

- Resolved items.
- Parent container.
- Design coordinates.
- Measured dimensions.
- Resolved visibility.
- Growth dimensions.
- Peer dependencies.
- Page fragments.
- Z-order.
- Final drawing coordinates.

The scene produces final page geometry first and draws the PDF second.

Benefits:

- Headers and footers know the final page count.
- Layout becomes inspectable without raster comparison.
- Body positioning and horizontal pagination share one model.
- Drawing code no longer makes incremental layout decisions.

No public API changes are required.

## Phase 4: Constraint-Based Two-Dimensional Body Layout

Do not replace flow layout with simple absolute positioning.

For each body or rectangle container:

1. Preserve declared `Left`, `Top`, width, and height.
2. Identify peer relationships.
3. Preserve declared minimum spacing.
4. Allow side-by-side items to remain side-by-side.
5. Propagate vertical growth only to dependent items below.
6. Propagate horizontal growth to dependent items on the appropriate side.
7. Apply conditional-visibility whitespace rules.
8. Keep rectangle-contained items in their own coordinate system.
9. Preserve true hard-page overlaps using z-order.
10. Detect dependency cycles or unsupported ambiguity and fail closed.

Initial synthetic fixtures:

- Two textboxes with identical `Top` and different `Left`.
- Side-by-side tablix and chart.
- Growing tablix with content below.
- Growing tablix with a peer item to its right.
- Nested rectangles.
- Conditionally hidden peer.
- Intentionally overlapping items.
- Explicit page break between peer bands.

Acceptance:

- Correct SSRS-like peer movement.
- Existing stacked reports retain layout and page counts unless the certification oracle proves the
  previous result wrong.
- No item enters the footer band.

## Phase 5: Horizontal Hard-Page Pagination

Implement this only after the layout scene is stable.

For PDF:

- Split oversized tablixes and matrices into horizontal page slices.
- Repeat eligible row-header columns.
- Preserve tablix corner cells.
- Respect row and column spans.
- Respect left-to-right and right-to-left direction.
- Combine horizontal and vertical slices deterministically.
- Number pages left-to-right, then top-to-bottom.
- Preserve explicit group and page breaks.
- Fail closed if one unsplittable column exceeds the printable width.

Do not gate pagination itself behind advanced materialization. Materialization determines the expanded
columns; the PDF layout stage determines page slices.

`DOCX_VISUAL` inherits this behavior through PDF. `DOCX_EDITABLE` and XLSX remain unchanged.

## Phase 6: Border Investigation Before Modification

Do not change double-border geometry without evidence. First certify:

- Stroke coordinates.
- Widths.
- Line caps and joins.
- Shared-edge deduplication.
- Double-border strand separation.
- Corners at 144 and 300 DPI.
- Behavior across representative PDF viewers and zoom levels.

Classify each observation:

- Incorrect vector geometry: renderer defect.
- Correct vectors but zoom-dependent raster thinning: viewer artifact.
- Anti-aliasing-only difference: certification-tolerance issue.

Only then modify double-border joins or edge-merging tolerance. Every change requires a synthetic regression
proving that distinct nearby lines are not fused.

## Phase 7: Global-Value Policy

Follow SSRS scope rules:

- `PageNumber` and `TotalPages` are supported in page headers and footers.
- Body usage fails closed or is explicitly documented as a non-SSRS extension.
- Capture `ExecutionTime` once per render and reuse it everywhere.
- Add tests for header/footer page fields across multi-page reports.

Do not add complex body `TotalPages` rendering unless an official SSRS-compatible construct requires it.

## Phase 8: Verification and Release

For every phase:

1. Add focused synthetic regression tests.
2. Run `npm test`.
3. Run `npm run audit:schema`.
4. Run the stress and nested-stress suites.
5. Render the corpus to PDF, DOCX, and XLSX.
6. Run the bundled-subreport suite.
7. Verify package integrity.
8. Verify fonts.
9. Run the private SSRS certification pack.
10. Run the Docker smoke test with licensed fonts.
11. Deploy a canary through `rdl.curasoftware.com`.

Release only when:

- No corpus regression remains unexplained.
- All artifacts are valid.
- Strict-font validation passes.
- Certification reports distinguish engineering validation from formal SSRS certification.
- The previous Docker image remains available for rollback.

## Recommended Implementation Order

1. Freeze the baseline.
2. Build the certification harness.
3. Address image, line, schema, and analysis gaps.
4. Introduce the layout scene.
5. Implement two-dimensional peer layout.
6. Implement horizontal pagination.
7. Make evidence-driven border changes.
8. Enforce the global-value policy.
9. Complete certification and deployment.

This sequence creates measurable evidence before changing the most regression-prone areas: body layout and
pagination.
