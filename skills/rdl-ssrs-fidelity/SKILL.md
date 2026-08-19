---
name: rdl-ssrs-fidelity
description: The authoritative model of how Microsoft SQL Server Reporting Services (SSRS) renders an RDL, and the rule that this engine must reproduce that behavior exactly. Use for ANY rendering defect RCA (borders, sizing/growth, value formatting, charts/axes, page headers/footers, positioning, pagination) before changing PDF, DOCX_EDITABLE, DOCX_VISUAL, or XLSX. Consult it to decide "what is the correct output?" — the answer is always "what SSRS produces," never "what makes this sample look right."
---

# RDL SSRS Fidelity

The goal of this service is a license-free renderer whose output is **indistinguishable from Microsoft SSRS**
for every supported construct. This skill is the reference for *what SSRS does*, so that every fix is a
step toward matching SSRS rather than a tweak toward one PDF.

## Governing principle

**The oracle is Microsoft SSRS, never the sample.** When output looks wrong, do not ask "what change makes
this report look right?" Ask "what does SSRS render for this RDL construct, and why does our engine differ?"
Fix the difference at the layer that owns the construct (parse → normalize → resolve expression → measure →
lay out → paginate → draw), generically, for every renderer that consumes it.

A defect is only understood once you can state the SSRS rule it violates. If you cannot state the rule,
you have not finished the RCA — read `references/` and, when a rule is genuinely unknown, obtain real SSRS
output and treat it as ground truth (see *Oracle*). Never encode a report name, filename, item name,
visible string, dataset value, page number, row count, or definition hash as a production predicate.

## Workflow for any rendering RCA

1. **Reproduce from the normalized model.** Render the exact RDL + parameters + datasets through the
   in-process renderer and capture the real drawing operations, not a stale artifact. A PDF someone handed
   you may predate the current code — regenerate before trusting it. Intercepting PDFKit draw calls
   (strokeColor/fillColor/moveTo/lineTo/rect/stroke/fill) or reading the layout trace gives exact geometry
   and color per edge; pixel rasters confirm the visible result.
2. **State the SSRS rule.** Identify the construct and look up the governing behavior in the relevant
   `references/` contract below. Name the rule in the RCA.
3. **Locate the lowest shared layer** that owns the construct and diverges from that rule.
4. **Fix generically**, and implement the matching adaptation in **every** applicable renderer in the same
   change (PDF, DOCX_EDITABLE, DOCX_VISUAL, XLSX). Record a renderer-impact matrix.
5. **Verify against SSRS semantics**, not against the sample: a synthetic minimal RDL isolating the
   construct plus its variants (literal and expression-backed, visibility, nesting, groups, spans, page
   boundaries, style inheritance, fonts, empty and large data). When a real SSRS reference from the same
   run is available, diff against it per the PDF certification skill.
6. If the correct SSRS semantics cannot be established safely, **fail closed** with `UNSUPPORTED_FEATURE`
   and document the unsupported variant — never ship a report-specific approximation.

## Hard rules (what SSRS does — reproduce exactly)

- **Formatting.** A value is rendered through the .NET format engine using its `Format` string. With no
  `Format`, SSRS uses the value's default .NET `ToString()` under the report `Language`/culture — never a
  host-language coercion. A `DateTime` with no format is the general date/time pattern (date **and** time),
  not a day-of-week/timezone dump and not a bare date. See `references/value-formatting.md`.
- **Sizing and growth.** Content grows the item; the item grows its container; the container's border and
  background are painted at the **rendered** extent, not the declared height. Coincident edges of a
  container and its grown child stay coincident. Page header/footer are fixed height and clip; the body
  grows. See `references/sizing-and-growth.md`.
- **Borders and shared edges.** Borders belong to each item/cell edge. Adjacent cells share an edge at the
  **same** coordinate, so their borders coincide and read as one line; a visible double is an engine bug
  (mismatched geometry or a redundant second draw), not SSRS behavior. See `references/border-resolution.md`.
- **Charts and axes.** The value axis uses SSRS "nice" 0-based 1/2/5·10ⁿ scaling; tick labels are formatted
  (no binary-float noise). Category labels are auto-fit by shrink → rotate → wrap and the plot area yields
  space to them — SSRS does **not** clip axis labels. See `references/charts-and-axes.md`.
- **Positioning and flow.** Report items are absolutely positioned by `Top`/`Left` on their container's
  canvas with `ZIndex` order and may overlap or sit side-by-side; they are not a vertical stack. Pagination
  happens in the growing body. See `references/page-and-flow.md`.
- **Determinism.** Same RDL + inputs + fonts + runtime → identical bytes. `Globals!ExecutionTime` is the
  one intentional exception.

## Oracle

The only proof of "matches SSRS" is a comparison against real SSRS output. A committed reference requires,
from the **same** SSRS run: the reference PDF, the exact parameters, the exact rows for every rendering
dataset, and the exact licensed font versions. Reconstructed/synthetic hydration proves the pipeline works
for those rows; it does not prove pixel equivalence. Certify with `rdl-pdf-layout-certification`
(page count/size exact, geometry within 0.5 pt, 144-DPI ≤0.5% pixel delta) and, for editable Word, with
`rdl-windows-word-fidelity`.

## Resources

- `references/value-formatting.md` — .NET format strings, culture, and no-format defaults SSRS applies.
- `references/sizing-and-growth.md` — CanGrow/CanShrink, container envelopes, fixed vs growing regions.
- `references/border-resolution.md` — per-edge borders, shared-edge coincidence, style precedence.
- `references/charts-and-axes.md` — axis scaling, tick/label formatting, auto-fit, no-data.
- `references/page-and-flow.md` — coordinate model, page header/footer, pagination, KeepTogether.
- `references/known-deviations.md` — currently-known gaps between this engine and SSRS to reconcile.
