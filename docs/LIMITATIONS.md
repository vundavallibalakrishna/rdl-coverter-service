# What is not supported, and why

The companion document, [SUPPORTED.md](./SUPPORTED.md), lists what renders. This one explains the refusals.

Every refusal here is deliberate. They fall into three groups:

1. **Security invariants** — supporting it would mean executing untrusted input. These will not change.
2. **Architectural boundaries** — the service has no database, network, or report server, by design.
3. **Not built yet** — no fundamental obstacle; simply not implemented, tested, and certified.

Knowing which group a refusal is in tells you whether asking for it is reasonable.

## Why refusal beats approximation

An unsupported construct is refused with `UNSUPPORTED_FEATURE`. It is never guessed at, approximated, or
silently dropped.

The reasoning: this service renders **compliance and risk reports**. A report that fails to render is a
visible, actionable problem. A report that renders but is subtly wrong — a missing row, a dropped total, a
cell that silently lost its content — is a *plausible-looking artifact that someone signs off on*. The second
failure is far more expensive than the first, and much harder to detect. So the default status of every
construct is `REJECTED`, and support is added only with a parser, a renderer, and tests.

Call `POST /v1/analyze` (or `converter.analyze()`) before rendering. It reports exactly which constructs an
RDL uses and whether they are supported, so refusals surface at integration time rather than in production.

---

## 1. Security invariants (permanent)

### Why data sources are never executed

`DataSource`, `ConnectString`, `Query`, `CommandText`, `QueryParameters`, and `IntegratedSecurity` are parsed
and classified `METADATA_ONLY`. The service reads them so `/v1/analyze` can report they exist. **It never
connects to anything, and never runs a query.**

The service accepts an uploaded RDL plus **caller-supplied rows**. The caller does data access, applies its
own authorization, and hands over the result set. That boundary is what makes it safe to accept RDL files
from users at all: an RDL is an executable document. If this service ran the queries inside one, anybody who
could upload a file could reach the database with the service's credentials, run arbitrary SQL and stored
procedures, and exfiltrate whatever it could see — a trivial SSRF-and-worse pivot into the private network.

Consequences, all intended:

- No live data. Rows come from you.
- No stored procedures, no `CommandType=StoredProcedure` execution.
- No `Source=External` / `Source=Database` images — no outbound fetches, no image-based SSRF.
- No XML external entities (XXE).

### Why custom code is never executed

| Construct | Status |
| --- | --- |
| `Code` (embedded VB) | `METADATA_ONLY` — reported, never run |
| `CodeModule`, `CodeModules` | `REJECTED` |
| `Code.*` calls in expressions | `REJECTED` |
| `CustomReportItem` | `REJECTED` |

`Code` in an RDL is arbitrary VB.NET authored by whoever produced the file. Running it would mean **remote
code execution as the service account**, from an uploaded document. There is no sandbox worth trusting here.

This is why expressions are parsed and interpreted rather than evaluated: no `eval`, no `Function`
constructor, no `vm` module, no dynamic code generation anywhere in the codebase. A report needing `Code.*`
must have that logic moved into the RDL expression surface or precomputed by the caller into a dataset field.

### `Aggregate()`

`Aggregate()` delegates to the *data provider's* aggregation — it means "ask the server to compute this."
There is no server here, so there is nothing correct to return. Rather than invent a plausible number for a
risk report, it fails closed. Use a supported aggregate (`Sum`, `Avg`, `Count`, …) or precompute the value.

---

## 2. Architectural boundaries (permanent)

### Subreports

`Subreport` is `REJECTED`. A subreport is a *reference to another report on a report server*, which the
renderer would have to resolve and execute with its own datasets. This service has no report server, no
catalogue, and no data access. Nothing it could do here would be correct.

**Workaround:** render the subreport as its own request and assemble the results caller-side.

### Maps and gauges

`Map*` (≈60 elements: shapefiles, spatial data, tile layers, projections, viewports) and `Gauge*`
(`RadialGauge`, `LinearGauge`, `GaugePanel`, scales, pointers) are `REJECTED`.

Maps additionally depend on spatial data sources and remote tile servers — the same execution and outbound
network problems as above. Gauges have no such obstacle; they are simply a large renderer that has not been
built. Both are refusals of scope, not of principle.

### Drillthrough, actions, document map

`ActionInfo`, `Drillthrough`, `ToggleItem`, `ToggleImage`, and `DocumentMapLabel` describe *interactive*
behaviour. PDF and Word are static artifacts. Interactivity is parsed where it is `METADATA_ONLY` and
otherwise rejected; toggled visibility renders in its statically-evaluated state.

---

## 3. Not built yet (may change)

These have no fundamental obstacle. They need a parser, a renderer, tests, and certification.

| Construct | Status | Note |
| --- | --- | --- |
| Gauges | `REJECTED` | Large renderer, no blocker |
| Chart types beyond bar/column/line/area/scatter/pie/doughnut | `REJECTED` | e.g. funnel, polar, range, candlestick |
| `Image` `Source=External`/`Database` | `REJECTED` | Blocked on security grounds above, but a caller-supplied image byte channel could be added |
| Non-`Textbox` tablix cell content | `REJECTED` | See below |

### Tablix cell content

A tablix cell may contain a `Textbox`, optionally wrapped in container `Rectangle`s. Anything else — an
`Image`, a `Line`, a nested `Tablix` — fails closed as `TablixCellContent:<Type>`.

### A note on how support gaps happen

This one is worth reading, because it explains a real defect class rather than a hypothetical one.

The capability catalogue classifies element **names**, globally. `Rectangle` was classified `SUPPORTED` —
which was true *at body level*, where `drawSimpleItem` handles it. But a tablix cell is its own rendering
context, and the cell materializer only understood `Textbox`. So a `Rectangle`-wrapped cell hit neither the
catalogue's refusal nor a renderer: `/v1/analyze` reported `compatible: true`, and the cell rendered **blank**.

That is precisely the silent-wrongness the fail-closed design exists to prevent, and it slipped through
because *support is contextual while the catalogue is not*. Access to the XSD did not help: XSD describes
where an element may legally appear, not whether this renderer draws it there.

The fix generalizes the lesson. `RENDERABLE_CELL_ITEMS` in `src/rdl/helpers.js` is an explicit allow-list of
what a cell renderer can actually draw, enforced in two places: at analyze time (`collectUnsupportedCell` in
`parser.js`, so `/v1/analyze` reports `TablixCellContent:<Type>` and `compatible: false`) and again at
materialization (which throws — defence in depth). Applying it immediately surfaced two more latent instances
of the same bug: `Image` and `Line` in cells, both of which had been silently rendering blank.

The general rule for contributors: **if a renderer consumes a restricted set of types, that set must be an
explicit allow-list, and anything outside it must be refused — not ignored.** A blank cell is a bug report
that never gets filed.

---

## Known behaviours that are not defects

Things that look wrong, are not, and have been investigated:

### Cells that render green with no data

VB coerces `Nothing` to `0`. An RDL guard like `IIF(IsNothing(a) And IsNothing(b), …)` combined with
non-short-circuiting `And` will colour empty rows as though they scored zero. **SSRS produces the same
output** for the same RDL and rows. This is report logic, not rendering — the fix belongs in the RDL.

### DOCX_EDITABLE pagination differs from the PDF

Word performs final layout, so page breaks land differently. `KeepTogether` and row spans are honoured via
`cantSplit`/`keepNext`, which keeps content coherent, but exact page parity is not achievable in an editable
reflowing document. Use `DOCX_FIXED_EDITABLE` for editable page-for-page canonical geometry, or
`DOCX_VISUAL` for a non-editable raster copy.

The experimental `docx.nativePageFragments` render option (legacy alias: `docxNativePageFragments`) keeps
content as native Word tables while splitting large tablixes at PDF-like page break estimates. It can improve
reports where Word otherwise compresses too many rows onto a page, but it can make row-span-heavy reports
worse. Enable it only for a certified RDL/data set after page-by-page Word export comparison.

`/v1/analyze` returns `structuredEditable` with the native-DOCX risk level, specific drift risks, and the
fragmentation recommendation for the RDL shape. That analysis is static: it does not replace rendering the
actual data set through Microsoft Word during release certification.

Certified structured-DOCX profiles can be mounted through `RDL_DOCX_PROFILE_PATH` and matched by the
`identity.definitionSha256` returned from `/v1/analyze`. Keep `RDL_DOCX_PROFILE_AUTO=false` until the profile
has passed page-by-page Word export certification for the exact report/data family. Auto-apply ignores
profiles where `certified` is not `true`; an uncertified candidate can still be selected explicitly for QA.
The service rejects malformed profile files instead of guessing: duplicate IDs, unsafe IDs, invalid
definition hashes, and unknown DOCX rendering keys return `CONFIG_INVALID` when a profile is requested or
auto-apply is enabled.

### DOCX_FIXED_EDITABLE is positioned, not reflowing

Every canonical PDF text line is an independent positioned Word text box. This keeps page count, page size,
headers, footers, lines, fills, images, and text anchors tied to the PDF, but large user edits can overflow a
box instead of pushing later rows or pages. Unsupported PDF operators and any text that cannot remain
editable fail with `UNSUPPORTED_FEATURE`; the renderer never hides raster text beneath editable overlays.

### Border thickness varies slightly between cells

Sub-pixel rasterization in the PDF viewer, not stroke width. Each grid line is stroked exactly once
(defer-and-merge in `pdf.js`), measured at 1.86px average versus SSRS's 1.91px. **The reference SSRS PDF
shows the same variance.** `RDL_BORDER_WIDTH_FLOOR_PT` exists to force a minimum stroke width, but it only
makes borders heavier rather than more uniform, so it defaults to `0` (off).

---

## Reference

Machine-readable, generated from Microsoft's published schema:

```bash
npm run audit:schema   # tmp/output/rdl-2016-capability-catalogue.json
```

Current classification of the 695 declared names: **160** `SUPPORTED`, **59** `METADATA_ONLY`, **476**
`REJECTED`. To check a specific RDL rather than the whole schema, use `POST /v1/analyze`.
