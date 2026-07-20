# Supported operations and functions

What this service renders. The companion document, [LIMITATIONS.md](./LIMITATIONS.md), lists what it refuses
and why.

> **Accuracy language.** Entries below are **Implemented** (code path and tests exist) unless marked
> otherwise. *Smoke-tested* means a supplied client RDL renders in all modes with clean temporary storage.
> *SSRS-certified* means output passed comparison against a reference SSRS PDF using the exact matching rows,
> parameters, and fonts. **Nothing in this service is currently SSRS-certified** — see
> [Certification status](#certification-status).

## How support is decided

Support is **fail-closed**: the default status of every RDL construct is `REJECTED`. A construct renders only
if it has been explicitly classified, given a parser and a renderer, and covered by tests. An RDL using
anything else is refused with `UNSUPPORTED_FEATURE` rather than being approximated or silently dropped.

Every element and attribute carries one of three statuses:

| Status | Meaning | Behaviour |
| --- | --- | --- |
| `SUPPORTED` | Parsed and rendered. | Affects output. |
| `METADATA_ONLY` | Parsed and reportable, deliberately never acted on. | Ignored at render time. Never executed. |
| `REJECTED` | Not supported. | Blocks rendering with `UNSUPPORTED_FEATURE`. |

`METADATA_ONLY` is the important category for security: `DataSource`, `ConnectString`, `Query`,
`CommandText`, and `Code` are all parsed so `/v1/analyze` can *tell you they are there*, and are then
ignored. They are never executed. See [LIMITATIONS.md](./LIMITATIONS.md#why-data-sources-are-never-executed).

### The generated catalogue is the source of truth

This document is prose. The machine-readable catalogue is generated from Microsoft's published RDL schema:

```bash
npm run audit:schema   # writes tmp/output/rdl-2016-capability-catalogue.json
```

It currently classifies **695 declared names** — 691 elements and 4 attributes:

| Status | Count |
| --- | --- |
| `SUPPORTED` | 160 |
| `METADATA_ONLY` | 59 |
| `REJECTED` | 476 |

`POST /v1/analyze` classifies *your* RDL against this catalogue and returns every construct it uses, so you
never have to guess whether a report will render — ask the service.

Expressions use a **separate** catalogue (`EXPRESSION_FUNCTION_CAPABILITIES`) because XSD describes document
structure, not SSRS/VB runtime semantics. That distinction is also the origin of a real defect class — see
[LIMITATIONS.md](./LIMITATIONS.md#a-note-on-how-support-gaps-happen).

## Report structure

- `Report`, `ReportSections`, `ReportSection`, `Body`, `PageHeader`, `PageFooter`
- Page geometry: `PageWidth`, `PageHeight`, `TopMargin`, `RightMargin`, `BottomMargin`, `LeftMargin`
- `PrintOnFirstPage`, `PrintOnLastPage`
- Item geometry: `Top`, `Left`, `Width`, `Height`, `ZIndex`
- Page breaks, `KeepTogether`
- `df:DefaultFontFamily`

## Report items

| Item | Notes |
| --- | --- |
| `Textbox` | `Paragraphs` / `TextRuns`, `CanGrow`, `CanShrink`, per-run styling |
| `Tablix` | See [Tablix](#tablix) |
| `Rectangle` | As a body-level container and as a tablix-cell container |
| `Image` | **`Source=Embedded` only.** `Sizing`: `Fit`, `FitProportional`, `Clip`, `AutoSize` |
| `Line` | Body level only |
| `Chart` | See [Charts](#charts) |

`Image` with `Source=External` or `Source=Database` is rejected (`ImageSource:<source>`) — the service does
not fetch remote resources or read databases.

## Tablix

Flat tables, grouped tables, and matrices (cross-tabs).

- Static and **dynamic** row groups; nested row-header hierarchies (verified to seven columns)
- **Dynamic column groups** (matrix / cross-tab) with `TablixColumnHierarchy`, `TablixHeader`, `TablixCorner`
- Group **header/footer subtotal** rows
- **Recursive / parent** row groups (`Group/Parent`)
- `Group/Variables`, resolved as `Variables!Name.Value` in the group's scope
- Group row spans, merged cells (`ColSpan` / `RowSpan`), repeated header rows
- `HideDuplicates`, conditional visibility, conditional styles
- Filters and sorts
- Per-side borders, `KeepTogether`, single-column `ColumnSpacing`
- Safe `MarkupType=HTML` text normalization (normalized to text — not an HTML renderer)

**Performance gate.** Matrices, parent groups, and subtotal rows are gated behind detection in
`validation.js needsAdvancedMaterialization`. A tablix using none of them takes the original flat path and
materializes byte-identically. Do not remove those gates.

Cell content is restricted to `Textbox` (optionally wrapped in container `Rectangle`s). Any other item type
in a cell fails closed as `TablixCellContent:<Type>` — see
[LIMITATIONS.md](./LIMITATIONS.md#tablix-cell-content).

## Charts

Rendered as vector graphics from caller-supplied rows.

| RDL `Type` | Renders as |
| --- | --- |
| `Bar` | Horizontal bar |
| `Column` (or omitted `Type`) | Vertical column |
| `Line` | Line |
| `Area` | Area |
| `Scatter`, `Point` | Scatter |
| `Shape` | Pie, or doughnut when `Subtype` contains `Doughnut` |

`Stacked` and `PercentStacked` subtypes apply to bar, column, and area. Legends, titles, and axis labels are
supported. Any other chart type is rejected rather than approximated.

## Datasets, fields, and parameters

- `datasets` values are arrays of row objects keyed by **exact RDL `DataField` names** — not the field
  `Name`. The two are frequently different, and confusing them is the most common integration error. The
  service maps `DataField` → `Name` so expressions can use `Fields!Name.Value`.
- A missing dataset is `DATASET_MISSING`; a missing field is `FIELD_MISSING`.
- Parameter types validated: `String`, `Integer`, `Float`, `Boolean`, `DateTime` (ISO). Defaults, multi-value,
  and `Nullable` are honoured. A violation is `PARAMETER_INVALID`.

## Expressions

No `eval`, no `Function`, no VM, no dynamic code generation — expressions are parsed and interpreted. This is
a hard security invariant, not an implementation detail.

### Member access

`Fields!X.Value`, `Parameters!X.Value`, `Globals!X`, `User!X`, `ReportItems!X.Value`, `Variables!X.Value`,
plus `.IsMissing`, `.Count`, `.IsMultiValue`, and indexed `Parameters!X.Value(0)`.

Globals available: `PageNumber`, `TotalPages`, `ReportName`, `ExecutionTime`.

### Operators

| Class | Operators |
| --- | --- |
| Arithmetic | `+` `-` `*` `/` `\` (integer divide) `^` (power) `Mod` |
| Concatenation | `&` |
| Comparison | `=` `<>` `<` `>` `<=` `>=` `Like` |
| Logical | `And` `AndAlso` `Or` `OrElse` `Xor` `Not` |

### Functions (101)

All of the following are supported. Anything **not** on this list is rejected — including `Aggregate()` and
any `Code.*` call.

**Aggregates (13)** — `Avg`, `Count`, `CountDistinct`, `CountRows`, `First`, `Last`, `Max`, `Min`, `StDev`,
`StDevP`, `Sum`, `Var`, `VarP`

**Scope and position (5)** — `InScope`, `Level`, `Previous`, `RowNumber`, `RunningValue`

**Lookup (3)** — `Lookup`, `LookupSet`, `MultiLookup`

**Flow control (3)** — `IIF`, `Choose`, `Switch`

**Text (20)** — `Contains`, `EndsWith`, `InStr`, `InStrRev`, `Join`, `LCase`, `Left`, `Len`, `LTrim`, `Mid`,
`Replace`, `Right`, `RTrim`, `Space`, `Split`, `StartsWith`, `StrComp`, `StrReverse`, `Trim`, `UCase`

**Math (18)** — `Abs`, `Atan`, `Ceiling`, `Cos`, `Exp`, `Fix`, `Floor`, `Int`, `Log`, `Log10`, `Power`,
`RGB`, `Round`, `Sign`, `Sin`, `Sqrt`, `Tan`, `Truncate`

**Conversion (12)** — `CBool`, `CByte`, `CDate`, `CDbl`, `CDec`, `CInt`, `CLng`, `CSng`, `CStr`, `Hex`,
`Oct`, `Val`

**Date and time (17)** — `DateAdd`, `DateDiff`, `DatePart`, `DateSerial`, `DateValue`, `Day`, `Hour`,
`Minute`, `Month`, `MonthName`, `Now`, `Second`, `TimeValue`, `Today`, `Weekday`, `WeekdayName`, `Year`

**Formatting (5)** — `Format`, `FormatCurrency`, `FormatDateTime`, `FormatNumber`, `FormatPercent`

**Inspection (5)** — `IsArray`, `IsDate`, `IsError`, `IsNothing`, `IsNumeric`

### Format strings

`Format()` and `Style/Format` implement the .NET format engine (`src/rdl/format.js`): standard numeric
(`C`, `P1`, `F2`, `N0`, …), custom numeric (`#,##0.00`, section formats), and date/time patterns.

## VB semantics worth knowing

The interpreter reproduces SSRS/VB behaviour, including behaviour that surprises people:

- `Nothing` coerces to `0` in numeric context. An expression like
  `IIF(IsNothing(a) And IsNothing(b), "green", …)` can therefore colour a cell "green" for rows with no data.
  **SSRS does exactly the same thing.** If output looks wrong this way, the RDL logic is the cause, not the
  renderer.
- That coercion applies to **equality** too: `Nothing = 0` and `Nothing = ""` are both **True**, and `<>` is
  its exact negation. This is not a curiosity — reports drive conditional formatting from query-computed
  row-number fields (`=IIF(Fields!rn.Value = 0, "Solid", "None")` on a `TopBorder`, `=IIF(Fields!rn.Value = 0,
  False, True)` on `Hidden`), so a NULL row number still counts as `0` and still draws the rule. Treating it
  as unequal silently drops borders and blanks cells with no error anywhere.
- `And`/`Or` do not short-circuit; `AndAlso`/`OrElse` do.

## Output modes

| Mode | How it is produced | Page count |
| --- | --- | --- |
| `PDF` | Directly from the normalized RDL model. Selectable text. | Exact |
| `DOCX_EDITABLE` | Native OpenXML, generated directly — **not** converted from PDF. Real tables, real text. | Unknown (`null` / `X-Page-Count: unknown`) — Word paginates |
| `DOCX_FIXED_EDITABLE` | Canonical PDF parsed into positioned editable text boxes, Word shapes, lines, and images. No page screenshot. | Exact canonical count |
| `DOCX_VISUAL` | PDF rasterized at 300 DPI, one full-page floating image per Word page. | Exact |
| `XLSX` | Native Excel workbook. Each tablix is a block of styled cells (fills, borders, merges, fonts); numeric and date fields are written as **live typed values** with a translated number format, not text. Charts and the logo embed as images. | Not paginated (`null`) — Excel owns print layout |

`DOCX_EDITABLE` is editable but Word owns final layout, so its pagination will not match the PDF exactly.
`POST /v1/analyze` returns `structuredEditable` with native-DOCX drift risks and the
`structuredEditable.nativePageFragments.recommendation` value for the RDL shape. The optional
`docx.nativePageFragments` render flag keeps real Word tables but must be certified per report/data set.
Certified structured-DOCX profiles may be mounted and matched by `identity.definitionSha256`; auto-apply is
off by default. Profile files fail closed on duplicate/unsafe IDs, malformed match hashes, or unknown DOCX
rendering keys.
`DOCX_FIXED_EDITABLE` locks pagination and geometry while keeping every report text line editable; large
user edits can overflow their positioned text boxes. `DOCX_VISUAL` is the non-editable raster contract.

`XLSX` is a data-first export. By default every tablix is stacked as a row block in one worksheet with
autofit column widths; because the columns are shared, blocks with different column counts will not align —
Excel is a grid, not a page. Set `excel.sheetPerTablix: true` on the render request to put each tablix on its
own worksheet (`Table 1`, `Table 2`, …) with its own columns, and collect the title band, charts, and
free-form text onto a leading `Overview` sheet — best when the goal is filtering/pivoting each table.

Values that resolve to a number (including `=Format(Fields!X.Value, "N2")`, whose number is recovered behind
the format) are written as live numbers so they stay summable/pivotable; multi-run, conditional, and
genuinely textual cells stay text. Untrusted cell text is stored as typed strings that Excel never evaluates
as formulas — no apostrophe-escaping is applied because, unlike CSV, an XLSX string cell carries an explicit
type. The page footer (page numbers/dates) is omitted as it has no meaning in a continuous sheet.

## Certification status

| Report | Status |
| --- | --- |
| Combined Assurance 2016 client profile | Smoke-tested |
| Incident Dashboard | Smoke-tested |
| KRI Report | Smoke-tested |
| Internal 520-row stress fixture | Implemented + verified by `npm run verify:stress` |

No report is **SSRS-certified**. That requires the reference package described in `AGENTS.md` — the exact
SSRS reference PDF, its parameters, its rows, and its licensed font versions from the same run — and the
regression tolerances passing. Until then, use `npm run verify:reference`, which reports regression evidence
but will not mark output certified without `--exact-inputs`.
