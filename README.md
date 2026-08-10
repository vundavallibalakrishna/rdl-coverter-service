# RDL Converter Service

Standalone, private-network RDL renderer. It takes an RDL definition plus **caller-supplied datasets** and
returns one PDF, DOCX, or XLSX.

It has no database, no persistent storage, and no Lumina dependency. It **never** executes RDL data sources,
SQL, stored procedures, custom report code, or external entities — see [Security model](#security-model).

Use it two ways, both supported and both backed by the identical pipeline:

- **[As an HTTP service](#option-a-run-as-an-http-service)** — run it, POST an RDL, get a file back.
- **[As a library](#option-b-use-as-a-library)** — `npm install`, call `createConverter()` in-process.

## Contents

- [Requirements](#requirements)
- [Quick start](#quick-start)
- [Option A: Run as an HTTP service](#option-a-run-as-an-http-service)
- [Option B: Use as a library](#option-b-use-as-a-library)
- [Which option should I use?](#which-option-should-i-use)
- [Output modes](#output-modes)
- [The request](#the-request)
- [Supplying subreports](./docs/SUBREPORTS.md)
- [Errors](#errors)
- [Configuration](#configuration)
- [Fonts](#fonts)
- [Security model](#security-model)
- [Runtime isolation](#runtime-isolation)
- [Docker](#docker)
- [What renders and what does not](#what-renders-and-what-does-not)
- [Working files](#working-files)
- [Testing and verification](#testing-and-verification)
- [Project layout](#project-layout)

## Requirements

| Requirement | Why |
| --- | --- |
| **Node.js ≥ 22** | ESM and native test runner; production may use Node 24 |
| **Poppler** (`pdftoppm`) | `DOCX_VISUAL`, and chart images in `DOCX_EDITABLE`/XLSX `REPORT`/`DATA` modes — rasterizes PDF/chart pages. `brew install poppler` / `apt install poppler-utils` |
| **Licensed fonts** | Production only. See [Fonts](#fonts) |

`GET /readyz` (or `readiness(config)`) reports whether all three are actually satisfied.

Microsoft Word is **not** installed or invoked by the service. All DOCX modes are generated as OOXML on
Linux. Microsoft Word for Windows is the sole `DOCX_EDITABLE` release-certification viewer.

## Quick start

```bash
npm install
RDL_STRICT_FONTS=false npm start     # http://localhost:7070
```

`RDL_STRICT_FONTS=false` substitutes fonts so you can render without licensed files. **Development only** —
output will not match SSRS metrics. See [Fonts](#fonts).

---

## Option A: Run as an HTTP service

```bash
npm start                            # node src/server.js
npm run dev                          # watch mode
npx rdl-converter-service            # if installed as a dependency
```

### Endpoints

| Method | Path | Purpose |
| --- | --- | --- |
| `GET` | `/healthz` | Process liveness. |
| `GET` | `/readyz` | Writable temp storage, required font variants, PDFKit, and Poppler readiness. `503` when not ready. |
| `GET` | `/` or `/test-ui` | Open end-to-end page for uploading one RDL plus its request JSON and downloading a selected output. |
| `POST` | `/v1/analyze` | Namespace, page settings, parameters, exact dataset fields, fonts, detected constructs, structured-DOCX drift risks and fail-closed errors. **Does not render.** |
| `POST` | `/v1/render` | One completed `PDF`, `DOCX_EDITABLE`, `DOCX_VISUAL`, or `XLSX` artifact. |

### Open test page

Start the service and open `http://localhost:7070/` in a browser. Select an `.rdl` file and a JSON request
containing its `parameters`, `datasets`, and optional bundled `subreports`, then choose the required output.
The page posts the files to the same `/v1/render` endpoint used by API clients and downloads the completed
artifact. Any `output` or `rdlBase64` property already present in the JSON is replaced by the page selection
and uploaded RDL respectively.

The page intentionally has no application authentication or role checks. It does not add CORS and remains
subject to this service's private-network-only deployment boundary. Do not expose the service directly to
the public internet.

### Analyze before you render

`/v1/analyze` answers "will this RDL render, and what exactly does it need from me?" without rendering it.
It returns `compatible`, every construct the RDL uses classified `SUPPORTED` / `METADATA_ONLY` / `REJECTED`,
`blockingErrors` for anything refused, and — most usefully — the exact `DataField` names each dataset
expects.

```bash
curl -X POST http://localhost:7070/v1/analyze \
  -H 'content-type: application/json' \
  --data "{\"rdlBase64\":\"$(base64 -i report.rdl)\"}"
```

### Render (JSON)

```bash
curl -X POST http://localhost:7070/v1/render \
  -H 'content-type: application/json' \
  --output report.pdf \
  --data @request.json
```

```jsonc
// request.json
{
  "rdlBase64": "PD94bWwg…",
  "output": "PDF",                      // PDF | DOCX_EDITABLE | DOCX_VISUAL | XLSX
  "outputFileName": "combined-assurance",
  "parameters": { "ReportYear": 2026 },
  "datasets": {
    "MainDataset": [                    // keys are exact RDL DataField names
      { "RiskName": "Vendor concentration", "Rating": 3 }
    ]
  }
}
```

For a parent RDL containing `Subreport` items, the request must also carry each child RDL and
invocation-scoped child rows. See [Supplying subreports](./docs/SUBREPORTS.md) for the complete HTTP JSON,
multipart, nested-grandchild, and validation contract.

### Render (multipart)

Avoids base64 inflation on large RDLs.

```bash
curl -X POST http://localhost:7070/v1/render \
  -F 'rdl=@report.rdl;type=application/xml' \
  -F 'request={"output":"PDF","parameters":{},"datasets":{}};type=application/json' \
  --output report.pdf
```

### Response headers

`Content-Type`, sanitized `Content-Disposition`, `Content-Length`, `X-Request-Id`, `X-Page-Count`,
`X-Render-Duration-Ms`; XLSX responses also include `X-Xlsx-Layout-Mode`. DOCX responses use
`X-Docx-Layout-Mode` plus `X-Docx-Editable-Text-Ratio`. The artifact is rendered completely before any
header is sent, so a `200` means a finished file.

`PDF`, `DOCX_EDITABLE`, and `DOCX_VISUAL` return numeric canonical PDF page counts. `DOCX_EDITABLE` reports
`X-Docx-Layout-Mode: windows-paged-editable`.

---

## Option B: Use as a library

Same parser, same validation, same worker isolation, same fail-closed guarantees — no HTTP hop.

### Install

The package is `private: true`, so it is **not** published to the public npm registry. Install it from the
repository or a tarball:

```bash
npm install git+ssh://git@your-host/rdl-converter-service.git
npm install github:your-org/rdl-converter-service#v0.1.0
```

```bash
npm pack                                       # -> rdl-converter-service-0.1.0.tgz
npm install /path/to/rdl-converter-service-0.1.0.tgz
```

To publish to an **internal** registry, remove `"private": true` from `package.json` and add a
`publishConfig.registry`. It is deliberately left in so a stray `npm publish` cannot push a client-data-
handling converter to the public registry. The `files` allow-list ships only `src/`, `docs/`, and the
top-level docs; `test/` and anything under `tmp/` — including client samples — never enter the tarball.

### Render a report

```js
import { createConverter } from 'rdl-converter-service';
import fs from 'node:fs/promises';

const converter = await createConverter();

try {
  const rendered = await converter.render({
    rdl: await fs.readFile('report.rdl'),   // Buffer | Uint8Array | string
    output: 'PDF',                          // PDF | DOCX_EDITABLE | DOCX_VISUAL | XLSX
    parameters: { ReportYear: 2026 },
    datasets: {
      MainDataset: [{ RiskName: 'Vendor concentration', Rating: 3 }],
    },
  });

  await fs.writeFile('report.pdf', rendered.buffer);
  console.log(rendered.pageCount, rendered.mimeType, rendered.totalRows);
} finally {
  await converter.close();
}
```

The library accepts the same `subreports` object as the HTTP API. Child RDL bytes are base64 inside that
object even though the library accepts the parent `rdl` directly as a `Buffer`. See
[Supplying subreports](./docs/SUBREPORTS.md#library-example).

### Check compatibility first

```js
const analysis = converter.analyze(await fs.readFile('report.rdl'));
if (!analysis.compatible) {
  console.error(analysis.blockingErrors);   // [{ feature: 'TablixCellContent:Image', … }]
}
```

`analyze()` is pure and synchronous — no worker, no temp files.

### Handle errors

```js
import { ServiceError } from 'rdl-converter-service';

try {
  await converter.render({ rdl, output: 'PDF', datasets });
} catch (error) {
  if (error instanceof ServiceError) {
    console.error(error.code, error.statusCode, error.details);   // stable codes
  } else throw error;
}
```

### API

| Export | Description |
| --- | --- |
| `createConverter({ config, env })` | Creates a converter. Resolves temp storage and owns a worker pool. |
| `converter.render({ rdl, output, parameters, datasets, subreports, signal })` | → `{ buffer, mimeType, extension, pageCount, size, totalRows }` |
| `converter.analyze(rdl)` | → compatibility report. Synchronous. |
| `converter.close()` | Terminates in-flight workers. Call on shutdown. |
| `converter.config` | The resolved, frozen config. |
| `buildApp({ config, logger })` | The Fastify instance — embed the HTTP service in your own server. |
| `analyzeRdl(rdl, limits)` / `parseRdl(rdl, limits)` | Lower-level, no converter needed. |
| `readiness(config)` / `checkFonts(config, families)` | Health checks for your own probes. |
| `ServiceError` / `toServiceError(error)` | Stable error type. |
| `loadConfig(env)`, `OUTPUTS`, `RenderRunner`, `sanitizedFilename` | Advanced use. |

**Use `createConverter()`, not `new RenderRunner(config)`.** The runner assumes `config.tempRoot` already
exists with `0700` permissions (the HTTP app creates it at boot), so a bare runner fails on first render.
`createConverter()` performs that setup.

Configure it in-process without environment variables:

```js
const converter = await createConverter({
  config: { maxConcurrency: 4, renderTimeoutMs: 60_000, fontDir: '/srv/fonts' },
});
```

Cancel a long render with an `AbortSignal`:

```js
await converter.render({ rdl, output: 'PDF', datasets, signal: AbortSignal.timeout(30_000) });
```

### Which option should I use?

| | HTTP service | Library |
| --- | --- | --- |
| Language | Any | Node.js 22.x only |
| Isolation | Separate process **and** host; independently restartable | Worker per render; shares your host |
| Resource limits | Contained by the service's own memory/CPU budget | Renders compete with your app |
| Operations | Another deployment to run | Nothing extra |
| Overhead | Network + base64/multipart transfer | None |

Rule of thumb: **library** for a Node.js caller that already trusts the RDLs it renders; **HTTP service** for
anything multi-language, multi-tenant, or accepting user-uploaded RDLs — the extra blast-radius isolation is
worth the deployment.

Both paths run each render in a forked worker with the same timeout, heap cap, concurrency admission, and
guaranteed temp cleanup. That isolation is a property of the pipeline, not of the server.

---

## Output modes

| Mode | How it is produced | Text | Page count |
| --- | --- | --- | --- |
| `PDF` | Directly from the normalized RDL model. | Selectable | Exact |
| `DOCX_EDITABLE` | Native OpenXML built from the canonical PDF renderer's resolved layout trace. It is not a screenshot or an external PDF conversion. | Editable | Canonical PDF count |
| `DOCX_VISUAL` | Renders PDF, rasterizes every page at 300 DPI, one full-page floating image per Word page. | Images | Exact |
| `XLSX` | Native Excel workbook. The default `REPORT` layout creates one native-cell worksheet per explicit RDL section; `DATA` preserves the legacy stacked/per-tablix export. | Live cells | Not paginated (`null`) |

`DOCX_EDITABLE` generates the canonical PDF internally, records its resolved page geometry, and constructs
one native fixed-layout Word section for each PDF page. Tables and text remain editable; PDF-measured line
breaks, grids, row heights, borders, headers, footers, and fonts are materialized explicitly. Declared RDL
images and charts may remain pictures. The result is designed and certified only for Microsoft Word on
Windows; editing may change later pagination. Unsupported Word geometry fails closed.

The former continuous renderer, document profiles, and native-fragment switches have been removed.
`docx.nativePageFragments`, `docxNativePageFragments`, `docx.profile`, and their former environment settings
are rejected with `RDL_INVALID`.

Choose `DOCX_VISUAL` when editability is unnecessary and a raster page image is acceptable. Choose the
default XLSX `REPORT` layout for an editable, PDF-styled workbook without PDF pagination. Set
`"excel": { "layoutMode": "DATA" }` for the legacy data-first workbook.

In XLSX `REPORT` mode, report text, tablixes, rectangles, lines, fills, and borders remain native editable
cells. Declared RDL images and visible charts are anchored pictures in their resolved report regions;
side-by-side chart peers retain their coordinate relationship. Chart pictures are intentionally not native
Excel chart objects, while the surrounding workbook remains editable.

## The request

| Field | Required | Notes |
| --- | --- | --- |
| `rdlBase64` | JSON only | The RDL. Multipart uses the `rdl` file part instead. |
| `output` | ✅ | `PDF` \| `DOCX_EDITABLE` \| `DOCX_VISUAL` \| `XLSX` |
| `datasets` | ✅ | Object of `datasetName` → array of row objects. |
| `parameters` | — | Validated against the RDL's declared types and defaults. |
| `subreports` | — | Render-time bundle of child `rdlBase64` definitions and invocation-scoped parameter/dataset instances. Supported for `PDF`, `DOCX_EDITABLE`, `DOCX_VISUAL`, and XLSX `REPORT`; see [Supplying subreports](./docs/SUBREPORTS.md). |
| `pagination.continuationMarkers` | — | `PDF` and `DOCX_EDITABLE`. When `true`, places “Continued from previous page” above the next table fragment for renderer-confirmed logical-row continuations. |
| `excel.layoutMode` | — | `XLSX` only, case-insensitive. `REPORT` (default) or legacy `DATA`. |
| `excel.sheetPerTablix` | — | `XLSX` DATA mode only. Existing `true` requests without `layoutMode` continue to select DATA automatically. |
| `outputFileName` | — | Sanitized for `Content-Disposition`; also `Globals!ReportName`. |

> **`datasets` rows are keyed by the exact RDL `DataField` name — not the field `Name`.**
> The two are often different, and confusing them is the most common integration error. The service maps
> `DataField` → `Name` so expressions can use `Fields!Name.Value`. `POST /v1/analyze` returns the exact
> `DataField` names each dataset expects — use it rather than guessing.

For a parent containing `<ReportName>/Shared/Child</ReportName>`, supply every concrete invocation:

```json
{
  "subreports": {
    "/Shared/Child": {
      "rdlBase64": "PD94bWwg...",
      "instances": [
        {
          "parameters": { "EntityID": 42 },
          "datasets": { "ChildData": [{ "Entity ID": 42, "Label": "Resolved child row" }] }
        }
      ]
    }
  }
}
```

Child rows follow the same exact-`DataField` rule. Missing or duplicate invocation parameters, unused
definitions, cycles, excessive nesting, unsupported child body items, and absent child datasets fail closed.
The service never executes the child RDL query or resolves `ReportName` outside this request.

The short example above shows the shape only. The complete guide explains parameter type canonicalization,
empty child datasets, reusing signatures, multipart calls, library calls, and child-to-grandchild reports:
[Supplying subreports](./docs/SUBREPORTS.md).

## Errors

JSON body: `{ "error": { "code", "message", "details" } }`, with stable codes:

| Code | HTTP | Meaning |
| --- | --- | --- |
| `RDL_INVALID` | 400 / 413 | Malformed, empty, or oversized RDL/request. |
| `UNSUPPORTED_FEATURE` | 400 | Fail-closed refusal. See [LIMITATIONS.md](./docs/LIMITATIONS.md). |
| `PARAMETER_INVALID` | 400 | Parameter missing or wrong type. |
| `DATASET_MISSING` | 400 | The RDL needs a dataset you did not supply. |
| `FIELD_MISSING` | 400 | Rows are missing a required `DataField`. |
| `FONT_MISSING` | 503 | A declared font has no file on the render host. Characters an *installed* font cannot draw are substituted instead — see [Fonts](#fonts). |
| `BUSY` | 503 | At capacity. Includes `Retry-After`. |
| `RENDER_TIMEOUT` | 504 | Exceeded `RDL_RENDER_TIMEOUT_MS`. |
| `RENDER_FAILED` | 500 / 499 | Worker died, or client disconnected mid-render. The response message is scrubbed by design; the underlying exception is logged server-side as the `diagnostic` field of the `RDL request failed` log line. |

Messages are deliberately free of report content — see [Security model](#security-model).

## Configuration

Environment variables (see `.env.example`). Library callers can pass the same values via
`createConverter({ config })`.

| Variable | Default | Purpose |
| --- | --- | --- |
| `PORT` | `7070` | HTTP port. |
| `HOST` | `0.0.0.0` | Bind address. |
| `RDL_TEMP_ROOT` | `<tmpdir>/rdl-converter` | Private scratch root. Created `0700`. |
| `RDL_FONT_DIR` | `<cwd>/fonts` | Licensed font directory. |
| `RDL_STRICT_FONTS` | `true` | `false` substitutes missing fonts. **Development only.** |
| `RDL_ALLOW_COMPATIBLE_FONT_FALLBACKS` | `false` in strict mode; `true` in non-strict mode | Permit explicitly catalogued, glyph-validated font fallbacks. This is never an unrestricted system-font substitution. |
| `RDL_MAX_RDL_BYTES` | `10485760` (10 MB) | Max RDL size. |
| `RDL_MAX_REQUEST_BYTES` | `26214400` (25 MB) | Max request size. |
| `RDL_MAX_ROWS` | `100000` | Max rows per render. |
| `RDL_MAX_CONCURRENCY` | `2` | Concurrent renders. Beyond this → `BUSY`. |
| `RDL_RENDER_TIMEOUT_MS` | `120000` | Hard render deadline. |
| `RDL_WORKER_MEMORY_MB` | `512` | Baseline V8 heap cap for an ordinary worker. |
| `RDL_WORKER_MEMORY_MAX_MB` | `4096` | Hard cap for deterministic PDF/page-locked DOCX workload-based heap scaling. Never lower than the baseline. |
| `RDL_MAX_XML_NODES` | `250000` | XML expansion guard. |
| `RDL_MAX_XML_DEPTH` | `256` | XML nesting guard. |
| `RDL_PDFTOPPM_PATH` | `pdftoppm` | Poppler binary for `DOCX_VISUAL`. |
| `RDL_PDF_LAYOUT_OPTIMIZATIONS` | `true` | Generic PDF measurement-cache optimization. Set `false` for an immediate rollback to the v0.1.0 computation path. |
| `RDL_EXPRESSION_PLAN_CACHE` | `true` | Bounded structural expression-plan cache shared by all renderers. Set `false` to reparse every expression evaluation. |
| `RDL_PDF_FONT_SELECTION_CACHE` | `true` | Bounded request-scoped cache for repeated PDF font and glyph-coverage selections. Set `false` to resolve every text run independently. |
| `RDL_BORDER_WIDTH_FLOOR_PT` | `0` | Minimum PDF border stroke, in points. `0` honours the RDL exactly. |
| `LOG_LEVEL` | `info` | Fastify log level. |
| `RDL_SAMPLES_DIR` | `<repo>/tmp` | Dev only. Where client samples live for tests and smoke scripts. See [Client samples](#client-samples). |

### Production launchers

The repository includes production launchers with strict fonts, bounded input/XML limits, full console
telemetry, one concurrent render, a five-minute render deadline, and workload-scaled worker heaps from
512 MB up to 8192 MB. Both launchers automatically load `.env.production` from the repository root when
that file exists:

```bash
./start-production.sh
```

```bat
start-production.bat
```

Create it from the tracked template, then replace the host-specific paths and large-report settings:

```bash
cp .env.example .env.production
```

```bat
copy .env.example .env.production
```

Configuration precedence is: existing process/service environment, `.env.production`, then launcher
defaults. Operators can therefore use IIS/systemd/Docker settings as authoritative overrides without
editing the launcher or environment file. Set `RDL_ENV_FILE` to use a file in another location. On Windows,
for example:

```bat
set "RDL_FONT_DIR=C:\rdl-fonts"
set "RDL_PDFTOPPM_PATH=C:\poppler\Library\bin\pdftoppm.exe"
set "RDL_WORKER_MEMORY_MAX_MB=12288"
start-production.bat
```

The main server intentionally starts without `--max-old-space-size`: every isolated render worker receives
its own workload-estimated heap bound from `RDL_WORKER_MEMORY_MB` and `RDL_WORKER_MEMORY_MAX_MB`.

### Render telemetry

At the default `LOG_LEVEL=info`, every render emits bounded JSON phase events to stdout with
`event: "render.phase"` and the Fastify request ID. The HTTP layer, runner, and isolated worker report:

- one startup `runtime.profile` event with the OS/Node architecture, available parallelism, physical memory,
  concurrency, worker-heap bounds, and timeout used by that process;
- request decoding and structural input sizes;
- temporary-storage preparation and input-write time;
- worker-memory estimation, selected heap limit, worker startup, and cleanup;
- input reads, JSON decoding, RDL parsing, bundled-subreport resolution, font checks, and validation;
- requested renderer-module loading (unused PDF, DOCX, and XLSX dependency graphs are not loaded);
- renderer start/completion, artifact write/read, page/sheet/row counts, and output size;
- for direct PDF output: canonical body layout (including aggregate tablix materialization, setup, initial
  measurement, and drawing time), page header/footer bands, PDFKit serialization, and final PDF validation;
- for page-locked editable DOCX: its canonical PDF phases, trace validation, font loading, bounded page
  construction progress, OOXML packing, font-variant packaging, and internal-artifact cleanup;
- elapsed and phase duration, process RSS/heap/external memory, CPU time, and event-loop utilization;
- timeout, abort, worker-exit, validation, rendering, and cleanup failure phases using stable error codes.
  Fatal worker exits are classified as `V8_HEAP_OUT_OF_MEMORY`, `NATIVE_RUNTIME_ABORT`, `PROCESS_ABORT`, or
  `WORKER_EXIT`; raw worker stderr is bounded in memory and never written to logs or returned to callers.

The terminal `RDL rendering completed` event remains the request summary. Fastify's subsequent
`request completed` event includes response transmission time, so comparing the two separates server-side
rendering from a slow client/network download.

Telemetry is intentionally structural and bounded. It never contains RDL XML, queries, parameter values,
dataset values, temporary paths, output binary content, request/response bodies, or authorization headers. Raising
`LOG_LEVEL` changes log filtering only; it does not weaken that content boundary.

`RDL_BORDER_WIDTH_FLOOR_PT` exists for a reported "uneven borders" symptom that turned out to be viewer
sub-pixel rasterization, not stroke width. A floor only makes borders *heavier*, not more uniform, so it is
opt-in and off by default.

## Fonts

Font metrics drive layout. Wrong fonts means wrong pagination, so production is **strict by default**: a
missing font is `FONT_MISSING` (503) rather than a silent substitution that quietly shifts every page break.

Two different conditions are treated differently, because they are not the same problem:

| Condition | Behaviour |
| --- | --- |
| The declared family has **no file** on this host | Strict mode fails closed (`FONT_MISSING`, 503). Substituting would change the advance widths of every run, and therefore every page break. |
| The declared family **is installed but has no glyph** for some character | The run is drawn in an installed font that does cover it. Never fails. |

The second case is a property of a few characters, not of the document: Arial has no `✓` (U+2713), `✗`
(U+2717) or `☹` (U+2639), and no Latin family covers CJK or Indic. The declared font still draws every
other run at its own metrics, so a single character in a single cell must not cost the whole export. No
flag gates this — the alternative is a 503 for a report that renders correctly everywhere else.

- Coverage stand-ins are tried metric-compatible first (Liberation Sans has Arial's advance widths, so when
  it covers the run the page breaks are identical to SSRS), then widest-coverage: Segoe UI Symbol, Noto
  Sans Symbols 2, Arial Unicode MS, Lucida Sans Unicode, Microsoft Sans Serif, DejaVu Sans, Noto Sans.
  A candidate is used only when it covers **the whole run** — a face covering the missing character but not
  the surrounding text would trade one unrenderable character for an unrenderable run.
- A candidate must also be **embeddable**, not merely cover the text. PDFKit subsets an embedded TrueType
  font by decoding each glyph it used; colour fonts (COLR/CBDT/sbix — Segoe UI Emoji is one) return glyphs
  with no decoder and abort the render with `glyph._decode is not a function` at the very end, as an opaque
  `RENDER_FAILED` 500. fontkit reports coverage for those glyphs happily, so coverage alone is not a safe
  test. Colour fonts are therefore skipped, including when the report declares one directly.
- When nothing on the host covers the run, the declared font is kept and only the characters it lacks come
  out as `.notdef`, as SSRS renders them. The export still succeeds.
- Substitutions are reported, never silent: the `X-Font-Substitutions` response header and the
  `fontSubstitutions` field of the render log line list `requested => substituted (reason, runs)`.
- To widen coverage without a code change, drop a single-face file into `RDL_FONT_DIR`:
  `NotoSansSymbols2-Regular.ttf`, `NotoSansCJK-Regular.ttf`, or `NotoSansDevanagari-Regular.ttf`. Font
  collections (`.ttc`, e.g. Windows `Nirmala.ttc` / `msgothic.ttc`) are **not** usable — a collection needs
  a face name that the path-based font handoff to PDFKit cannot carry.
- Licensed fonts are **never committed to the repository or baked into the image** (see `AGENTS.md`). Mount
  them at runtime into `RDL_FONT_DIR` (`/app/fonts` in the container).
- Reports commonly need Arial, Times New Roman, Segoe UI, and Segoe UI Symbol. The exact consumed
  regular/bold/italic/bold-italic faces must be mounted.
- `Segoe UI Emoji` may use a mounted `NotoEmoji-Regular.ttf` as an explicitly enabled compatible PDF
  fallback for the *absent-family* case. The renderer validates the actual glyphs before layout and
  embedding; supplementary-plane emoji never fall through to Helvetica. Enable this in strict mode only
  with `RDL_ALLOW_COMPATIBLE_FONT_FALLBACKS=true`.
- `GET /readyz` reports exactly which variants are missing. Note it checks file *presence*, not glyph
  coverage — a host can be ready and still substitute for individual characters.
- `RDL_STRICT_FONTS=false` substitutes them for local development. Never in production.

## Security model

The service treats every RDL as untrusted input, because an RDL is an executable document.

- **No data source execution.** `DataSource`, `ConnectString`, `Query`, `CommandText` are parsed so
  `/v1/analyze` can report them, then ignored. No SQL, no stored procedures, no connections. Rows come from
  the caller, who has already applied its own authorization.
- **No code execution.** Embedded `Code` is reported, never run. `CodeModules` and unknown `Code.*` calls
  are rejected. `Code.CalculateColor(yValue, xValue)` is the sole allowlisted compatibility call and
  dispatches to a fixed service-owned 5×5 heat-map function, never to report VB.
- **No dynamic evaluation.** Expressions are parsed and interpreted. No `eval`, no `Function`, no `vm`, no
  code generation anywhere in the codebase.
- **No outbound fetches.** No external entities (XXE), no `Source=External` images.
- **No content logging.** RDL contents, queries, parameters, row values, and rendered artifacts are never
  logged. Logs carry sizes, counts, durations, and request IDs only.
- **Fail-closed by default.** Unclassified constructs are refused rather than approximated.
- **Unauthenticated by design.** It is a private-network component. Never expose it publicly without adding
  authentication, TLS, and network allowlisting at the ingress.

Full reasoning: [docs/LIMITATIONS.md](./docs/LIMITATIONS.md).

## Runtime isolation

Every render gets a `0700` directory with `0600` files, runs in a forked child process with a bounded V8
heap, completes the whole artifact before response headers are sent, and has its directory removed on
success, failure, timeout, disconnect, **and** shutdown.

Defaults: 10 MB per RDL, 25 MB per request, 100,000 rows, 120 s, two concurrent workers, a 512 MB
baseline heap, and a 4096 MB hard per-worker cap for exceptionally wide or text-heavy PDF/page-locked
DOCX workloads. The cap is a V8 limit, not a reservation; ordinary renders remain at the baseline.
For a large page-locked DOCX workload on a sufficiently provisioned Windows host, operators may raise
`RDL_WORKER_MEMORY_MB` and `RDL_WORKER_MEMORY_MAX_MB`; restart the service after changing either value and
use the `memory-estimated` plus fatal-category telemetry to verify the result. On hosts that cannot safely
accommodate two multi-gigabyte workers, reduce `RDL_MAX_CONCURRENCY` rather than lowering a worker below the
memory required to construct its complete native OOXML package.
For a low-core Windows VM serving large PDF/DOCX requests, start with `RDL_MAX_CONCURRENCY=1`. Compare one
isolated render before increasing it: two simultaneous CPU-bound layouts can increase latency while a large
DOCX also raises the host working set by several gigabytes.

## Docker

```bash
RDL_FONT_HOST_DIR="/path/to/licensed-font-directory" docker compose up --build
```

The Compose profile drops Linux capabilities, uses a read-only root filesystem with a bounded tmpfs, and
binds to loopback only. Production ingress must add TLS, network/IP allowlisting, request-size enforcement,
and a container memory limit.

## What renders and what does not

Two documents, kept deliberately separate:

- **[docs/SUPPORTED.md](./docs/SUPPORTED.md)** — every supported construct: report items, tablix features
  (including matrices, subtotals, and recursive groups), charts, parameters, operators, and all
  **101 expression functions**.
- **[docs/LIMITATIONS.md](./docs/LIMITATIONS.md)** — what is refused **and why**, split into permanent
  security invariants, architectural boundaries, and not-built-yet. Also covers known-but-correct
  behaviours (`Nothing`-coerces-to-`0`, DOCX pagination, border sub-pixel variance).

The short version: a table-report subset, fail-closed. Textbox, embedded image, rectangle, line, chart, and
normalized tablix reports; exact `DataField` mapping; parameters, visibility, conditional styles, filters,
sorts, nested and dynamic groups, merged cells, repeated headers, per-side borders, duplicate suppression,
safe HTML-to-text normalization, z-order, keep-together, page settings and breaks, and a catalogued safe
expression subset. Expression-capable paragraph spacing and line height are resolved per row, and
`TextRun/Value@EvaluationMode="Constant"` preserves a leading `=` as literal text. Caller-bundled,
invocation-scoped tablix subreports render in PDF, editable DOCX, visual DOCX, and XLSX `REPORT`; unresolved
subreports and subreports in XLSX `DATA` remain fail-closed. Maps, gauges, custom code, and non-embedded
images are rejected.

To ask the service instead of reading docs, `POST /v1/analyze`. To dump the whole schema catalogue:

```bash
npm run audit:schema   # tmp/rdl-2016-capability-catalogue.json
```

Of the 695 declared names in Microsoft's 2016 schema (691 elements, 4 attributes): **169** `SUPPORTED`,
**62** `METADATA_ONLY`, **464** `REJECTED`. The default is `REJECTED`.

> **Accuracy language.** *Implemented* = the code path and tests exist. *Smoke-tested* = the supplied RDL
> renders in all modes with clean temporary storage. *SSRS-certified* = output passed comparison against the
> reference SSRS PDF using the exact matching rows, parameters, and fonts. **Nothing here is currently
> SSRS-certified.** Do not describe output as pixel-perfect or certified until the reference package has been
> supplied and the `AGENTS.md` regression tolerances have passed.

## Working files

Everything generated — renders, smoke output, verifier reports, rasterized pages — is written under `tmp/`,
which is git-ignored in full. Nothing a render produces should ever reach version control: artifacts embed
client report content and row values.

`.gitignore` also blanket-ignores `*.pdf`, `*.docx`, `*.png`, and font files, so an artifact written outside
`tmp/` is still caught. `test/fixtures/**` is the one deliberate exception — those are synthetic and the
suite needs them.

### Client samples

Client RDLs, request fixtures, and row data are **client property and are never committed**. They carry real
report definitions, queries, and row values, so they live directly in **`tmp/`** — outside version control —
not in a report/run/format subfolder.

A fresh clone therefore legitimately has no samples, and that is not a broken checkout. Tests and scripts
that need them **skip** rather than fail:

```
✔ resolves CSS/X11 colour names to hex so PDFKit fills render
﹣ the incident dashboard with charts passes the capability gate (skipped)
```

To run the sample-backed tests, put the client files directly in `tmp/`, or point at them elsewhere:

```bash
RDL_SAMPLES_DIR="/secure/path/to/client-samples" npm test
```

Both tests and scripts resolve this through `scripts/lib/samples.js` — the single source of truth for the
location. Do not hard-code sample paths.

## Testing and verification

```bash
npm test                                    # required for every change
npm run test:coverage
```

For parser or rendering changes, also:

```bash
npm run smoke:sample -- "/path/to/report.rdl"
RDL_FONT_HOST_DIR="/path/to/licensed-font-directory" npm run smoke:docker -- "/path/to/report.rdl"
npm run audit:schema
npm run verify:stress -- --require-pass
npm run verify:stress:nested -- --require-pass
npm run verify:reference -- "/path/to/report.rdl" "/path/to/reference.pdf" "/path/to/hydration.json"
npm run certify:windows-word -- "/path/to/report.rdl" "/path/to/request.json" "/path/to/reference.pdf"
```

Inspect representative PDF/DOCX output. **An HTTP `200` is not verification.**

- **`verify:stress`** renders a deterministic 470-row table as direct PDF and page-locked editable DOCX.
  Its canonical PDF must be 30–40 pages. The gate requires exact-once row and
  oversized-cell markers, native tables, repeated headers, interlocking horizontal/vertical merges, and a
  structurally closed bottom border on every explicit table fragment. It also raster-checks the complete
  bottom rule on every table-bearing PDF and DOCX page, including pages created by a cell taller than the
  page. Every text/field run must explicitly name its RDL font; the fixture materially mixes Arial, Times New
  Roman, and Segoe UI with bold, italic, underline, colour, conditional fills, per-side/dashed borders,
  wrapped text, nested grouping, and a final explicit page break. Artifacts and the JSON audit are written
  directly under `tmp/`. This certifies renderer mechanics — it does not claim SSRS or Microsoft Word visual
  certification.
- **`verify:stress:nested`** runs the same structural, traceability, font-grounding, package-integrity, and
  all-page raster-border gates against a deterministic five-level hierarchy. It combines repeatable
  multi-band headers, nested row-header spans, group headers/footers and scoped totals, conditional
  visibility/styles, explicit group page breaks, and three independently page-taller editable cells. The
  canonical PDF is constrained to 30–40 pages. Every emitted PDF-derived page fragment must carry a
  continuous full-grid closing rule. Artifacts are direct children of `tmp/`.
- **`verify:reference`** renders a hydrated PDF, compares page count and dimensions, checks text anchors,
  rasterizes the first three pages at 144 DPI, and writes a comparison report. Add `--exact-inputs
  --require-match` **only** when the hydration and font versions come from the exact SSRS reference run.
- **`certify:windows-word`** builds the page-locked DOCX, invokes the bundled PowerShell/COM harness on a
  Windows QA host, forces Word pagination, exports to PDF without updating fields, audits the package, and
  compares it with the canonical PDF at 144 DPI. Release certification requires exact page count and
  dimensions, identical displayed text and ordering, geometry within 0.5 pt, and no page with more than
  0.5% of pixels differing by over 16 colour levels. Word for Mac, LibreOffice, Google Docs, and browser
  previews are not certification authorities.

## Project layout

```
src/
  index.js          Public library entry point (createConverter, buildApp, analyzeRdl…)
  server.js         HTTP entry point (npm start)
  app.js            Fastify app: routes, error mapping
  config.js         Environment → frozen config
  request.js        Transport parsing (multipart / base64), filename sanitization
  readiness.js      Temp storage, font, PDFKit, Poppler probes
  rdl/
    parser.js       RDL XML → normalized model; analyze-time fail-closed guards
    validation.js   Parameters, datasets, tablix materialization
    capabilities.js Schema capability catalogue (SUPPORTED/METADATA_ONLY/REJECTED)
    expression.js   Expression parser + interpreter (no eval)
    functions/      The 101-function registry
    format.js       .NET format-string engine
  render/
    pdf.js          PDF renderer and canonical layout trace producer
    layoutTrace.js  Deterministic point-based PDF layout manifest
    docx.js         Windows-paged editable DOCX entry point
    pagedDocx.js    Trace-driven native Word section/table renderer
    windowsWordCompatibility.js  Word platform limits and request policy
    visualDocx.js   Rasterized full-page DOCX
    excel.js        Native XLSX renderer
    chart.js        Vector charts
    fonts.js        Font resolution + strict-font enforcement
  worker/
    runner.js       Concurrency admission, timeouts, temp lifecycle
    renderWorker.js Forked worker: parse → font check → validate → render
scripts/
  lib/samples.js    Resolves client samples (tmp, or RDL_SAMPLES_DIR)
  …                 Smoke, stress, reference, and schema-audit tooling
test/               Node test-runner suite
tmp/                All generated output AND client samples. Git-ignored in full.
```

Contributors: read `AGENTS.md` first. It carries the authoritative architecture, security invariants,
supported RDL boundary, and verification workflow.
