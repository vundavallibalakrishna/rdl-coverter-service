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
| **Node.js ≥ 22** | ESM, native test runner |
| **Poppler** (`pdftoppm`) | `DOCX_VISUAL`, and chart images in `DOCX_EDITABLE`/XLSX `DATA` mode — rasterizes PDF/chart pages. `brew install poppler` / `apt install poppler-utils` |
| **Licensed fonts** | Production only. See [Fonts](#fonts) |

`GET /readyz` (or `readiness(config)`) reports whether all three are actually satisfied.

Microsoft Word is **not** installed or invoked by the service. All DOCX modes are generated as OOXML on
Linux; Word for Mac/Windows is used only as the authoritative release-certification viewer.

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
`X-Docx-Layout-Mode` plus
`X-Docx-Editable-Text-Ratio`. Structured DOCX responses also include
`X-Docx-Native-Page-Fragments`; if a profile was applied they include `X-Docx-Profile-Id` and
`X-Docx-Profile-Certified`. The artifact is rendered completely before any header is sent, so a `200` means
a finished file.

`DOCX_EDITABLE` returns `X-Page-Count: unknown` — Word performs its own final pagination. `PDF`,
`DOCX_VISUAL` returns exact page counts.

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
| Language | Any | Node.js ≥ 22 only |
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
| `DOCX_EDITABLE` | Native OpenXML, generated directly — **not** converted from PDF. Real tables and text. | Editable | Unknown — Word paginates |
| `DOCX_VISUAL` | Renders PDF, rasterizes every page at 300 DPI, one full-page floating image per Word page. | Images | Exact |
| `XLSX` | Native Excel workbook. The default `REPORT` layout creates one native-cell worksheet per explicit RDL section; `DATA` preserves the legacy stacked/per-tablix export. | Live cells | Not paginated (`null`) |

Because Word performs `DOCX_EDITABLE` pagination itself, the service cannot know the count: the library
returns `pageCount: null` and the HTTP layer sends `X-Page-Count: unknown`. `XLSX` is also `pageCount: null`
(a spreadsheet is continuous; Excel decides print pagination). The other modes return numeric page counts.

Choose `DOCX_EDITABLE` when users need normal Word table editing and reflow (Word owns pagination, so page
breaks will not match the PDF). Choose `DOCX_VISUAL` when page-for-page fidelity matters and editing is not
required — it is one full-page image per page. Choose the default XLSX `REPORT` layout for an editable,
PDF-styled workbook without PDF pagination. It splits only at explicit RDL page breaks, keeps layout and
tablix content in native cells, preserves RDL group row spans as editable merged regions, and permits only
declared embedded RDL images as pictures. Set `"excel": { "layoutMode": "DATA" }` for the legacy data-first
stacked workbook; `sheetPerTablix` remains available only in that mode.

`DOCX_EDITABLE` splits large tablixes into multiple native Word tables by default using PDF-like page-break
estimates. Each fragment repeats declared header rows and carries a physical closing border while remaining
editable. Set `"docx": { "nativePageFragments": false }` on a render request (the legacy top-level
`docxNativePageFragments: false` is still accepted), or set `RDL_DOCX_NATIVE_PAGE_FRAGMENTS=false`, to
restore one continuous Word-owned table. Fragmentation is not a page-parity guarantee: some RDLs move closer
to the PDF while row-span-heavy reports can drift further. `/v1/analyze` returns
`structuredEditable.nativePageFragments.recommendation` so callers can decide whether to retain the default
or opt out for a certified report/data family.

For certified reports, you can also mount a structured DOCX profile file and let the service apply the
certified options only to matching RDL definitions:

```json
{
  "profiles": [
    {
      "id": "incident-dashboard-native",
      "certified": true,
      "match": {
        "definitionSha256": "sha256-from-v1-analyze"
      },
      "docx": {
        "nativePageFragments": true
      }
    }
  ]
}
```

Set `RDL_DOCX_PROFILE_PATH=/app/config/docx-profiles.json`. With
`RDL_DOCX_PROFILE_AUTO=false`, callers opt in using `"docx": { "profile": "incident-dashboard-native" }`.
With `RDL_DOCX_PROFILE_AUTO=true`, the first matching profile with `certified: true` is applied
automatically. Candidate profiles generated by `certify:docx` intentionally remain `certified: false`; they
match in `/v1/analyze` but do not auto-apply until a reviewer marks them certified after the exact SSRS
reference run. Explicit request selection (`"docx": { "profile": "..." }`) is still allowed for QA.
Request flags such as `"docx": { "nativePageFragments": false }` override the profile.
Profile files are validated before use. Each profile must have a unique header-safe `id`
(`A-Z`, `a-z`, `0-9`, `.`, `_`, `:`, `-`; max 128 characters), a `match` object containing
`definitionSha256`, `name`, or `namespace`, and only known `docx` options. Duplicate IDs, unknown rendering
keys, malformed SHA-256 hashes, and unsafe IDs fail closed with `CONFIG_INVALID`.

Fixed position does not mean read-only content. Fixed DOCX packages have no Word document/write protection,
every drawing anchor is emitted unlocked, and every visible report text line can be edited in its text box.
The service fails generation if package protection or a text-edit lock is detected.

## The request

| Field | Required | Notes |
| --- | --- | --- |
| `rdlBase64` | JSON only | The RDL. Multipart uses the `rdl` file part instead. |
| `output` | ✅ | `PDF` \| `DOCX_EDITABLE` \| `DOCX_VISUAL` \| `XLSX` |
| `datasets` | ✅ | Object of `datasetName` → array of row objects. |
| `parameters` | — | Validated against the RDL's declared types and defaults. |
| `subreports` | — | Render-time bundle of child `rdlBase64` definitions and invocation-scoped parameter/dataset instances. Supported for `PDF` and `DOCX_VISUAL`; see [Supplying subreports](./docs/SUBREPORTS.md). |
| `pagination.continuationMarkers` | — | `PDF` and `DOCX_EDITABLE`. When `true`, places “Continued from previous page” above the next table fragment for renderer-confirmed logical-row continuations. |
| `docx.nativePageFragments` | — | `DOCX_EDITABLE` only. Defaults to `true`; set `false` to restore one continuous Word-owned table. |
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
| `FONT_MISSING` | 503 | A required font is unavailable. See [Fonts](#fonts). |
| `BUSY` | 503 | At capacity. Includes `Retry-After`. |
| `RENDER_TIMEOUT` | 504 | Exceeded `RDL_RENDER_TIMEOUT_MS`. |
| `RENDER_FAILED` | 500 / 499 | Worker died, or client disconnected mid-render. |

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
| `RDL_WORKER_MEMORY_MAX_MB` | `2048` | Hard cap for deterministic PDF workload-based heap scaling. Never lower than the baseline. |
| `RDL_MAX_XML_NODES` | `250000` | XML expansion guard. |
| `RDL_MAX_XML_DEPTH` | `256` | XML nesting guard. |
| `RDL_DOCX_NATIVE_PAGE_FRAGMENTS` | `true` | Split large `DOCX_EDITABLE` tablixes into editable native table fragments at PDF-like break estimates. Set `false` for continuous Word-owned tables. |
| `RDL_DOCX_PROFILE_PATH` | unset | Optional JSON file containing certified structured-DOCX profiles matched by RDL hash/name/namespace. |
| `RDL_DOCX_PROFILE_AUTO` | `false` | Automatically apply the first matching structured-DOCX profile. Keep off unless every profile is release-certified. |
| `RDL_PDFTOPPM_PATH` | `pdftoppm` | Poppler binary for `DOCX_VISUAL`. |
| `RDL_BORDER_WIDTH_FLOOR_PT` | `0` | Minimum PDF border stroke, in points. `0` honours the RDL exactly. |
| `LOG_LEVEL` | `info` | Fastify log level. |
| `RDL_SAMPLES_DIR` | `tmp/samples` | Dev only. Where client samples live for tests and smoke scripts. See [Client samples](#client-samples). |

`RDL_BORDER_WIDTH_FLOOR_PT` exists for a reported "uneven borders" symptom that turned out to be viewer
sub-pixel rasterization, not stroke width. A floor only makes borders *heavier*, not more uniform, so it is
opt-in and off by default.

## Fonts

Font metrics drive layout. Wrong fonts means wrong pagination, so production is **strict by default**: a
missing font is `FONT_MISSING` (503) rather than a silent substitution that quietly shifts every page break.

- Licensed fonts are **never committed to the repository or baked into the image** (see `AGENTS.md`). Mount
  them at runtime into `RDL_FONT_DIR` (`/app/fonts` in the container).
- Reports commonly need Arial, Times New Roman, and — for the Combined Assurance profile — Segoe UI and
  Segoe UI Symbol (the legend glyphs). Segoe UI is licensed on the client machine.
- `Segoe UI Emoji` may use a mounted `NotoEmoji-Regular.ttf` as an explicitly enabled compatible PDF
  fallback. The renderer validates the actual glyphs before layout and embedding; supplementary-plane emoji
  never fall through to Helvetica. Enable this in strict mode only with
  `RDL_ALLOW_COMPATIBLE_FONT_FALLBACKS=true`.
- `GET /readyz` reports exactly which variants are missing.
- `RDL_STRICT_FONTS=false` substitutes them for local development. Never in production.

## Security model

The service treats every RDL as untrusted input, because an RDL is an executable document.

- **No data source execution.** `DataSource`, `ConnectString`, `Query`, `CommandText` are parsed so
  `/v1/analyze` can report them, then ignored. No SQL, no stored procedures, no connections. Rows come from
  the caller, who has already applied its own authorization.
- **No code execution.** Embedded `Code` is reported, never run. `CodeModules` and `Code.*` are rejected.
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
baseline heap, and a 2048 MB hard per-worker cap for exceptionally wide or text-heavy PDF workloads.

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
invocation-scoped tablix subreports render in PDF and visual DOCX; unresolved subreports and subreports in
editable DOCX/XLSX remain fail-closed. Maps, gauges, custom code, and non-embedded images are rejected.

To ask the service instead of reading docs, `POST /v1/analyze`. To dump the whole schema catalogue:

```bash
npm run audit:schema   # tmp/output/rdl-2016-capability-catalogue.json
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
report definitions, queries, and row values, so they live in **`tmp/samples/`** — outside version control —
not in the repository.

A fresh clone therefore legitimately has no samples, and that is not a broken checkout. Tests and scripts
that need them **skip** rather than fail:

```
✔ resolves CSS/X11 colour names to hex so PDFKit fills render
﹣ the incident dashboard with charts passes the capability gate (skipped)
```

To run the sample-backed tests, put the client files in `tmp/samples/`, or point at them elsewhere:

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
npm run certify:docx -- "/path/to/report.rdl" "/path/to/request.json" "/path/to/ssrs-reference.pdf" --renderer=word --exact-inputs --require-match
```

Inspect representative PDF/DOCX output. **An HTTP `200` is not verification.**

- **`verify:stress`** renders a deterministic 470-row table as direct PDF and structured editable DOCX,
  then converts the DOCX back to PDF. Its canonical PDF must be 30–40 pages; structured-DOCX page count is
  advisory because Word/LibreOffice own native-table pagination. The gate requires exact-once row and
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
  canonical PDF is constrained to 30–40 pages; editable Word pagination remains advisory. Oversized Word
  rows must be split into bounded native rows, and every emitted page fragment must carry a continuous
  full-grid closing rule. Artifacts are direct children of `tmp/`.
- **`verify:reference`** renders a hydrated PDF, compares page count and dimensions, checks text anchors,
  rasterizes the first three pages at 144 DPI, and writes a comparison report. Add `--exact-inputs
  --require-match` **only** when the hydration and font versions come from the exact SSRS reference run.
- **`certify:docx`** renders the canonical PDF plus both structured DOCX variants (`nativePageFragments`
  off/on), exports the DOCX through Microsoft Word for Mac or LibreOffice, compares page count, dimensions,
  extracted text, and OpenXML editability, then writes a `docx-profile-candidate.json` matched by
  `identity.definitionSha256`. Use `--renderer=word --exact-inputs --require-match` for release
  certification. Without the exact SSRS reference PDF/data/fonts, it produces a candidate only and reports
  the blocker explicitly.
  Without `--exact-inputs` it reports regression evidence but will never mark output certified. Page-count or
  pixel differences against a reference built from different rows are not certification failures by
  themselves.

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
    pdf.js          PDF renderer (PDFKit)
    docx.js         Native OpenXML renderer (editable)
    visualDocx.js   Rasterized full-page DOCX
    excel.js        Native XLSX renderer
    chart.js        Vector charts
    fonts.js        Font resolution + strict-font enforcement
  worker/
    runner.js       Concurrency admission, timeouts, temp lifecycle
    renderWorker.js Forked worker: parse → font check → validate → render
scripts/
  lib/samples.js    Resolves client samples (tmp/samples, or RDL_SAMPLES_DIR)
  …                 Smoke, stress, reference, and schema-audit tooling
test/               Node test-runner suite
tmp/                All generated output AND client samples. Git-ignored in full.
  samples/          Client RDLs, request fixtures, row data — never committed
  output/           Renders, verifier reports, rasterized pages
```

Contributors: read `AGENTS.md` first. It carries the authoritative architecture, security invariants,
supported RDL boundary, and verification workflow.
