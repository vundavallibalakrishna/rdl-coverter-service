# Value Formatting Contract (SSRS)

SSRS formats every displayed value through the .NET format engine, governed by the textbox/run `Format`
property and the report `Language` (culture). Reproduce that engine — never a host-language (JavaScript)
coercion such as `String(value)` or `Date.prototype.toString()`.

## The rule

1. Resolve the value (expression → typed value: String, Integer/Decimal/Double, DateTime, Boolean, Null).
2. Resolve the effective culture from the report/textbox `Language` (default: invariant/host culture only
   if the report declares none — a real report almost always declares one, e.g. `en-ZA`, `en-GB`).
3. If a `Format` string is present, apply it via the .NET format engine (standard or custom pattern).
4. If **no** `Format` is present, render the value's default .NET `ToString(culture)`:
   - **DateTime** → general date/time **long** pattern "G": date **and** time, culture-ordered
     (e.g. `en-GB` → `18/08/2026 14:27:38`). Not `ddd MMM dd yyyy … GMT+…` (JS), not a bare date.
   - **Numeric** → general number "G": the significant digits with the culture decimal separator, no
     thousands grouping, no forced decimals.
   - **Boolean** → `True` / `False`.
   - **String** → the string unchanged.
   - **Null / Nothing** → empty string.

## Standard format strings to support (culture-aware)

- Numeric: `C[n]` currency, `N[n]` number w/ grouping, `F[n]` fixed, `P[n]` percent, `D[n]` integer,
  `E`/`e` scientific, `G` general, `X` hex.
- Date/time: `d` short date, `D` long date, `t` short time, `T` long time, `f`/`F` full, `g`/`G` general,
  `M`/`m` month-day, `Y`/`y` year-month, `s`/`o`/`R`/`u` round-trip/sortable.
- Custom numeric: `#,##0.00`, `0.##`, `0%`, section formats `pos;neg;zero`, literals and escapes.
- Custom date/time tokens resolve as whole tokens: `yyyy yy MMMM MMM MM M dddd ddd dd d HH H hh h mm m ss s
  tt t`. Never first-match-wins alternation (which truncates `MMMM`→`MMM`+`M`).

## Precision and correctness

- Do not leak binary floating-point artifacts into text. A computed tick/aggregate like `3 * 0.2` is
  `0.6000000000000001`; format to the value's own precision so it reads `0.6`. (See charts contract.)
- Percent multiplies by 100 and appends the culture percent sign.
- Currency uses the culture symbol/placement — do not hard-code `$`/`USD` for a non-US report.
- Dates format in a fixed, declared time zone so the same input is deterministic across hosts; SSRS renders
  the stored value, it does not localize to the render machine's zone.

## Engine placement

Formatting is resolved once, in the expression/format layer, and consumed identically by every renderer
(`styledTextForItem` / `textForItem` for PDF/DOCX text, the XLSX number-format mapping, DOCX runs). A Date
or number that reaches a renderer as a raw object must still pass through this engine before it becomes a
string — a renderer must never `String()` a typed value directly.

## Engine status

Reconciled: the no-format **DateTime** default renders general date/time (`formatValue` → `dd/MM/yyyy
HH:mm:ss`; the text path routes any Date-or-ISO-string with no format through the formatter; XLSX writes a
typed date). **Report culture** is honored — `formatValue`/`formatNet` take a culture, `globals.culture`
(from the report `Language`, canonicalized) reaches every renderer via `cultureFor(context)`, and `C`
currency uses the locale's currency (`currencyForCulture`); a report with no `Language` keeps the legacy
en-US-number / en-GB-date defaults. Explicit date formats are honored in XLSX too (translated to a live
Excel number format). Remaining gaps (see `known-deviations.md`): the *no-format* default date order is not
culture-reordered, and item-level `Language` overrides are not applied.
