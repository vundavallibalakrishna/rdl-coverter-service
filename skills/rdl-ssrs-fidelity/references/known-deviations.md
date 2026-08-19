# Known Engine ↔ SSRS Deviations

Running log of places this engine is known to differ from SSRS, or where fidelity is unproven. Each entry
names the SSRS rule, the current behavior, and how to reconcile. Remove an entry only when it is fixed and
verified against a real SSRS oracle. Never let an entry justify a report-specific workaround.

## Open

- **No-format DateTime default.** SSRS renders a `DateTime` with no `Format` as general date/time (date and
  time). The engine currently emits a bare date (`dd/MM/yyyy`) because a unit test pins that default. Fix:
  change the no-format Date default to general date/time and update the test; verify against an SSRS
  reference. Ref: `value-formatting.md`.

- **Body coordinate model is partly flow-based.** SSRS positions body items absolutely by `Top`/`Left`;
  the top-level body band currently stacks items vertically (`gap = max(0, item.top − previousBottom)`),
  which is only correct for stacked layouts. Rectangle children and header/footer items already use true
  absolute positioning. Fix: absolute 2-D positioning for the body band; preserve output for purely stacked
  reports. Ref: `page-and-flow.md`.

- **Wide-matrix horizontal pagination.** SSRS splits a matrix wider than the printable body across
  horizontal page columns, repeating row headers. The engine draws the natural width and can overflow the
  right edge. Fix behind the advanced-materialization path. Ref: `page-and-flow.md`.

- **Body-scope `TotalPages`.** Correct in the header/footer pass; a body/tablix cell referencing
  `Globals!TotalPages` currently sees a provisional value. Fix: resolve final page count for body scope, or
  confirm no supported report needs it and document. Ref: `page-and-flow.md`.

- **Non-embedded images.** `Image.Source ∈ {External, Database}` currently renders nothing silently, which
  violates fail-closed. Fix: classify as `UNSUPPORTED_FEATURE` at analyze time (metadata-only policy still
  forbids fetching external resources). Ref: service security invariants.

- **Redundant coincident border strokes.** At shared tablix/group edges the engine can emit two coincident
  strokes (e.g. black + light-grey) and rely on draw-order to hide one. It renders as a single line today
  and is stable under row growth, but is fragile. Fix: resolve conflicting collinear coincident edges to a
  single precedence winner in the edge-flush layer. Ref: `border-resolution.md`.

## Verified fixed (kept for regression awareness)

- **Container-growth double border.** A Rectangle wrapping a Tablix drew its border at the declared height
  while the tablix grew, stranding a second bottom line. Fixed by painting the container border/box at the
  rendered extent (and correcting the trace height so DOCX inherits it). Ref: `sizing-and-growth.md`.

- **Footer date dump / crop.** `=Globals!ExecutionTime` was coerced with JS `String(Date)`, producing a
  multi-line timezone string that overflowed the fixed footer. Fixed by routing no-format Date values
  through the .NET format engine. (See the open deviation above for the remaining date-vs-datetime default.)
  Ref: `value-formatting.md`.

- **Chart float axis label + cropped category labels.** Tick values leaked `0.6000000000000001`; long
  category labels wrapped and were clipped by the plot clip rectangle. Fixed by formatting tick values to
  the interval precision and reserving the measured category-label band before sizing the plot.
  Ref: `charts-and-axes.md`.

## How to use this log

Before changing a renderer, check whether the construct appears here. When you fix an open item, move it to
"Verified fixed" with the SSRS rule and the verification evidence. When you discover a new deviation during
RCA, add it here with the rule it violates — even if you are not fixing it in the same change.
