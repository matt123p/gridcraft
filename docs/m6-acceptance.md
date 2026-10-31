# M6 integration evidence: GridCraft

Updated 2026-10-31. This records the GridCraft portion of tiny-vba M6. M6.1's
Word object profile belongs to WordCraft; this document checks GridCraft against
the shared-editor and real-workbook portions of M6. GridCraft is on
`feat/tiny-vba-integration`; the tiny-vba dependencies are pinned to
`da8e14938028a956749712676c742c92fb30c362` (the M6 shared API).

## Status

The selected GridCraft workflows pass against its actual engine session and
workbook objects. This is useful fork-side M6 evidence, but it does not certify
the complete M6 milestone: WordCraft lifecycle delivery and Microsoft Office
reference observations remain outstanding. The shared editor is used by both
forks; GridCraft does not carry a forked interpreter or substitute mock host.

## Executable evidence

The following checks passed during M6 work on 2026-10-30:

- `cargo check -p gridcraft-vba --offline`
- `cargo test -p gridcraft-vba`
- `cargo test -p gridcraft-ui-egui --features vba --test vba_events`
  (5 tests)
- `cargo test -p gridcraft-ui-egui --features vba --test vba_forms`
  (1 test)
- `cargo test -p gridcraft-ui-egui --features vba --test vba`
  (21 tests)
- `cargo test -p gridcraft-ui-egui --features vba-window --test vba_windows`
  (4 tests)

The adapter tests exercise real range and worksheet objects, typed values,
mutation and undo, recalculation, trust denial/revocation, pause and resume,
event suppression and FIFO delivery, reset/close cleanup, and modal form waits.
The UI tests exercise shared editor embedding/window behavior and source save,
import, and return routing. Source import is inert.

No Excel reference build was exercised for this evidence set. These are authored
GridCraft contract tests, not Excel compatibility observations. See
[`vba.md`](vba.md) for the bounded object profile and compatibility limits.

## Module-context conformance and completion update (2026-10-31)

GridCraft completed its M6.7 acceptance on the shared tiny-vba revision
`da8e14938028a956749712676c742c92fb30c362` (the M6 resolution/With/implicit-member
API, immutable Git pin, no path override). The adapter declares each document
module's implicit `Me` members from the supported catalogue (`Bridge::configure_document`,
keyed by live workbook code names: worksheet modules get the Worksheet catalogue members,
`ThisWorkbook` gets the Workbook members) and declares host-global result types and
document instance types for the shared IDE's read-only completion.

The specifications' unshadowed module-kind binding and live host-global consultation
cases now run as authored workflows against real workbook state in
`crates/vba/tests/context_resolution.rs` (four tests, native; wasm-bindgen targets
included):

- worksheet-module bare `Range` writes the module's own sheet while a different
  sheet is active; bare `ActiveSheet` in the same handler resolves live;
- `ThisWorkbook`-module bare `Range` follows the active sheet while
  `Me.Worksheets("Data")` and `Me.Name` stay on the project workbook;
- pausing a worksheet event, activating another sheet and moving the selection
  externally, then resuming, proves per-invocation consultation of live state and
  that bare `Range` in the worksheet module never follows the active sheet;
- declared host globals and call-shaped receivers (`ThisWorkbook.`, `Range("A1").`,
  `Worksheets("Data").`, `Sheet1.`) complete through the shared editor with zero
  host dispatch and no cell writes.

The adapter test suite passes 74 native tests (`cargo test -p gridcraft-vba`);
`vba_events` (5), `vba_forms` (1) and `vba` (21) UI tests pass with the `vba`
feature. `Application.Run`, add-ins and chart sheets remain refused/NOT_RUN in
tiny-vba's [M6.5 check](../../../tiny-vba/docs/m6-context-resolution-check.md);
Microsoft Excel reference observations remain NOT_RUN (no Office session on this
machine). Package round trips and signatures stay in M7.
