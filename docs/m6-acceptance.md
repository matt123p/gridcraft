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

## Gate environment note (2026-10-31)

`cargo xtask ci` passes every stage in this checkout except two pre-existing
`interaction` UI tests (`column_autocomplete_folds_case_beyond_ascii`,
`language_switch_renders_the_ribbon_and_persists`) that render glyphs (ẞ, ẻ) only
the optional `craft-fonts` checkout provides; `CRAFT_FONTS_DIR` is not available on
this machine and both tests fail identically on the pre-M6 commit, so they are
environment-blocked, not regressions. Verified here: i18n-report, fmt, strict
Clippy, the workspace test suite with those two tests skipped (all others pass,
including the 4 new `context_resolution` workflows and 847 UI interaction tests),
doctests, assets, layers, parity, keytips and the wasm gate.

## Attribute/code-view completion audit (2026-11-01)

The additional real-workbook `context_resolution` workflow imports an
attribute-only ThisWorkbook component, checks blank projected code and its code
name, then executes an independently authored Macro2 with the exact recorded
`VB_ProcData.VB_Invoke_Func` attribute. Import is inert, execution needs consent,
A1 becomes 42, raw source remains byte-identical and one workbook undo restores
the original cell. All five context-resolution workflows pass against the local
shared implementation. Full adapter/CI results are tracked in tiny-vba's `docs/m6-acceptance.md`.

Document code-name rename requests refuse explicitly: imported host identities
remain read-only until binary project writeback is designed. The shared raw
metadata inspector is read-only; Excel Macro Options/shortcut policy is not
implemented by displaying stored invocation metadata. Fresh Excel observations
remain NOT_RUN because no Office installation/session is available.

The final shared pin is published revision
`d40a283699ce0eb0e998493d3443731622050b3c` in manifests and Cargo.lock, with no
local Cargo override. All 75 adapter tests pass against that Git source. Full
`cargo xtask ci` passes all ten stages with the cached optional craft-fonts
checkout, resolving the earlier font-dependent acceptance failures. The optional
VBA Macros dialog also preserves its selected entry and shape-assignment context.

Final optional-feature checks: all 21 VBA UI workflows pass (including the two
macro/shape-context regressions); strict Clippy passes; the VBA-enabled web host
checks on wasm32-unknown-unknown. The WASM build reports a pre-existing
backstage Modified enum dead-code warning; it succeeds.
