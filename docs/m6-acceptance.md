# M6 integration evidence: GridCraft

Updated 2026-10-30. This records the GridCraft portion of tiny-vba M6. M6.1's
Word object profile belongs to WordCraft; this document checks GridCraft against
the shared-editor and real-workbook portions of M6. GridCraft remains on
`feat/m5-document-events-udfs`; the tiny-vba dependencies are pinned to
`98c7038a297ca9453c01d7e4a5530e8fc8865a10` (the merged M5 API).

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
