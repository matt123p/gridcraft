# M3 — Excel range and worksheet coverage

Updated 2026-10-29 on `feat/m3-ranges-and-worksheets`, based on completed M1
`55d3b1bb`. GridCraft owns all Excel definitions and workbook effects. The immutable tiny-vba dependency advances to completed M2
`3e16574eeb14395f00d5e7c47c412a4a501873ca`: the old M1 pin rejected `Is Nothing`
in normal Find workflows. All tiny-vba dependencies use that same revision;
unrelated locked packages remain unchanged. M3 adds no Office definitions or
application dependencies to tiny-vba. Workbook source storage still supports only
standard modules: classes/forms are explicitly refused before saving, and dirty
designs block implicit close. Shared form source/design export remains separate;
M6 workbook form persistence is not silently claimed by this dependency update.

The local ignored `plan/` and shared `craftrules` checkouts referenced by AGENTS.md
are absent. This work follows committed AGENTS.md, architecture and VBA contracts.
All source/fixtures are independently authored. Public Microsoft documentation is
behavior evidence, not a fresh Excel oracle observation.

## Stages

| Stage | Work | Status |
|---|---|---|
| M3.1 | Lightweight references, Cells zero/one/two indexes, column letters, Rows/Columns, Item, CountLarge | Implemented; final audit recorded below |
| M3.2 | Relative Range, Offset/Resize, quoted qualification, scoped names, unions/intersections and Areas | Implemented; final audit recorded below |
| M3.3 | Sparse UsedRange/End/Find/repeat and native structural edits, reference maintenance, protection/undo/recalc | Implemented; final audit recorded below |

## Contract and bounded profile

Reference descriptors retain sheet ownership and bounded geometry/collection
shape. Construction never allocates cells. Whole-sheet `Cells`, `Rows.Count`,
`Columns.Count`, `CountLarge`, last-row and sparse Find workflows work without
whole-sheet arrays. Count is Long and overflows with 6; CountLarge is Double.
Rows/Columns count and enumerate subranges, independently from Cells. Item uses
shared VBA Long binding (including numeric-string coercion and banker rounding),
relative one/two indexes and column letters. Unknown argument names fail with 448;
Null with 94, numeric overflow with 6 and out-of-sheet geometry with 1004.

Worksheet.Range accepts local A1/full-row/full-column strings, two endpoints,
quoted worksheet names (including escaped apostrophes), reference-valued local and
workbook names, colon, disjoint comma areas and intersections. Application.Range
can resolve another sheet. Qualified Worksheet/Range references and Range endpoint
objects must retain the receiver's sheet. Range.Range offsets literal addresses
from its first area's origin (dollar signs do not stop that offset); object endpoints
and defined-name targets retain absolute geometry. Offset defaults to zero;
Resize preserves omitted dimensions and rejects nonpositive or out-of-grid sizes.

Areas supports Count, Item/default Item and For Each. Cell iteration visits each
area in source order, row-major within it; Rows/Columns use only the first area.
Multi-area Value/Value2/Formula, ClearContents and structural edits refuse before
effects (1004). Overlapping areas, external/3-D/structured/computed names and
parenthesized expressions are explicitly unsupported; names are reference grammar,
not arbitrary formula evaluation. There are at most 64 disjoint areas, 4 KiB per
reference expression, 32 name/expression levels and 1,024 resolution steps.

UsedRange exposes the engine's current stored-cell/spill extent, or A1 on an empty
sheet. It does not emulate Excel's historical used extent after clearing or infer
extent from whole-row/column formatting. End scans stored cells on one lane;
formulas count as occupied even when their result is empty. It handles adjacent
regions, gaps and worksheet edges without stepping through empty cells. Merged,
hidden and filtered navigation is not a separately verified Excel profile.

Find supports nonempty scalar What, same-sheet single-cell After, xlValues or
xlFormulas, xlWhole/xlPart, row/column order, next/previous direction, MatchCase,
`*`/`?` wildcards and `~` escaping. It wraps once and returns Nothing on no match,
without moving the selection. LookIn/LookAt/SearchOrder persist across successful
calls; FindNext/FindPrevious retain the successful What/options. Invalid calls do
not replace them. Default LookIn is formulas, LookAt partial and order by rows.
Search uses General-format scalar text or formula source; format-dependent text,
comments, empty-cell searching, MatchByte=True and SearchFormat=True are refused
or outside this profile. The state is bridge-local, not synchronized with the UI
Find dialog. Case folding uses Rust Unicode lowercase, not locale-specific Excel
collation. End/Find cap each scan at 65,536 stored entries, and all scans share a
one-million-entry macro budget. Wildcard matching shares eight million steps per
macro, including failed calls, so an error handler cannot repeat expensive work
without consuming the limit.

EntireRow/EntireColumn and Insert/Delete route typed geometry into the same engine
bodies as ribbon/control structural commands. Whole rows/columns infer their axis;
partial-cell shifts require explicit xlShiftDown/xlShiftToRight or
xlShiftUp/xlShiftToLeft. CopyOrigin below/right and whole-sheet edits are refused;
format inheritance follows the existing engine profile. Protected partial shifts
are refused; line edits obey engine permissions. Boundary/table/push-off checks,
formula/name rewrites, recalculation, recording and undo stay in the engine. The
native hot path does not serialize/parse JSON; successful journal envelopes are
serialized for replay. The shared whole-line envelope routes also preserve the
explicit target sheet.

Saved range/area handles are preflighted and rebased through the engine's formula
reference adjustment before an edit; deleted references fail with 1004. A failure
leaves workbook and handle coordinates unchanged. Partial shifts follow the
engine's rectangle reference adjustment; split/discontiguous reference inference
is not fabricated. M4 supersedes the opening-sheet/topology restrictions with stable ownership and
dynamic roots; external structural edits/undo explicitly invalidate saved range
coordinates. See [M4 acceptance](vba-m4-identity-values-and-members.md).

Values retain the 65,536-cell/8 MiB budget. Enumeration preflights 8,192 shared
handles before producing loop entries. Structural line-format work is capped at
65,536 rows/columns per call. Geometry/index arithmetic is checked before effects.
Earlier successful writes remain in the macro's single undo group after a later
error or stop; failed operations add no partial worksheet effects.

## Acceptance

`crates/vba/tests/ranges.rs` executes original unchanged source for whole-sheet
counts, typed row/cell loops, relative ranges, quoted/local/global names,
intersections/areas, last-row, End, wildcard Find/repeat/wrap/no-match, formulas,
structural reference maintenance, nonactive sheets, protection, bounds and work
refusals. Existing adapter tests retain exact value/error tags, host ownership,
foreign/stale references, handle-budget atomicity, source identity and pause/trust
behavior. `crates/engine/tests/structure.rs` verifies native/envelope structural
parity, explicit-sheet journal replay, recalculation, undo and refusal atomicity.

Completion audit: the initial 48 native adapter tests and two native structural
engine tests pass. The integrated M3/M4 suite passes 58 native and 58 actual
JavaScript/WASM adapter tests; full GridCraft CI passes all ten gates (2,593
workspace test/doctest results). The final full CI rerun, 1,150 optional embedding tests, strict window Clippy
and browser feature compilation pass as recorded in the M4 report. This supersedes the old tiny-vba “M3 not started”
status. M3.1–M3.3 retain their original scope. Main branches are unchanged;
M5 events and Office package writeback are separate gates. M4 closes the dependency gate with published tiny-vba `22cc845e5ee74a1bcd15eec32cdb8a8f18d316f8`;
final acceptance uses the immutable Git dependency without a local override.

## Public behavior references

- [Item](https://learn.microsoft.com/office/vba/api/excel.range.item), [Cells](https://learn.microsoft.com/office/vba/api/excel.range.cells), [Rows](https://learn.microsoft.com/office/vba/api/excel.range.rows), [Columns](https://learn.microsoft.com/office/vba/api/excel.range.columns), [CountLarge](https://learn.microsoft.com/office/vba/api/excel.range.countlarge).
- [Range.Range](https://learn.microsoft.com/office/vba/api/excel.range.range), [Offset](https://learn.microsoft.com/office/vba/api/excel.range.offset), [Resize](https://learn.microsoft.com/office/vba/api/excel.range.resize).
- [UsedRange](https://learn.microsoft.com/office/vba/api/excel.worksheet.usedrange), [End](https://learn.microsoft.com/office/vba/api/excel.range.end), [Find](https://learn.microsoft.com/office/vba/api/excel.range.find), [Insert](https://learn.microsoft.com/office/vba/api/excel.range.insert), [Delete](https://learn.microsoft.com/office/vba/api/excel.range.delete).
