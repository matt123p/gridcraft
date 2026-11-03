# Excel VBA source compatibility audit

Reviewed 2026-11-03 against public Microsoft Learn documentation and the actual
`gridcraft-vba` adapter and the tiny-vba Git revision pinned in its manifest/lockfile.
Nominal type acceptance is recorded in [the M1.1 report](vba-m1-types.md).
Unchanged Excel workbook source is measured as compatibility evidence, not the
implementation target. `GridCraft.Excel.v1` is a bounded standard-module profile;
it must not be advertised as an identical Excel object model or general XLSM
execution support.

The accepted scope is narrower: VBA projects only behavior GridCraft already
supports. Binding and adapter work is in scope; adding a missing GridCraft
feature is not. Such members remain visible as `OUT_OF_SCOPE_GRIDCRAFT` and do
not count as unfinished VBA integration. See the
[accepted v1 priority report](vba-object-model-v1-priority-final.md).

For the documentation-led audit of the **entire Excel object-model surface**
and Luna's implementation packets, use the
[object-model implementation plan](vba-object-model-implementation-plan.md)
and [pinned 5,996-page inventory](vba-object-model-inventory.json).
That workstream covers methods, properties, events, collections and enums;
VBA built-ins and worksheet-function implementations are separate surfaces.

This audit checks every exposed dispatch branch, imports, catalogue, typed value
conversion, source import and worksheet-function path. It is not an exhaustive
enumeration of Excel's object library. Evidence is local source inspection and
independently authored native/JavaScript-WASM regression tests. No fresh Excel
oracle run or third-party workbook compatibility result is claimed. A passing
refusal test proves a clear boundary, not compatibility with the rejected source.

## Current surface and differences

| Area | Implemented behavior | Difference affecting unchanged source |
|---|---|---|
| Application | Code-owner ThisWorkbook; dynamic ActiveWorkbook/Sheet/Cell/Selection, collections/ranges; Calculation, Calculate, EnableEvents and narrow Goto | Workbooks and general WorksheetFunction/Evaluate remain absent. ScreenUpdating is unsupported; events are limited to the declared open/change profile |
| Workbook | Name, Application/Parent, ActiveSheet, Worksheets/Sheets collection objects and indexed access | Range/Cells/ThisWorkbook/ActiveWorkbook now correctly rejected on Workbook; Open/Save/Close, Names, paths and other workbook members absent |
| Worksheet | Stable identity, read/write Name through native rename, Application/Parent, Range/Cells/Rows/Columns/UsedRange, document CodeName bindings, narrow Activate/Select/Paste | Names collection, lifecycle methods and general sheet grouping remain absent; deleted sheets fail with 91 |
| Range construction | Lightweight A1/full-row/full-column references, bounding endpoints, relative Range.Range, Offset/Resize, quoted qualification, reference names, disjoint unions/intersections | No computed/external/3-D/structured references; overlapping areas refused; reference-only names use bounded local/workbook lookup |
| Cells | Whole-sheet/cell collection, zero/one/two indexes, column letters, Item names and VBA coercion | Dynamic unqualified root; concrete assigned objects retain ownership; reference construction does not materialize cells |
| Collections | Actual Worksheets/Sheets roots, Count, Item(index/name), default Item and bounded For Each; typed collection variables and Parent/Application | Add/mutation and multi-index selection remain absent. Both collections expose the worksheet kinds GridCraft actually supports; chart/macro sheets are not fabricated |
| Range ownership and geometry | Application, Parent/Worksheet, Long Row/Column/Count, Double CountLarge, Rows/Columns, Areas, Offset/Resize, sparse End and Find/repeat, EntireRow/EntireColumn | Rows/Columns use first area; multi-area scalar/array reads/writes and structural edits refuse explicitly; see M3 report for search and structure profile boundaries |
| Range values | Format-derived Date/Currency Value versus numeric Value2; scalar/1-based 2D reads; lower-bound-aware array cropping/error fill or scalar broadcast | No RangeValueDataType, Formula2/FormulaR1C1; fictitious Excel leap-day Date is refused |
| Value conversion | Typed native atomic values/formats, known CVErrs, General input/formula conversion and literal Text-format strings | en-US engine conversion profile; no universal locale compatibility; invalid UTF-16/unknown errors and inexact oversized integers refuse |
| Range clearing/Address | `Clear`, `ClearContents` and `ClearFormats` use bounded native clear commands; Address supports optional/named A1/R1C1, absolute/mixed/relative/external forms | Clearing is limited to 65,536 selected cells; saved filesystem paths/closed external references remain outside profile; M3 structural limits retained |
| Implicit roots | Application, ThisWorkbook, active workbook/sheet/cell/selection, collections, Range/Cells/Rows/Columns | Fresh getter resolution at source access; worksheet event calls scope unqualified Range construction to the code-name owner, and `Me` forwards through the captured live document owner. Unsupported catalogue members still refuse |
| Language and references | Selected Office host bindings, Object/Variant late binding, registered Excel nominal types, defaults and For Each | This is a bounded Excel adapter projection, not a complete typelib/COM loader. Shared-language classes/forms support does not establish the corresponding Excel object-model or binary-control support |
| Project import | Inert read-only standard/class/document-module extraction; original VBA bytes preserved | The declared M5 event profile is accepted: Workbook_Open is explicit, Worksheet_Change ordering/lifecycle and reviewed native edit sources are tested, and document `Me` dispatch binds to live workbook/sheet objects. Unsupported members still refuse. Source edits do not rewrite vbaProject.bin or signatures |
| Worksheet UDFs | Public standard-module functions, direct Range reference arguments, read-only Range values, Caller/Volatile, persistent project state, tested Date/Currency `Value` versus numeric `Value2`, and bounded execution | The declared M5 UDF profile is accepted against authored GridCraft workbook state; context-dependent document roots and cross-workbook/add-in registration remain unsupported |
| Buttons/toolbars/forms | Explicit public no-argument Sub assignment; inert DrawingML macro association; `.tvform` Label/TextBox/CommandButton runtime through the embedded IDE and workbook host | Not Office binary UserForm resources, ActiveX/control/CommandBars/Ribbon object models or workbook form persistence |
| Lifetime/resources | Per-target consent, document UIDs, stable sheet tokens, per-document undo groups and M3 resource bounds | External structural edits/undo invalidate saved coordinates with 1004; deleted sheets/closed documents refuse instead of retargeting |

The catalogue describes this implemented projection for completion. It is not an
Excel type library. The current collection/default profile and its executable
evidence are in [M1.2/M1.3 acceptance](vba-m1-defaults-and-collections.md).
M3 range/navigation/structural coverage is recorded in [M3 acceptance](vba-m3-ranges-and-worksheets.md).
M4 identity/value/member coverage and immutable dependency pin are recorded in [M4 acceptance](vba-m4-identity-values-and-members.md).

## Corrections made during this audit

2026-11-02: added explicit same-workbook Range destinations for
[`Application.Goto`](https://learn.microsoft.com/en-us/office/vba/api/excel.application.goto),
including inactive-sheet activation, multi-area selection, active-cell ownership
and named arguments. `crates/vba/tests/ranges.rs` covers the parenthesized Range
expression from the reported navigation pattern and negative argument paths.
Scroll=True, omitted-reference history, string/procedure destinations and
cross-workbook navigation remain outside this profile.

2026-11-02: the reported `Selection.Copy` failure is corrected alongside the
following worksheet/range Select and Worksheet.Paste calls. Range.Copy without
Destination uses the engine clipboard; `Range.Copy Destination` pastes a
single-area source directly to a single-area destination in the current workbook,
including another worksheet. Paste without Destination (Link omitted or False)
preserves cells, rebased formulas and formatting within one workbook.
Worksheet.Activate and Select with Replace omitted/True are supported.
Range.Select requires an active worksheet. The implementation delegates workbook
effects to `edit.copy`, `selection.set`, `sheet.activate` and `edit.paste`;
the adapter validates arguments, caps copied/pasted cells at 65,536 and queues
Worksheet_Change only after successful mutation. No system clipboard API is used.
`Worksheet.Paste Destination` also targets one Range in the current workbook,
activating its worksheet when needed. `Range.Cut` with or without Destination
stages an internal move; same-workbook paste clears the source only after a valid
target has been checked, and one Worksheet_Change receives both changed areas
when they are on the same sheet. `Application.CutCopyMode` reports
False/xlCopy/xlCut and cancels or stages a bounded selection. Cross-workbook
clipboard operations remain unsupported. Cut Destination moves a bounded
single-area source directly to a single-area destination in the current workbook.
`Range.PasteSpecial` supports the
native engine's all/formula/value/format/comment/validation/width modes and
arithmetic, SkipBlanks and Transpose options. Enum IDs and optional arguments
follow Microsoft's [Range.PasteSpecial](https://learn.microsoft.com/en-us/office/vba/api/excel.range.pastespecial),
[XlPasteType](https://learn.microsoft.com/en-us/office/vba/api/excel.xlpastetype)
and [XlPasteSpecialOperation](https://learn.microsoft.com/en-us/office/vba/api/excel.xlpastespecialoperation)
references. Source-theme and conditional-format merge paste modes remain refused.
`Range.Clear`, `ClearContents` and `ClearFormats` now use native clear commands
for selections up to 65,536 cells. Content-clearing paths preflight and queue a
Worksheet_Change target; format-only clearing does not queue that event. Runtime
acceptance for this addition has not yet been run.
Evidence: `crates/vba/tests/clipboard.rs` covers the unchanged recorded sequence
with two populated meal slots, both destination forms, Cut/Paste movement and
event targets, PasteSpecial enum/operation/transpose/skip-blanks paths,
formula/style fidelity, one macro undo, snapshot retention over pause/resume,
tiled paste callback targets, protection, empty clipboard and explicit refusals.
Link=True, sheet grouping,
multi-area/cut/foreign-workbook clipboard pastes remain outside
this profile. Semantics follow public [Range.Copy](https://learn.microsoft.com/en-us/office/vba/api/excel.range.copy),
[Range.Cut](https://learn.microsoft.com/en-us/office/vba/api/excel.range.cut),
[Application.CutCopyMode](https://learn.microsoft.com/en-us/office/vba/api/excel.application.cutcopymode),
[XlCutCopyMode](https://learn.microsoft.com/en-us/office/vba/api/excel.xlcutcopymode),
[Worksheet.Paste](https://learn.microsoft.com/en-us/office/vba/api/excel.worksheet.paste)
and [cell selection](https://learn.microsoft.com/en-us/office/vba/excel/concepts/cells-and-ranges/selecting-and-activating-cells)
documentation. No Food.xlsm execution or fresh Excel oracle result is claimed.

The host-configurable bounded execution profile now permits the fully populated
20-slot loop while retaining explicit resource failures. The focused clipboard
suite exercises the complete loop and one grouped undo.

Workbook no longer inherits Application-only Range, Cells, ThisWorkbook or
ActiveWorkbook. Application/Parent links retain object identity across supported
owners; Range.Worksheet returns the containing sheet. Worksheet.Range accepts
optional Cell2, named argument order, reversed corners and rectangular same-sheet
Range endpoints. Foreign/stale/wrong-type/cross-sheet endpoints fail explicitly.
Range.Row, Column and Count return Long values and reject writes/extra arguments.
Existing native mutation, protection, undo/recalculation and trust remain in use.

`crates/vba/tests/workbooks.rs` covers unchanged source for the corrected subset,
named endpoints, bounding rectangles, owner identity, exact error paths, no writes
on failure, handle validation and range budgets. New nominal tests execute typed
Excel declarations unchanged; refusal tests retain unregistered-type,
broader collection and Address gaps so they cannot be mistaken
for successful workbook execution. Existing project/worksheet tests cover source
import refusals, inert package preservation, UDFs and their boundaries.

Earlier exposed-profile audit validation on Windows with Rust 1.99.1: all 29 adapter tests pass natively and
under the actual Node-hosted `wasm-bindgen-test-runner`; all 1,150 optional
`gridcraft-ui-egui/vba-window` tests pass; strict all-target adapter/window UI
Clippy passes. Native UI tests use the pinned optional craft-fonts checkout and
serial test execution. Original source/fixture code only; no external workbook
corpus run is included in these counts.
The final `cargo +1.99.1 xtask ci` passes all ten gates (including native tests,
format/Clippy, assets/layers and browser compilation).

## Stages toward the bounded GridCraft projection

M1 registered types, defaults and selected collection workflows are implemented;
see the acceptance reports. M5 document modules, events and Range-aware UDFs are
accepted for the declared profile; see [the M5 evidence](../../tiny-vba/docs/m5-acceptance.md).
Remaining stages:

M3 range/worksheet and M4 identity/value/member implementations are described in
[M3 acceptance](vba-m3-ranges-and-worksheets.md) and
[M4 acceptance](vba-m4-identity-values-and-members.md). M4 uses the published immutable tiny-vba dependency; no local path override remains.

3. Expand application-specific support only for underlying behavior already
   implemented by GridCraft. Coordinate reusable language work in tiny-vba and
   binding work in GridCraft, retaining separate repositories. Missing GridCraft
   features stay on the ordinary product roadmap and are not dependencies of this
   VBA integration plan.
4. Select a reviewed, licensed workbook corpus with declared dependencies and
   independently established expected results. Report load/compile/run/save
   results separately, including source-byte identity, exact value/error tags,
   effects/events and preservation. Add binary project writeback only as its own
   reviewed implementation. Do not enable arbitrary imported macros on open.

## Public behavior references and provenance

Consulted 2026-10-28. These links support behavioral comparisons, not a claim of
Microsoft verification. Descriptions/tests/implementation are independently
authored; no tutorial source, Office binary, proprietary artwork, screenshots or
unlicensed fixture was copied or committed.

- [Application object](https://learn.microsoft.com/en-us/office/vba/api/excel.application%28object%29)
- [Workbook object](https://learn.microsoft.com/en-us/office/vba/api/excel.workbook)
- [Application.Range](https://learn.microsoft.com/en-us/office/vba/api/excel.application.range)
- [Workbook.Worksheets](https://learn.microsoft.com/en-us/office/vba/api/excel.workbook.worksheets)
- [Worksheet.Range](https://learn.microsoft.com/en-us/office/vba/api/excel.worksheet.range)
- [Worksheet.Cells](https://learn.microsoft.com/en-us/office/vba/api/excel.worksheet.cells)
- [Range.Item](https://learn.microsoft.com/en-us/office/vba/api/excel.range.item)
- [Range.Value](https://learn.microsoft.com/en-gb/office/vba/api/excel.range.value)
- [Range.Value2](https://learn.microsoft.com/en-us/office/vba/api/excel.range.value2)
- [Range.Address](https://learn.microsoft.com/en-us/office/vba/api/excel.range.address)
- [Range.Parent](https://learn.microsoft.com/en-us/office/vba/api/excel.range.parent)
- [Worksheet.Parent](https://learn.microsoft.com/en-us/office/vba/api/excel.worksheet.parent)
- [Range.Worksheet](https://learn.microsoft.com/en-us/office/vba/api/excel.range.worksheet)
- [Range.Row](https://learn.microsoft.com/en-us/office/vba/api/excel.range.row)
- [Range.Count](https://learn.microsoft.com/en-us/office/vba/api/excel.range.count)

2026-10-29: M3 supersedes the historical range/Cells omissions above. Public
documentation and original native/WASM fixtures remain the evidence; no Office
oracle or general workbook compatibility claim is added.
