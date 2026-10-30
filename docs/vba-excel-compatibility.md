# Excel VBA source compatibility audit

Reviewed 2026-10-29 against public Microsoft Learn documentation and the actual
`gridcraft-vba` adapter and the tiny-vba Git revision pinned in its manifest/lockfile.
Nominal type acceptance is recorded in [the M1.1 report](vba-m1-types.md).
Target: existing Excel workbook source
running unchanged. **That target is not met.** `GridCraft.Excel.v1` is a bounded
standard-module profile; it must not be advertised as an identical Excel object
model or general XLSM execution support.

This audit checks every exposed dispatch branch, imports, catalogue, typed value
conversion, source import and worksheet-function path. It is not an exhaustive
enumeration of Excel's object library. Evidence is local source inspection and
independently authored native/JavaScript-WASM regression tests. No fresh Excel
oracle run or third-party workbook compatibility result is claimed. A passing
refusal test proves a clear boundary, not compatibility with the rejected source.

## Current surface and differences

| Area | Implemented behavior | Difference affecting unchanged source |
|---|---|---|
| Application | Code-owner ThisWorkbook; dynamic ActiveWorkbook/Sheet/Cell/Selection, collections/ranges; Calculation and Calculate | No Workbooks, WorksheetFunction/Evaluate or event model. ScreenUpdating/EnableEvents are explicitly unsupported |
| Workbook | Name, Application/Parent, ActiveSheet, Worksheets/Sheets collection objects and indexed access | Range/Cells/ThisWorkbook/ActiveWorkbook now correctly rejected on Workbook; Open/Save/Close, Names, paths and other workbook members absent |
| Worksheet | Stable identity, read/write Name through native rename, Application/Parent, Range/Cells/Rows/Columns/UsedRange | No CodeName roots, Activate/Select or Names collection; deleted sheets fail with 91 |
| Range construction | Lightweight A1/full-row/full-column references, bounding endpoints, relative Range.Range, Offset/Resize, quoted qualification, reference names, disjoint unions/intersections | No computed/external/3-D/structured references; overlapping areas refused; reference-only names use bounded local/workbook lookup |
| Cells | Whole-sheet/cell collection, zero/one/two indexes, column letters, Item names and VBA coercion | Dynamic unqualified root; concrete assigned objects retain ownership; reference construction does not materialize cells |
| Collections | Actual Worksheets/Sheets roots, Count, Item(index/name), default Item and bounded For Each; typed collection variables and Parent/Application | Add/mutation and multi-index selection remain absent. Both collections expose the worksheet kinds GridCraft actually supports; chart/macro sheets are not fabricated |
| Range ownership and geometry | Application, Parent/Worksheet, Long Row/Column/Count, Double CountLarge, Rows/Columns, Areas, Offset/Resize, sparse End and Find/repeat, EntireRow/EntireColumn | Rows/Columns use first area; multi-area scalar/array reads/writes and structural edits refuse explicitly; see M3 report for search and structure profile boundaries |
| Range values | Format-derived Date/Currency Value versus numeric Value2; scalar/1-based 2D reads; lower-bound-aware array cropping/error fill or scalar broadcast | No RangeValueDataType, Formula2/FormulaR1C1; fictitious Excel leap-day Date is refused |
| Value conversion | Typed native atomic values/formats, known CVErrs, General input/formula conversion and literal Text-format strings | en-US engine conversion profile; no universal locale compatibility; invalid UTF-16/unknown errors and inexact oversized integers refuse |
| ClearContents/Address | Undoable native clearing/structure; Address optional/named A1/R1C1, absolute/mixed/relative/external forms | External owner title/sheet supported; saved filesystem path/closed external references remain outside profile; M3 structural limits retained |
| Implicit roots | Application, ThisWorkbook, active workbook/sheet/cell/selection, collections, Range/Cells/Rows/Columns | Fresh getter resolution at source access; worksheet event calls scope unqualified Range construction to the code-name owner, and `Me` forwards through the captured live document owner. Unsupported catalogue members still refuse |
| Language and references | Standard-module subset, Object/Variant late binding, explicit properties and registered Excel Application/Workbook/Worksheet/Range/Worksheets/Sheets nominal types, defaults and For Each | The supplied reference snapshot is a bounded adapter profile, not a complete Excel typelib/COM loader. Broader collections, classes/UserForms and unregistered Excel types remain unsupported |
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

## Stages toward unchanged workbooks

M1 registered types, defaults and selected collection workflows are implemented;
see the acceptance reports. M5 document modules, events and Range-aware UDFs are
accepted for the declared profile; see [the M5 evidence](../../tiny-vba/docs/m5-acceptance.md).
Remaining stages:

M3 range/worksheet and M4 identity/value/member implementations are described in
[M3 acceptance](vba-m3-ranges-and-worksheets.md) and
[M4 acceptance](vba-m4-identity-values-and-members.md). M4 uses the published immutable tiny-vba dependency; no local path override remains.

3. Expand application-specific support beyond the accepted M5 profile. Coordinate
   reusable language work in tiny-vba and Office behavior in GridCraft, retaining
   separate repositories.
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
