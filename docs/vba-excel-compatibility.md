# Excel VBA source compatibility audit

Reviewed 2026-10-28 against public Microsoft Learn documentation and the actual
`gridcraft-vba` adapter, tiny-vba revision
`ea38c3f346b4010d6017b3a6d27e645ebad7ef79`. Target: existing Excel workbook source
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
| Application | Application/Parent, ThisWorkbook/ActiveWorkbook, ActiveSheet, indexed Worksheets/Sheets, Range and two-index Cells | Restricted to the editor's workbook; no Workbooks collection, Selection/ActiveCell, WorksheetFunction, Evaluate, Calculate, Caller, Run, Volatile or application settings/events |
| Workbook | Name, Application/Parent, ActiveSheet, indexed Worksheets/Sheets | Range/Cells/ThisWorkbook/ActiveWorkbook now correctly rejected on Workbook; Open/Save/Close, Names, paths and other workbook members absent |
| Worksheet | Name, Application/Parent, Range(Cell1, Cell2), Cells(row, column) | No CodeName roots, Activate/Select, UsedRange, Rows/Columns, Names, structural edits or formatting objects; topology changes invalidate the bridge |
| Range construction | Local rectangular A1 string; optional second string or same-sheet Range endpoint; bounding rectangle; case-insensitive named Cell1/Cell2 | No names, unions/intersections, qualified addresses or relative Range.Range; single-argument Range object is rejected; references over 65,536 cells are refused even before reading |
| Cells | Two positive integer arguments, one-based worksheet coordinates | Excel returns an all-cell Range with default Item; zero/one-argument forms, column letters, full Variant coercion and Excel Item argument names are absent. Current Row/Column named arguments are adapter extensions, not Excel Item signatures |
| Collections | Worksheets(index/name), Sheets(index/name) return a worksheet | These are projections, not collection objects. Count, Item, Add, enumeration and multi-index selection are absent. Sheets does not expose chart/macro sheets. Unqualified aliases are fixed-member imports, not Excel collection roots |
| Range ownership and geometry | Application, Parent/Worksheet, Row, Column, Count; read-only, one-based Long coordinates/count | Single rectangular area only; CountLarge, Rows/Columns, Areas, Item, Offset, Resize, End, Find and default-member dispatch absent |
| Range values | Explicit Value/Value2/Formula Get/Let, scalar or 1-based 2D Variant reads; matching-shape arrays or scalar broadcast writes | Value still follows Value2: no format-derived Date/Currency. No RangeValueDataType, Formula2/FormulaR1C1 or default Value access. Shape-mismatched arrays are refused rather than Excel's partial/error filling behavior |
| Value conversion | Empty, booleans, text, doubles and known cell CVErrs; typed native writes through engine commands | Literal Value strings are retained rather than fully matching Excel's input conversion. Date/Currency writes become numeric storage; integers beyond exact double precision, unknown errors and invalid UTF-16 are refused |
| ClearContents/Address | ClearContents through native undoable command; Address returns absolute local A1 | Address's five optional arguments, relative/R1C1/external references absent; other range mutation/formatting methods absent |
| Implicit roots | Application, ThisWorkbook, ActiveWorkbook, ActiveSheet, Worksheets, Range, Cells | Unqualified ActiveSheet is captured on editor opening; Application.ActiveSheet and global Range/Cells use the current sheet. Cross-workbook active/code-owner distinctions are not implemented |
| Language and references | Supported standard-module VBA subset, Object/Variant late binding, explicit properties | Worksheet/Excel.Range typed declarations fail compilation. Host type library/constants, document codenames, general default members, classes/UserForms and Office reference binding are not implemented |
| Project import | Inert read-only standard-module extraction; original VBA bytes preserved | Executable workbook/worksheet/class modules reject import; no Workbook_Open/Worksheet_Change event dispatch, controls/forms designer or P-code execution. Source edits do not rewrite vbaProject.bin or signatures |
| Worksheet UDFs | Public standard-module functions, value-only arguments/results, recalculation/spill and bounded execution | Reference arguments are value arrays, not Range objects; no Caller/Volatile, persistent UDF state or host access. Thus Range-typed/context-sensitive Excel UDFs do not work unchanged |
| Buttons/toolbars | Explicit public no-argument Sub assignment; inert DrawingML macro association and GridCraft toolbar metadata | Not full Excel Forms/ActiveX/control/event/CommandBars/Ribbon object models; no UserForm execution |
| Lifetime/resources | Explicit trust, document identity, undo groups, 8,192 handles, 65,536 cells and 8 MiB value payload caps | Cross-document and sheet topology changes stop execution; resource ceilings intentionally differ from Excel. Successful effects before error/stop remain undoable |

The catalogue describes this implemented projection for completion. It is not an
Excel type library. In particular its indexed Worksheets and Cells function
signatures must not be used as evidence of Excel collection/property parity.

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
on failure, handle validation and range budgets. Separate refusal tests pin typed
Excel declaration, collection, Cells and Address gaps so they cannot be mistaken
for successful workbook execution. Existing project/worksheet tests cover source
import refusals, inert package preservation, UDFs and their boundaries.

Validation on Windows with Rust 1.99.1: all 29 adapter tests pass natively and
under the actual Node-hosted `wasm-bindgen-test-runner`; all 1,150 optional
`gridcraft-ui-egui/vba-window` tests pass; strict all-target adapter/window UI
Clippy passes. Native UI tests use the pinned optional craft-fonts checkout and
serial test execution. Original source/fixture code only; no external workbook
corpus run is included in these counts.
The final `cargo +1.99.1 xtask ci` passes all ten gates (including native tests,
format/Clippy, assets/layers and browser compilation).

## Stages toward unchanged workbooks

1. Add generic host type/default-member/collection binding in tiny-vba, with
   generic tests only. GridCraft supplies Excel types/constants and real
   Worksheets/Sheets/Range Item collections. Preserve existing native ABI and
   named/omitted argument semantics. Accept typed Excel declarations and common
   collection/default-property source unchanged before widening the profile.
2. Separate lightweight range references from materialization budgets. Implement
   Cells/Rows/Columns/CountLarge/Item, relative Range, Offset/Resize, named ranges,
   UsedRange and End using engine state and native commands. Verify bounds,
   protection, undo/recalc and native/WASM behavior for each addition.
3. Implement correct active/code-owner roots and stable worksheet identities;
   Value versus Value2 coercion, Address signatures, formatting objects and
   application calculation/settings semantics. Never add success-returning stubs.
4. Integrate document modules/codenames and events with defined ordering,
   EnableEvents, cancellation and recursion budgets; then Range-aware UDFs and
   control events. Coordinate reusable language work in tiny-vba and Office
   behavior in GridCraft, retaining separate repositories.
5. Select a reviewed, licensed workbook corpus with declared dependencies and
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
