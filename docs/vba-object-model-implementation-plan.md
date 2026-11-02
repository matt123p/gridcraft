# Excel VBA object-model audit and Luna implementation plan

Reviewed 2026-11-02. Status: audit snapshot and implementation handoff; new work
below is **NOT_STARTED**. Scope is the **third surface only: Excel's VBA object
model**, including its properties, methods, events, collections and enums.
This is a documentation-led audit, not an audit of a particular workbook.

## Scope and the evidence to use

Use Microsoft's [Excel object-model reference](https://learn.microsoft.com/en-us/office/vba/api/overview/excel/object-model)
as the target, not the list of operations already exposed by GridCraft. An engine
command, worksheet formula or visible UI feature does not establish a working
VBA object-model member. A catalogue entry does not prove every optional argument
or every Excel behavior works.

Do not audit or reimplement VBA intrinsic functions, syntax, or the worksheet
calculation library in this project. `Application.WorksheetFunction` is included
only as an object-model bridge to the existing formula engine; its mathematical
implementations and per-function parity remain a separate surface. Generic IDE
budget, continuation and host-service changes are allowed dependencies where
necessary to make this object model work, not a new language milestone.

The accompanying [machine-readable inventory](vba-object-model-inventory.json)
contains **all 5,996 `api/Excel.*.md` pages** in Microsoft's public documentation
repository at commit `b202c96d9d53896593b73d6fae57f5230601a5e3`. The Git tree was
not truncated; its full Excel page set was reconciled against the downloaded
archive. Only identifiers, classification, source paths, links and independently
written audit notes are included. No Microsoft sample implementation, artwork,
Office binary, workbook specimen or personal workbook source is included.

| Document heading kind | Pages |
|---|---:|
| Method | 1,351 |
| Property | 3,899 |
| Event | 123 |
| Object | 344 |
| Enumeration | 272 |
| Other reference topic | 7 |
| Total | 5,996 |

These are **page counts, not unique callable-feature counts**. The reference
includes legacy Excel Graph material, aliases and standalone member pages.
Some pages require reconciliation before becoming implementation requirements.
Member signatures, defaults, return types, read/write modes and platform remarks
must be expanded per packet; the snapshot is not a completed semantic oracle.

Implementation baseline: GridCraft HEAD
`c2c5ca8031a607089aae0206e1dbb69d57a74dcf` **plus the existing uncommitted
Goto/clipboard changes**, and its pinned tiny-vba revision
`0b41fd7bf7c78a62323133b290f2f9646667f9d3`. Rebase the inventory against the
actual checkout before coding. Do not discard, overwrite or reimplement those
uncommitted changes. The local ignored `plan/` files referenced by AGENTS.md are
absent; follow committed repository instructions and architecture.

## What the implementation actually exposes

The shipping adapter is principally `crates/vba/src/objects.rs`, with argument
binding in that file, type/enum registration in `types.rs`, completion metadata
in `lib.rs::catalog`, and specialized range/value/format/clipboard helpers.
`worksheet.rs` supplies a separate constrained UDF execution context. The
engine implements many more features than this adapter exposes.

| Surface | Current baseline | Missing work |
|---|---|---|
| Application | Dynamic workbook/sheet/cell/selection roots; calculation, EnableEvents; explicit same-workbook Range Goto | Workbooks/Windows access, Union/Intersect, general Run/Evaluate/WorksheetFunction, screen/alert/status settings, scheduling, UI and platform APIs; complete Goto variants |
| Workbooks/Workbook | Code-owner and active identity, selected properties and worksheet collections | Workbooks collection; Add/Open/Close/Save/SaveAs/SaveCopyAs, windows, names, metadata, links, protection and the remaining documented members |
| Worksheets/Sheets/Worksheet | Worksheet-only collections, Count/Item/enumeration, name, ranges, UsedRange, narrow Select/Activate/Paste | Add/Copy/Move/Delete and grouping; visibility/CodeName/Index; chartsheets; protection, data/print/drawing objects and remaining members |
| Range | Selected construction, default access, values/formulas, geometry, areas, offset/resize, End/Find, contents clearing, bounded structural operations and Select/Copy | Cut, Destination variants, PasteSpecial, fills, Replace, SpecialCells, CurrentRegion, merges, formatting breadth, formulas/local/R1C1/dynamic arrays, filtering/sorting and remaining methods/properties |
| Clipboard | Single-area Copy, same-workbook current-selection Worksheet.Paste, bounded tiling, formula/style preservation, Change callback | Destination forms, Cut, cross-workbook operations, PasteSpecial options, linked paste, multi-area rules, CutCopyMode, CopyPicture/system formats |
| Formatting | NumberFormat; selected Font properties; Interior.Color | Borders, alignment, themes, patterns, remaining font/fill members, Styles, conditional formatting, dimensions and page-facing formatting |
| Events | Explicit Workbook_Open and bounded Worksheet_Change delivery | Remaining Application/Workbook/Worksheet events, cancellation ByRef, ordering and host UI integration |
| UDF host | Range-aware arguments, restricted reads, Caller/Volatile context | General WorksheetFunction/Evaluate/Run are not implied by this path; preserve read-only UDF restrictions |
| All other object families | Engine/UI/package functionality may exist | No general VBA projection for Sort/SortFields, Names, ListObjects, Windows, PageSetup, charts/drawings, pivots, connections and most other documented objects |

Inventory status meanings:

| Status | Snapshot pages | Meaning |
|---|---:|---|
| `implemented-subset` | 87 | Positive catalogue/dispatch or event evidence; expand option/semantic gaps before claiming completion |
| `context-only` | 2 | Caller/Volatile on the UDF path, not general macro support |
| `object-subset` | 10 | An object-page alias for a type represented in the adapter |
| `not-exposed` | 4,618 | No matching member in the reviewed adapter projection |
| `object-not-exposed` | 334 | No general projection of this documented object |
| `formula-bridge-not-exposed` | 390 | WorksheetFunction member pages; bridge work only in this plan |
| `partial-definition` | 10 | Existing enum type registration; reconcile actual values/aliases |
| `definition-gap` | 262 | Enum definitions absent from the projection or awaiting reconciliation |
| `reference-triage` | 283 | Standalone/legacy/reference page; reconcile receiver, alias and platform before counting a feature |

No row is marked fully Excel-compatible. The groups overlap conceptually across
objects and members; do not turn these counts into a compatibility percentage.
The JSON packet field is an initial routing aid. The packet contract below takes
precedence where a member has multiple concerns, and OM20 must resolve the tail.

## Ownership and implementation rules for Luna

- GridCraft owns Excel object definitions, lifetimes, collections, enums,
  workbook/document state, UI, undo/recalc, file persistence and services.
  Reuse native engine commands; add engine primitives when no correct one
  exists. Do not implement a second workbook or duplicate an engine algorithm
  inside the VBA adapter.
- tiny-vba owns generic runtime/debugger/IDE/host protocol changes. Use its
  repository-local development and focused skills. Do not add Excel classes to
  tiny-vba. Keep dependency revisions immutable and synchronized across GridCraft;
  publish/test a shared change before updating the fork pin. Coordinate the
  equivalent WordCraft dependency only when a shared contract changes.
- Each member needs exact argument names/order, Missing/Empty/Null/Nothing,
  optional defaults, return tag, Get/Let/Set/Call/enumeration behavior, errors,
  collection indexes/default members, receiver ownership and effect metadata.
  Use the existing typed foreign protocol. Never coerce host values through JSON;
  JSON command envelopes may identify an operation without serializing VBA values.
- A supported receiver must return the correct live object, including captured
  identity across active-sheet/workbook changes. Respect generations, host
  identity, deleted/closed objects, rename/reorder/structural edits, and read-only
  inspection. Use one shared registry, not parallel identity stores.
- Validate the entire request, budgets, ownership, trust, protection and event
  capacity before mutation. Preserve partial prior macro writes and the existing
  undo grouping policy on runtime error/Reset. Never silently truncate a copy,
  paste, sort, enumeration or allocation to make it fit.
- Keep long work bounded and cancellable. Native/browser/headless consumers
  need explicit service availability; no ambient filesystem/network/shell/COM
  access inside the VM and no generic desktop automation in the adapter.
- Unsupported, unavailable, deprecated and platform-restricted are distinct
  outcomes. Record every documented row even where implementation needs a new
  service or policy contract. Do not hide difficult APIs by omitting them from
  the denominator or declaring all of an object supported after one method works.
- Preserve current roadmap IDs and completed intrinsic work. OM IDs below are
  a new cross-repository object-model workstream, not replacements for M1-M8.

## Work packets and order

All packets below are NOT_STARTED except the explicitly described existing
baseline. Split each packet into small PRs by object family and behavior. Do not
make one PR containing the whole object model. First finish OM00/OM01, then
OM02-OM07; continue through the remaining families with their prerequisites.
OM20 classification starts early and must be complete before final sign-off.

| Packet | Initial pages | Deliverable | Prerequisites |
|---|---:|---|---|
| OM00 | 272 | Complete inventory, signatures, enum and constant contracts | None |
| OM01 | Cross-cutting | Host execution policy, member registry and diagnostics | OM00 schema |
| OM02 | 179 | Selection, clipboard, editing and range operations | OM01 |
| OM03 | 200 | Application roots, settings, calculation and dispatch | OM01 |
| OM04 | 268 | Workbook/sheet collections and lifecycle | OM01, OM03 |
| OM05 | 528 | Formatting, styles and conditional formatting | OM02, OM04 |
| OM06 | 172 | Sort/filter/validation/outline/data tools | OM02, OM05 |
| OM07 | 497 | Names, formulas and evaluation/function bridges | OM03, OM04 |
| OM08 | 162 | Windows, panes, view and navigation | OM03, OM04 |
| OM09 | 94 | Protection, permissions and protected document state | OM04 |
| OM10 | 132 | Page setup, printing and export | OM04, OM05, OM08 |
| OM11 | 123 | Application/workbook/worksheet event completeness | OM01, OM04; incremental with every mutating packet |
| OM12 | 1,219 | Charts, drawings, images and associated objects | OM04, OM05, OM08 |
| OM13 | 832 | Pivot objects, slicers, timelines, cube/model APIs | OM04, OM06, OM14 where externally backed |
| OM14 | 295 | Connections, query/import/XML and refresh | OM04, explicit services |
| OM15 | 105 | Dialogs, interaction, proofing and built-in controls | OM03, OM08, OM11 |
| OM16 | 135 | Sharing, collaboration, publishing and remote document APIs | OM04, OM09, explicit services |
| OM17 | 558 | Legacy/Graph/OS/COM/XLL/VBE interoperability decisions and adapters | OM00, explicit platform/service contracts |
| OM18 | 131 | Structured tables and table styles | OM04, OM05, OM06, OM07 |
| OM19 | 94 | Notes/comments, hyperlinks and document metadata | OM04, OM05, OM11 |
| OM20 | Cross-cutting | Reconcile aliases/reference pages and finish acceptance; zero unrouted pages | Starts at OM00; closes after all packets |

Page routing is not a project estimate: for example, one chart method may need
several engine capabilities, whereas many enum pages can be registered together.
Record reassignments in the inventory, keeping the original pinned source set.

### OM00 — inventory and contracts first

Implement a Rust maintenance exporter/checker, preferably `cargo xtask vba-audit`,
that reads a locally cached pinned documentation snapshot without fetching at
build/test time. Reproduce the current page set and provenance. Preserve every
input page; recognize aliases, legacy Excel Graph topics and enum pages.
Maintain stable receiver-qualified IDs and retain source page IDs separately
when several pages map to the same member.

Expand rows with member kind, signature, argument modes/defaults, return type,
read/write support, event parameters/cancellation, enum values/aliases, platform
notes, dependencies, profile, implementation symbol, test IDs and result state.
Use states such as MISSING, PARTIAL, IMPLEMENTED_TESTED, PLATFORM_SERVICE_NEEDED,
LEGACY_REVIEW and REFERENCE_ALIAS. A service-needed row remains outstanding;
it is not an automatic completion exception. Fully implemented states require
parameter-level acceptance. Catalogue/source comparison is evidence, not a
substitute for tests. Check declared metadata against actual dispatch branches.

Register constants in GridCraft's `types.rs` with typed Long values, correct
enum identities, qualified/unqualified lookup and aliases. Prioritize clipboard,
sort/filter, alignment, sheets, print and protection enums with their consumers.
Missing `xl...` identifiers in code without Option Explicit can be mistaken for
implicit variables; test the constants' actual values, not merely compilation.

Acceptance: no missing/duplicate source paths, no unassigned canonical members,
no silently discarded aliases, stable deterministic output, baseline dispatch
reconciled, and a report that lists every unresolved row without a misleading
overall compatibility percentage.

### OM01 — cross-cutting execution and diagnostics

The shared IDE hardcodes `Limits::default()` at normal launch and idle document
callback creation (`tiny-vba/crates/vba-ide/src/lib.rs`); the VM default currently
permits 1,024 host calls. Add an explicit host-configurable bounded execution
policy through the shared IDE, then choose and test GridCraft's macro budget.
Do not just remove the VM limit or increase every consumer's global defaults.
Validate configuration, freeze it for a running/paused/waiting session, use it
on every launch/callback route, and keep inspection/UDF budgets distinct.
Count work over slices/continuations; a pause must not replenish the budget.
Reset/new invocation semantics must be documented and tested.

Use a unified member descriptor source for catalogue, dispatch validation,
completion, object browser and the audit inventory. Include supported argument
combinations and contextual availability. Diagnostics must name the receiver,
member, offending option and source location; distinguish unknown member 438
from unsupported option, unavailable service and exhausted resource budget.
Add a nonexecuting compatibility report for statically resolved receivers.
Dynamic receivers must be reported as UNKNOWN, never silently certified.

Acceptance: low-budget deterministic refusal, larger bounded macro completion,
responsive cancellation, pause/resume without duplicate writes, stale source/
handles rejected, and missing/partial/context-only members visible before or
during execution. Existing trust gates must still guard pure code and host calls.

### OM02 — finish recorded editing and clipboard workflows

Work in `clipboard.rs`, `objects.rs`, `ranges.rs`, `structure.rs`, `references.rs`
and the relevant engine command modules. Implement these in separate slices:

1. Full `Range.Copy`/`Range.Cut` Destination and clipboard forms;
   `Worksheet.Paste` Destination/Link; `Range.PasteSpecial` Paste/Operation/
   SkipBlanks/Transpose; `Worksheet.PasteSpecial` documented format variants;
   `Application.CutCopyMode`, clipboard cancellation and CopyPicture.
2. ActiveCell/Selection/Activate/Select, grouped worksheet selection, complete
   Goto/history/scroll variants, Union/Intersect (shared with OM03/OM07), and
   multi-area rules. Do not substitute Goto for Select and change activation.
3. Clear/ClearFormats and other documented clearing members; Replace; complete
   Find/repeat state; CurrentRegion, CurrentArray, SpecialCells; fill/AutoFill/
   series; insert/delete/entire-row-column completeness and range grouping.

Reuse `edit.copy`, `edit.cut`, `edit.paste`, `edit.pasteSpecial`, `selection.set`
and structural primitives where their semantics match. Internal workbook
clipboard and system clipboard formats are different host capabilities. A cut
must not erase its source until a valid move/paste commits; protection, grid
overflow, overlap and cancellation must not lose source data. Validate source
and destination consent separately across workbooks. Preserve formulas, styles,
merges, notes, validation, images and links according to each paste mode rather
than promising all modes copy all metadata.

Acceptance: ascending-size/large bounded ranges, overlapping moves, cross-sheet/
workbook ownership, numeric/text/error/formula/style fidelity, relative formula
rebasing, transpose/skip blanks/arithmetic, source preservation on failed cut,
correct selection, one macro undo group, callback footprint/order and native/
browser service-unavailable results. Check each optional argument's combination,
not just the no-argument path. See Microsoft [Cut](https://learn.microsoft.com/en-us/office/vba/api/excel.range.cut),
[Copy](https://learn.microsoft.com/en-us/office/vba/api/excel.range.copy),
[Paste](https://learn.microsoft.com/en-us/office/vba/api/excel.worksheet.paste)
and [PasteSpecial](https://learn.microsoft.com/en-us/office/vba/api/excel.range.pastespecial).

### OM03 — Application API

Add missing global roots/collections and preserve unqualified root behavior.
Implement ScreenUpdating, DisplayAlerts, StatusBar/DisplayStatusBar, Cursor,
calculation state/full/rebuild modes and other settings through real host state.
ScreenUpdating=False must suppress workbook painting where appropriate while
keeping the editor, cancellation and logic responsive; it is not a no-op flag.
DisplayAlerts=False needs documented default responses, not blind approval of
destructive operations. Coordinate scope/restoration with normal completion,
error, Reset and host close according to documented behavior and explicit host
termination policy.

Add bounded Application.Run, Undo/OnUndo/OnRepeat, OnKey/OnTime/Wait, path/
environment/version/locale queries, custom lists and documented utility methods.
Schedule through injected host time/queues; avoid blocking the UI thread and
cancel retained callbacks on host/project teardown. Treat Evaluate and formula
bridges as OM07. Route remote/OS/legacy APIs to OM16/OM17 rather than faking them.
Acceptance must exercise separate documents, live state, callback ownership,
teardown and headless/browser availability. [Application reference](https://learn.microsoft.com/en-us/office/vba/api/excel.application(object)).

### OM04 — workbook and worksheet lifecycle

Implement Workbooks and the full documented collection/default/enumeration
contracts, then workbook Add/Open/Activate/Close/Save/SaveAs/SaveCopyAs and sheet
Add/Copy/Move/Delete/visibility/index/CodeName/chart-sheet support. Include
optional Before/After, omitted destination/new workbook, count/template, naming
collisions, workbook structure protection and stale references. Model sheet
grouping explicitly so editing a grouped selection has the correct scope.

Use existing document/window/file commands and services, not raw VM IO. Suppress
macro execution during import; event delivery remains explicit and ordered.
Preserve package parts and source identities on persistence. Original binary
VBA preservation is not permission to claim edited binary project writeback;
coordinate that work with the existing package milestone and signing contracts.
Test cancelled/failed open/save/close, unsaved prompts, read-only/password paths,
cross-workbook moves/copies, event cancellation, active/code-owner separation,
all involved undo groups, file round trips and closed/deleted handles.
[Workbooks](https://learn.microsoft.com/en-us/office/vba/api/excel.workbooks),
[Workbook](https://learn.microsoft.com/en-us/office/vba/api/excel.workbook),
[Worksheet](https://learn.microsoft.com/en-us/office/vba/api/excel.worksheet).

### OM05 — formatting and conditional formatting

Complete Font, Interior, Borders/Border, Styles/Style and the range alignment,
wrap/shrink/indent/orientation, row/column sizing/hidden, merge and number-format
families. Implement mixed-range return values, theme/indexed/RGB color identities
and conversions, clear/copy behavior, localized format variants and protection.
Then project FormatConditions and each documented condition-specific object:
rules, scales, bars, icons, priorities, StopIfTrue, AppliesTo and modification.
Use style interners and existing conditional-format models/commands.
Test mixed reads, scalar/bulk writes, theme round trips, reference rebasing,
format-only pastes, merged geometry, large-range budgets and undo.
No separate styling state should diverge from what the workbook renders/saves.

### OM06 — sorting, filtering, validation and data manipulation

Add live Sort/SortFields/SortField state, SetRange, keys, Header/MatchCase/
Orientation/SortMethod and Apply. Register all relevant enum constants; support
named/omitted arguments and defaults exactly. Clear/Add changes configuration;
only Apply performs the sort. Repeated Worksheet.Sort getters must reference the
same logical sheet sort configuration, not fresh empty objects. Sorting an
inactive but explicitly owned range must not silently sort the active sheet.

Map valid state to `data.sort`; explicitly provide nonempty keys. The current
engine opens a sort dialog when keys are empty, which must not become a
programmatic VBA success path. Bound keys/footprints and preflight before
mutation. Its current Unicode-lowercase/codepoint text comparison is not proof
of locale/Pinyin/Stroke order. Implement/inject and test required collation or
declare the exact unsupported profile; do not accept SortMethod and ignore it.

Then implement Range/Worksheet AutoFilter, Filters/Filter and criteria operators,
ShowAllData, AdvancedFilter, validation objects, RemoveDuplicates, TextToColumns,
Subtotal, outline/grouping, scenarios/consolidation/goal-seeking and the remaining
data-tool members routed to this packet. Reuse native sort/filter/validation/
outline commands. Test blanks/errors/mixed types/headers/stability, descending
order, multiple keys, metadata/formula movement, protected/merged cells, saved
state, event footprints, cancellation and zero writes on invalid configuration.
[Sort](https://learn.microsoft.com/en-us/office/vba/api/excel.sort),
[SortFields.Add](https://learn.microsoft.com/en-us/office/vba/api/excel.sortfields.add).

### OM07 — names and formula-facing object-model APIs

Implement workbook/sheet Names/Name: scope, Add/Delete, RefersTo variants,
RefersToRange, visibility and identity. Complete Formula/FormulaR1C1/Formula2,
array/spill and local forms without confusing displayed text and formula source.
Add Range/Worksheet/Application Evaluate and Application.WorksheetFunction
through a typed context-sensitive bridge to existing engine functions. Preserve
Range arguments, arrays, errors and Excel return/error differences between
Evaluate, WorksheetFunction calls and worksheet formulas. Unsupported calculator
functions remain explicit; this packet does not implement their mathematics.
Inventory the 390 WorksheetFunction member pages as bridge requirements only.
Test reference ownership, dynamic arrays/spills, default members, errors versus
raised runtime errors, UDF read-only policy, locale and recursion/payload budgets.

### OM08 — windows, panes and navigation

Project Windows/Window/Panes/Pane and custom/sheet views onto real per-window
state: ActiveWindow, activation, selection, scrolling, freeze/split panes, Zoom,
gridlines/headings, display modes, bounds and window lifecycle. Worksheet-owned
state and window-owned state must remain distinct. Goto scrolling depends on
this packet. Test two windows on one workbook, two workbooks, focus changes,
saved views, closed handles, narrow/native/browser layouts and headless results.
[Window reference](https://learn.microsoft.com/en-us/office/vba/api/excel.window).

### OM09 — protection and permissions

Project workbook/sheet/range protection and AllowEditRanges, protection status,
password/encryption options, protected-view and document permissions through
existing protection models plus explicit services where needed. Preserve option
semantics such as UI-only protection and allowed edits; do not treat a password
string as proof that encryption exists. Validate protection for every newly
exposed mutation. Test locked/unlocked cells, structure protection, failed
authentication, save/reopen, cancellation and retained trust boundaries.

### OM10 — printing and export

Implement PageSetup, print areas/titles, margins, paper/orientation/scaling,
headers/footers, page breaks and Workbook/Worksheet/Range PrintOut/PrintPreview/
ExportAsFixedFormat. Delegate to existing print/PDF services; implement missing
engine options rather than acknowledging ignored arguments. Printer selection,
preview and native printing are service-backed; pure file/PDF export can have a
headless path. Test page geometry and selected-sheet/range scope by rendering
output and inspecting it, plus cancellation, unavailable printer and file errors.
[PageSetup](https://learn.microsoft.com/en-us/office/vba/api/excel.pagesetup).

### OM11 — event completeness

Inventory every documented event and ByRef/Cancel parameter. Extend explicit
open/change delivery to activate/deactivate/selection/calculate/before-save/
before-close/before-print/new-sheet/input events and the dependent chart/pivot/
connection/control events. Host commands, UI edits and VBA writes need one
defined ordering and suppression model. EnableEvents must apply at the correct
scope; UDF execution must not accidentally enable mutating events.
Use cancellable precommit events where documented; do not deliver BeforeSave or
BeforeClose after irreversible effects. Test reentrancy/recursion limits, FIFO,
ByRef cancellation, cross-document callbacks, error policy, modal waits, pause/
resume, source edits, Reset/close and event subscriptions across saved projects.
Do not retain document borrows or mutable locks across callbacks.

### OM12 — charts and drawing objects

Implement ChartObjects/ChartObject, Charts/Chart, ChartArea/PlotArea, Series/
SeriesCollection, axes/titles/legends/labels, data binding and chart settings;
then Shapes/Shape/ShapeRange, pictures, text frames, connectors, groups, geometry,
z-order, formatting, placement, OnAction, export and copy/paste. Include all
associated child/collection objects from the inventory, not just creation APIs.
Reuse chart/drawing models, commands and renderer. Engine feature gaps are
separate dependent tasks within this packet. Test object identity, data edits,
anchor transforms, grouping/ungrouping, protected sheets, undo, source deletion,
OOXML round trips and rendered native/browser/export output.
[Shapes](https://learn.microsoft.com/en-us/office/vba/api/excel.shapes).

### OM13 — pivots, slicers and data models

Project PivotCaches/PivotCache, PivotTables/PivotTable, fields/items/layout,
filters/calculated fields, refresh/drilldown, charts, slicers/timelines and the
documented cube/model/relationship/measure objects. Use existing pivot models;
do not create superficial objects that cannot calculate or persist. Separate
local pivot operations from external OLAP/data-model services. Test cache reuse,
refresh, missing data, field identity, filter/layout results, saved files, events,
undo and service cancellation. Unsupported provider functionality remains a
visible outstanding requirement, not an empty successful refresh.

### OM14 — connections, import and refresh

Implement Connections/WorkbookConnection, QueryTables/QueryTable, queries,
parameters, text/ODBC/OLEDB/XML/data-feed connection objects, mappings and refresh
lifecycle. Native parser/CSV/OOXML facilities can back local operations; external
drivers/network/authentication require injected services and explicit host
availability. Capture credential ownership without saving secrets into generic
metadata/logs. Test disconnected/cancelled/failed/async refresh, typed data,
dependencies, callback order, partial-result policy and package preservation.

### OM15 — interaction and built-in controls

Implement documented Excel dialogs, Application.InputBox type/Cancel semantics,
file pickers, status/interaction surfaces, dialog-sheet/form-control objects and
their action/event wiring. Scope built-in Office controls according to existing
tiny-vba form contracts; third-party ActiveX/OLE controls remain outside the
adopted policy and must be listed explicitly. No fake dialogs or success without
user-visible effects. UI interaction needs suspend/resume with released borrows,
bounded callbacks and explicit headless/browser behavior. Test modal ownership,
Cancel and typed results, selection/focus, project teardown and accessibility.

### OM16 — remote, collaboration and publishing APIs

Keep every sharing/review/mail/server/SharePoint/publishing/cloud-related row
visible. Identify which GridCraft feature/service can reproduce the documented
contract and which needs a new provider or a platform-specific profile. There
are currently no general host adapters for these APIs. Split implementation by
provider and data model, including authentication, conflicts, async completion,
failure/cancellation and package metadata. Do not reinterpret SendMail as a
no-op, automatically send messages, or silently replace a Microsoft service with
a different one. Implementation activation requires an explicit host feature;
planning does not authorize external sends or remote state changes.

### OM17 — legacy and platform interoperability

Reconcile Excel Graph legacy topics, XLM/Excel4 methods, DDE, COM/XLL/add-ins,
OLE, OS-specific operations, VBProject/VBE and similar APIs. Classify aliases,
deprecated-but-documented contracts and true platform dependencies. Preserve
all rows even when a native portable implementation is inappropriate.
The existing Rust-only/no-ambient-COM/no-VM-IO and third-party-control boundaries
still apply. Implement portable equivalents only when their behavior is explicit;
otherwise retain PLATFORM_SERVICE_NEEDED or POLICY_BLOCKED with exact reasons.
Any policy expansion needs its own reviewed design, not an implicit exception in
an implementation PR. A complete portable profile and complete Windows Excel
interoperability are different acceptance claims; publish both denominators.

### OM18 — structured tables

Implement ListObjects/ListObject, ListRows/ListRow, ListColumns/ListColumn,
table ranges/totals/header rows, resize/convert/add/delete, styles, filters and
structured references through actual table models. Include TableObject and
connection-backed variants with OM14. Test empty-table ranges, headers and
totals, calculated columns, row insertion/deletion, reference maintenance,
filter/sort interaction, protection, formula spill interaction and saved files.
[ListObject](https://learn.microsoft.com/en-us/office/vba/api/excel.listobject).

### OM19 — annotations and metadata

Implement Comments/Comment, threaded comments/persons/replies where documented,
Hyperlinks/Hyperlink, custom and built-in document properties, custom worksheet
properties and annotation formatting/visibility. Distinguish local notes from
collaborative threaded state and external follow-link actions. Use current
metadata/comment/link commands and package models. Test copy/move/delete,
structural rebase, author/thread identities, formatting, undo, file round trips
and unavailable remote/person services.

### OM20 — finish the tail and acceptance

Every current page has a family packet; there are zero unrouted OM20 pages.
Reconcile every `reference-triage` row within its assigned family, including
names with unusual receiver spelling and legacy standalone pages. Ensure smart
tag actions/proofing (OM15), phonetics/characters (OM05), add-ins (OM17), cell
watches (OM08), file converters/recovery (OM04) and the smaller object families
are implemented rather than lost behind the headline objects. New source pages
must be routed to a precise packet or a new explicit family packet. Do not
introduce an unexamined miscellaneous bucket at sign-off. Resolve duplicates/
aliases without deleting source-page provenance.

Run a final object graph reachability audit: a member cannot be complete when
its receiver cannot be obtained, its return type has no usable projection, or
its enum/default/child collection is missing. Check parameter combinations,
events, lifecycle, UI, services and persistence independently. Generate the
supported/partial/missing/platform-restricted reports from the same inventory.

## Per-PR handoff and verification

Each implementation PR must contain the specific inventory row IDs and Microsoft
links, owning layer, documented signature/options, intended profile, implementation
symbols, synthetic regression tests and an updated result ledger. Use original
authored workbooks and cases; no proprietary corpus or downloaded macros are
needed for the ordinary gates. Do not copy Microsoft's example code as fixtures.

Required behavior tests, scaled to each feature:

- Positive results with exact values/tags, object identity and observable state.
- Named/positional/omitted/default arguments, wrong receiver/kind and error paths.
- Invalid/foreign/stale/deleted/closed objects and captured versus active context.
- Protection, read-only/trust gates, budgets and zero partial writes on validation failure.
- Undo/recalc and persistence; appropriate event ordering/cancellation/footprints.
- Pause/wait/resume/Reset and callback counts proving effects are not replayed.
- Native and wasm compilation; actual WASM/WASI execution for shared portable behavior.
- Render/inspect affected workbook UI, print/export or graphics; an engine test
  does not prove the corresponding user-visible API works.

GridCraft: run focused adapter/engine/UI tests, `cargo +1.99.1 fmt --all -- --check`,
strict affected-target Clippy and `cargo +1.99.1 xtask ci` before a commit, with
VBA/window feature tests and wasm coverage explicitly enabled as applicable.
tiny-vba changes additionally require its documented CI/wasm/wasm-run and any
affected UI/native/browser gates. Keep network documentation retrieval outside
these gates; tests must work offline from the checked-in inventory/fixtures.

Report actual commands and results, NOT_RUN/unsupported services, exact remaining
rows and the next smallest packet. Microsoft reference-based tests establish a
declared profile; fresh Excel observations, when performed, are separate evidence.
Do not infer broad Excel parity from the absence of runtime error 438.

## Recommended first Luna assignment

Start with **OM00a + OM01a**: make the metadata inventory reproducible and
reconcile the current adapter, then expose a bounded host execution policy in the
shared IDE with tests and a synchronized GridCraft pin. Next do **OM02a**:
Destination-aware Copy/Cut/Paste and PasteSpecial, including their enums,
cross-document ownership and negative paths. Then **OM06a** Sort/SortFields plus
enum/collation contracts. Keep the rest of this plan as the ordered full-surface
backlog; do not turn it into a single large implementation change.

## Audit validation performed in this planning task

- Downloaded Microsoft's nontruncated public Git tree and pinned archive; SHA256
  `285218108a034fc33acc8296d714395fd182397ada27c848b93c3d5423efd75e`.
- Reconciled every one of the 5,996 Excel API source paths, classified headings
  and matched the current catalogue/selected generic dispatch/context paths.
- Inspected GridCraft object registry/dispatch, clipboard/formatting helpers,
  native edit/sort/file commands and tiny-vba IDE/VM budget construction.
- Research scope excludes workbook specimens, intrinsic implementations and
  worksheet-function mathematics. No application code or feature tests were
  changed for this plan; future packets still need their own verification.

The inventory is an auditable breadth baseline, with explicit uncertainty about
unreviewed parameter/platform semantics. It is not a claim that all 5,996 pages
have independently observed Excel behavior or completed implementation designs.
