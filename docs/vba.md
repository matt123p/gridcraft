# Optional tiny-vba integration

Updated 2026-10-30. Profile: `GridCraft.Excel.v1`, over tiny-vba's M5 native
protocol and M8 shared IDE. This is a bounded first integration, not full Excel VBA.
The [Excel source-compatibility audit](vba-excel-compatibility.md) records corrected
members, remaining mismatches and stages toward running existing workbooks
unchanged. Arbitrary XLSM execution without source changes is not yet supported.

For a separate desktop editor window, run
`cargo +1.99.1 run -p gridcraft --features vba-window`.
For an embedded desktop editor, use `--features vba`. The `vba-window` feature
includes `vba`; it adds only native shell presentation, not another interpreter or
process. Browser builds use `gridcraft-web/vba` and always embed the editor.
Default application builds keep VBA disabled.

Developer → Visual Basic or Alt+F11 opens the editor; Macros → Edit selects a
procedure in it. Enable Developer under File → Options → Customize Ribbon →
Main Tabs. View → Macros → View Macros and Alt+F8 open the macro picker.
There are no permanent VB/Macros launch buttons on the Quick Access Toolbar;
only macros the user explicitly assigns appear there. Existing workbook source
is read without execution; new projects start blank rather than loading a demo.
No source change rewrites
`vbaProject.bin`; save source to the workbook's GridCraft storage or export modules
as plain source. The original binary project stays unchanged.

The interpreter, debugger, source importer and editor remain in the separate
[tiny-vba repository](https://github.com/matt123p/tiny-vba), pinned by full Git
revision in Cargo manifests/lockfile. No sibling checkout or copied interpreter
code is needed. WordCraft can use the same native protocol and revision with its
own document objects. Workspace-root egui patches select the shared reviewed fork.
The pin is `98c7038a297ca9453c01d7e4a5530e8fc8865a10`; corpus/reference
submodules are opt-in and are not fetched by dependency builds.

## Ownership and commands

`gridcraft-vba` is a toolkit-free L6 adapter above engine. The engine does not
depend on it. `ui-egui/vba` adds a small shell module which mounts `vba-ui-egui`
with existing GridCraft theme/font/scale tokens, source IO services and bounded
logic ticks. Workbook writes call the native `Session::set_range_contents` API
for the generic engine-owned `range.setContents` command;
recalc, protection, recording and history stay in the engine. Typed native VBA
values are converted only at that engine-command/workbook storage boundary;
JSON is not the member-dispatch or mutation protocol. Successful writes serialize
only their replayable journal/recording envelope. See [the native content API](native-cell-contents.md).

These UI-only command IDs go through `SheetApp::run` and `engine.execute` on the
desktop control channel. Programmatic calls do not open dialogs:

| Command | Parameters / behavior |
|---|---|
| `vba.open` | `{source?: string, project?: bool, separateWindow?: bool, entry?: string}`; open/focus shared project, optionally navigate to a procedure; never executes |
| `vba.compile` | Compile current standard modules |
| `vba.run` | `{entry?: string}`; Run/Continue, bounded by subsequent logic ticks |
| `vba.stepInto`, `vba.stepOver`, `vba.stepOut` | Same shared debugger actions as toolbar |
| `vba.pause`, `vba.reset` | Pause or stop; retain prior workbook writes |
| `vba.status` | Mode, output and diagnostics |
| `vba.showDeveloper` | `{on: bool}`; persisted ribbon preference, initially hidden |
| `vba.apply` | Save current standard-module source to workbook and register its Functions |
| `vba.entries` | Discover public no-argument Subs and public Functions, with signatures |
| `vba.invoke` | `{entry: string}`; explicitly run a public no-argument Sub from an assignment |
| `vba.assign` | `{entry, id?: shape id, at?: cell, caption?: string}`; assign existing shape or create a button at the selection/`at` |
| `vba.pin` | `{entry, remove?: bool}`; add/remove a workbook macro on the Quick Access Toolbar |
| `vba.designMode` | `{on: bool}`; select/move assigned shapes instead of executing them |
| `vba.trust` | `{allow: bool}`; explicit consent for this editor/document instance |
| `vba.return` | Return to the workbook, keeping source/debugger state |
| `vba.close` | `{discard?: bool}`; refuse dirty source unless explicitly discarded |

The shared editor's Save/Export writes plain source through the shell's existing
file picker/download service. Import adds UTF-8 source. Browser source import can
also use `vba.open {source: ...}`. This feature does not install a second editor.

## Implemented object profile

| Object | Supported members |
|---|---|
| Application | Application/Parent, ThisWorkbook/ActiveWorkbook, ActiveSheet, Worksheets/Sheets, Range, Cells, Rows, Columns |
| Workbook | Name, Application/Parent, ActiveSheet, Worksheets/Sheets |
| Worksheet | Name, Application/Parent, Range(Cell1, optional Cell2), Cells, Rows, Columns, UsedRange |
| Range | Item/default Item, Cells/Rows/Columns, Areas, Row/Column/Count/CountLarge, relative Range, Offset/Resize, EntireRow/EntireColumn, End, Find/FindNext/FindPrevious, Insert/Delete; Value/Value2/Formula Get/Let and default Value, ClearContents, absolute Address; Application/Parent/Worksheet |
| Worksheets/Sheets | Count, Item(index/name), default Item, bounded For Each, Parent/Application |
| Areas | Count, Item/default Item, bounded For Each, Parent/Application |

References are lightweight, including whole worksheets. Cells supports zero/one/two
indexes, Excel Item argument names, column letters and shared VBA Long coercion.
Rows/Columns expose subrange collections; Cells returns cell collections. Worksheet
Range accepts A1, quoted sheet qualification, workbook/local reference names,
colon, disjoint comma areas and intersections. Application.Range can resolve another
worksheet; Worksheet/Range endpoints enforce ownership. Range.Range offsets literal
addresses from its first area's origin; Range object endpoints remain absolute.
Offset/Resize use checked bounds. Noncomputed defined names resolve with local-scope
precedence and bounded recursion. See [M3 acceptance and exact boundaries](vba-m3-ranges-and-worksheets.md).

Registered nominal types include Application, Workbook, Worksheet, Range,
Worksheets, Sheets and Areas. Wrong classes fail with 13; stale/foreign/closed
references fail with 91. Structural commands rebase existing references and deleted
references fail with 1004. [M1.1 acceptance](vba-m1-types.md) records the generic type boundary.

Imports include Application, ThisWorkbook, ActiveWorkbook, ActiveSheet, ActiveCell,
Selection, Worksheets/Sheets, Range and Cells/Rows/Columns. Getter aliases resolve
on each source access, while captured objects retain workbook/sheet ownership.
ThisWorkbook is the code owner; active roots follow the visible application
workbook. Scoped native targets restore the visible window after calls. Stable
sheet identities preserve handles through rename/reorder; deleted sheets and
closed workbooks fail explicitly. Consent remains per workbook.

Value reads infer Date/Currency from formatting; Value2 retains Double. Typed
Value writes respect date systems and atomically apply default formats to General
cells, with Currency storage rounded to cents. Value2 preserves numeric storage.
Reads return 1-based 2D Variant arrays; writes broadcast scalars or crop/fill a 2D
array from arbitrary lower bounds (uncovered cells get CVErr(2042)). General
strings use the engine's en-US input conversion; Text-formatted Value/Value2 text
stays literal. Formula explicitly parses leading `=` even in Text cells.

Address implements its five optional arguments and A1/R1C1, relative and external
forms. Live NumberFormat, Font Bold/Italic/Name/Size/Color and Interior.Color use
bounded protected native commands. Calculation settings and Calculate have real
engine effects. Unsupported settings/members refuse explicitly; ScreenUpdating
remains unsupported, while `EnableEvents` has the bounded M5 behavior below. See the
[M4 contract and acceptance](vba-m4-identity-values-and-members.md) for exact
boundaries, immutable dependency pin and executable evidence.

Limits: 8,192 interned object handles; 64 disjoint areas; 65,536 materialized cells
per read/write; 8 MiB value payload; 65,536 stored cells per End/Find scan, one million
scanned entries and eight million wildcard matching steps per macro;
tiny-vba's project/string/runtime budgets also apply. Unknown members fail with
VBA error 438. Invalid indices/ranges, type errors, stale references and trust
failures return explicit numbered errors and are visible to VBA handlers.

## Lifecycle and boundaries

File provenance requires explicit consent before **any** execution, including
host-free code. Every tick and every late-bound host invocation rechecks trust
and document identity. Script trust does not grant VBA consent or change file
provenance. There is no automatic open-event execution.

Each macro/event batch uses a spanning engine undo group, retained through
break/step/resume. Completion and failure finalize successful writes as one undo
entry; failures do not roll back earlier effects. Reset, close and trust revocation
cancel queued callbacks. No mutable workbook borrow survives a host call or tick.
The bridge remaps retained sheet references by stable sheet identity after topology
changes; deleted or stale owners fail explicitly. Switching away from the workbook
bound to this editor cancels active execution/event delivery.

The implemented M5 slice supports inert import of standard, class and document
modules, explicit trusted Workbook_Open, bounded Worksheet_Change delivery with
EnableEvents, and the Range-aware worksheet-function surface documented below.
The shared M2 designer/runtime can run `.tvform` forms through the workbook host;
forms remain separate exchange files rather than workbook VBA metadata. The profile
still lacks structural sheet/range mutations, cross-document objects, the complete
Office event/catalogue surface, VBA binary project writeback/signature handling and
external IO/COM. This remains a bounded compatibility projection, not general XLSM
execution.

## Spreadsheet workflows

1. Build with `--features vba-window` for the separate desktop window or
   `--features vba` for embedded mode. Developer → Code → **Visual Basic**, or
   **Alt+F11**, opens/focuses the same shared editor. File → Options → Customize
   Ribbon → Main Tabs → **Developer** exposes the initially hidden, persisted tab.
   Developer → Code → **Macros**, View → Macros → **View Macros**, and **Alt+F8**
   open the selectable procedure list. Choose **Edit** or **Step Into** to open
   the selected procedure; **Run** executes it without opening the editor.
2. Edit standard modules in the shared editor. **Save Source to Workbook** applies
   source and registers public Functions; the workbook must then be saved normally.
   Source is chunked through existing hidden-name storage, not a new VBA binary.
   Opening/importing source never runs it. For file provenance, give explicit
   session permission before running macros or calculating VBA functions.
3. In Macros choose a public no-argument Sub and **Run**, **Edit** or **Step Into**.
   Functions are excluded, as in Excel; they appear under Insert Function instead.
   GridCraft-specific **Assign Button** / **Add to Toolbar** actions are under the
   collapsed **GridCraft customizations** section. Developer → Controls → Insert
   → Button opens the assignment picker with an **Assign** action. New buttons
   are placed at the active cell (or `vba.assign.at`).
   Right-click an existing shape → Assign Macro to assign/reassign it. Ordinary
   clicks run its Sub; Design Mode allows selection, movement, resizing and deletion.
   Source plus button insertion/assignment is one undo step, using engine commands.
4. Public Functions are available in cell formulas and Insert Function's **User
   Defined** results after applying source. For this independently authored example,
   `=Triple(A1)` returns 21 when A1 is 7; normal dependency recalculation updates it:

```vb
Public Function Triple(ByVal number As Double) As Double
    Triple = number * 3
End Function
```

### Window ownership and launch overrides

With `vba-window`, the desktop shell creates one immediate native viewport after
rendering workbook windows. Workbook editing stays visible and usable; the child
renders the same `EditorUi`, `Ide` and `Bridge`. Only parent logic advances the VM.
Macros launched from the worksheet bring the editor forward on a breakpoint or
unhandled runtime failure.
Repeated launch focuses the existing window. Alt+F11 or the editor's Return action
returns focus to the workbook. Closing the editor's OS window hides it and retains
source, breakpoints and state; closing the project or workbook still observes the
existing dirty-source and cancellation guards. Native viewport state is excluded
from the WASM build, which retains in-page presentation.

`vba.open {"separateWindow": false}` explicitly embeds even in a window-enabled
desktop. `true` requires `vba-window` and a native window host; otherwise it returns
an error. Omitting it chooses the build/host default on first launch and retains
the current project's mode when reopening. Switching modes never creates another
project or VM. `vba.status` reports `separateWindow` and `visible`.

The desktop feature uses the existing egui viewport backend, not IPC, filesystem
source transfer, or a copied editor. `vba_windows` tests exercise nested native
viewport input/close/focus and assert rendering does not run another VM slice.

Functions accept scalar values and bounded rectangular reference values as
one-based two-dimensional Variant arrays; return scalars or rectangular arrays
through the normal spill engine. Ordinary ByRef parameters use owned temporaries,
and omitted trailing optional parameters use the language compiler's defaults.
Public Subs with parameters, Private procedures and Option Private Modules are
excluded from the picker. Built-in worksheet functions retain precedence.

The calculator exposes only the generic `calc::CustomFunctions` native provider,
installed on one open document; no VBA dependency enters calc/engine. The adapter
compiles independent wrapper procedures to retain the shared ByVal entry contract
and VBA's own argument coercion/default handling. Each call has a fresh bounded VM:
100,000 instructions, depth 64, 8 MiB payload; a pass shares one million instructions.
Source is limited to 64 modules/4 MiB encoded storage; 256 functions, 32 MiB compiled
estimate, 64 arguments, 65,536 cells per reference, 256 assignments per sheet and
32 toolbar assignments. Exhausted/malformed calls return `#VALUE!` without writes;
explicit supported CVErr values retain their worksheet errors.

UDFs cannot invoke workbook operations, prompts or asynchronous host calls.
The accepted M5 profile passes arguments declared `Range` or `Object` as read-only host Range
objects. `Variant` and scalar parameters keep their normal value/array coercion while
the calc bridge retains the original reference provenance. The profile also supports
`Application.Caller`, `Application.Volatile`, workbook value reads and persistent
project state. The read-only caller/range context exposes Value/Value2,
Address, Row/Column, Count, Worksheet, Offset and Resize; worksheet Name/CodeName
and parent workbook CodeName are readable. Cross-workbook/add-in registration and
explicit interior omitted arguments remain unsupported. Consent
is rechecked/revocable; closing the editor project or invalidating its document
disables registered functions. Reopening stored source is inert; Apply or explicit
consent re-registers it. Replaced/undone source cannot execute a stale compiled
version: apply the desired source again. Sheet handles follow stable identities
through supported topology changes; deleted owners/references fail instead of
retargeting by ordinal. Authored acceptance confirms Caller/Row/Column, Address,
Offset/Resize and Date/Currency `Value` typing versus numeric `Value2`; Currency
coercion retains four decimal places. Shape-owned assignments survive sheet renames.

`Worksheet_Change` currently covers VBA Range writes and native `cell.set`,
`range.setContents`, `range.setValues`, `range.fill`, contents clear, Clear All,
cut/paste, AutoFill, Series fill, directional fills, Flash Fill, sorting, removing
duplicates and Text to Columns. Text to Columns reports its committed output range
to the host event hook. Formats-only paste/fill operations are suppressed.
The reviewed native mutation list is covered by authored event-source tests, including
cell-editor commit, direct-control edits and pointer drag-cut delivery for both committed ranges.
Format-only operations and no-op duplicate removal do not queue events.
If a native edit commits but the bounded event queue is full, the edit remains
committed and GridCraft reports that the event could not be queued. A VBA Range
write checks the event budget before mutating and refuses before the write.
If a callback fails, remaining callbacks in the queued event batch are cancelled;
successful writes already made remain in the grouped undo step. Reset, close and
trust revocation also cancel queued callbacks before releasing the host.
M5 document-event integration is accepted for the declared profile. Workbook and worksheet code names
are imported inertly; `vba.workbookOpen` explicitly queues `Workbook_Open` after
host authorization, including from Design mode without first running a separate macro.
The IDE creates a verified idle session and delivers the handler through its normal
debugger path. Worksheet callbacks scope
unqualified Range/Cells/Rows/Columns construction to the owner sheet without
changing active selection. `Me.Range` and declared sheet/workbook members dispatch
through the captured live host owner; unsupported members still fail explicitly.
GridCraft's embedded IDE imports `.tvform` files through File → Import File (also
available as the `vba.importSource` UI command) and runs the selected Label/TextBox/
CommandButton profile through tiny-vba's shared form designer/runtime. TextBox and
button callbacks use the workbook bridge; authored integration evidence covers cell
writes and modal completion/close without a parked wait token. This does not add Office
binary UserForm resources or workbook VBA project persistence for forms.
Event notifications enter a FIFO queue and are delivered only after the current
invocation/callback reaches a boundary; callback writes append later events instead
of re-entering VBA. `EnableEvents = False` suppresses new notifications and discards
the queued remainder; setting it True resumes future notifications. Each event chain
is capped at 256 queued/delivered events. Reset/close/trust revocation and callback
failure discard the remaining queue; an event is removed before dispatch and is never
retried. A pause retains the undelivered tail until resume or cancellation.
The complete scenario status and exact evidence are tracked in tiny-vba's
[M5 acceptance map](../../tiny-vba/docs/m5-acceptance.md). The profile does not claim
complete Excel compatibility: unsupported catalogue members, structural mutations,
binary form resources and VBA project writeback remain outside scope.

Buttons use original GridCraft shapes with inert `on_action` metadata.
They implement the documented click/Assign Macro workflow, not Office Form Control
or ActiveX serialization. `shape.setMacro {id, entry}` stores/removes an assignment
without executing code. DrawingML's `macro` attribute and JSON preserve it on the
shape itself; deletion and ID reuse cannot transfer it to an unrelated object.
Source/toolbar data round-trip through GridCraft XLSX/JSON. Excel needs a real VBA
project containing matching procedures: saving GridCraft source does not update
that binary. Binary project writeback remains outside the shared import contract.

### Public behavior references and provenance

Consulted 2026-10-27: Microsoft's public tutorials for [running macros](https://support.microsoft.com/en-us/excel/run-a-macro-in-excel),
[assigning worksheet buttons](https://support.microsoft.com/en-us/excel/assign-a-macro-to-a-form-or-a-control-button),
[custom VBA functions](https://support.microsoft.com/en-us/excel/create-custom-functions-in-excel),
[keyboard shortcuts](https://support.microsoft.com/en-us/accessibility/excel/keyboard-shortcuts-in-excel)
and [Developer visibility](https://support.microsoft.com/en-us/office/add-ins/show-the-developer-tab).
Shape metadata also follows public [OnAction](https://learn.microsoft.com/en-us/office/vba/api/excel.shape.onaction)
and [DrawingML macro](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.spreadsheet.shape.macro)
documentation.
These are behavior references only. No tutorial code, text, images, Office binaries
or external workbook fixtures were copied. Implementation, icons and test sources
are independently authored; existing dependency licences remain unchanged. No
fresh Microsoft oracle observation or complete Excel compatibility is claimed.

## Validation and extension

Window/launch correction acceptance (2026-10-28): `cargo xtask ci` passes all ten
steps (2,553 tests). The full `vba-window` UI suite passes 1,150 tests, including
21 VBA workflow and four native-viewport cases; the embedded-only workflow suite
also passes. Strict all-target Clippy passes for embedded and window-enabled
desktop builds. Browser application compilation and a WASM build with the window
feature selected both pass; the latter still excludes native state and embeds.
The real Windows desktop build was launched in an isolated, preferences-disabled
session: distinct workbook/editor native windows were visually inspected, return
to the workbook retained the project, and the smoke instance was closed afterward.
Nine original light/dark/narrow renderings under ignored `target/vba-window-ui`
were regenerated and reviewed. Logs are `target/vba-window-*`. No hosted CI run
or manual macOS/Linux window test is claimed by this local evidence.

`crates/vba/tests/workbooks.rs` owns real workbook acceptance, including undo,
recalc, errors, typed arrays, trust, lifecycle and package preservation.
`crates/vba/tests/worksheet.rs` covers UDF defaults, owned ByRef inputs, dependencies,
scalar/errors/array spills, budgets, private scope and inert source/shape persistence.
Repeated shared strings exceeding 32-bit logical byte totals are rejected before
VBA UTF-16 allocation, with saturating payload accounting on native and WASM paths.
`crates/engine/tests/shape_macros.rs` covers assignment protection, escaping,
XLSX/JSON round trips and undo/removal; the authored ink case preserves inert
imported action metadata without losing its geometry.
`crates/ui-egui/tests/vba.rs` owns optional editor/control/accessibility acceptance.
Run both plus the ordinary workspace CI/WASM gates; default CI does not enable
optional features. Add any new Office member and its contract tests here, keeping
language semantics and editor code in tiny-vba.

Final acceptance evidence (2026-10-27): the complete ten-step GridCraft gate,
the full optional UI suite (1,143 tests including 18 VBA interaction scenarios),
strict optional all-target Clippy, actual desktop/browser builds and 25 adapter
cases (2 project, 13 workbook, 10 worksheet) pass. Adapter cases execute both natively
and in JavaScript-hosted WASM. Nine original editor/worksheet/picker PNGs were
rendered in light, dark and narrow/scaled modes and visually reviewed in ignored
`target/worksheet-ui`. The separate tiny-vba merged revision passes CI, WASM and
WASI execution. Logs are ignored under `target/worksheet-final-*`; this is local
acceptance, not a claim that the edited hosted CI workflow has already run.

The same adapter tests also use `wasm-bindgen-test` on wasm32-unknown-unknown.
Install `wasm-bindgen-cli` 0.2.129, set
`CARGO_TARGET_WASM32_UNKNOWN_UNKNOWN_RUNNER=wasm-bindgen-test-runner`, and run
`cargo test -p gridcraft-vba --target wasm32-unknown-unknown`. These execute in
JavaScript (Node), including GridCraft's browser clock imports. WASI is not a
supported GridCraft target. The optional CI job runs native, shared UI, strict
Clippy, browser compilation and JavaScript-hosted WASM acceptance explicitly.

The staged cross-repository evidence and resumption record lives in tiny-vba's
`docs/gridcraft-integration.md`.

2026-10-29: M3 adds bounded range geometry, sparse navigation/search and native
structural command coverage; see the M3 acceptance report for validation and gaps.

The M3 dependency update includes the shared M2 class/form editor. Current workbook
source persistence explicitly refuses classes/forms before effects; export them
separately. Dirty design state blocks implicit close. Workbook form storage remains
outside this profile; M5 document/event execution is partially implemented below.

2026-10-29: M4 adds stable document/sheet ownership, dynamic active roots, typed
Value conversion, Address arguments and live formatting/calculation. The shared getter prerequisite is published and pinned by immutable Git revision;
the temporary local override is removed.
