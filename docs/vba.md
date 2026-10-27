# Optional tiny-vba integration

Updated 2026-10-27. Profile: `GridCraft.Excel.v1`, over tiny-vba's M5 native
protocol and M8 shared IDE. This is a bounded first integration, not full Excel VBA.

Build the desktop with `cargo +1.99.1 run -p gridcraft --features vba` (or enable
`gridcraft-web/vba` for the browser). Default application builds keep VBA disabled.
View → Macros → VBA Editor opens the shared editor. Read VBA Project Source
decodes the original project without execution. No source change rewrites
`vbaProject.bin`; save source to the workbook's GridCraft storage or export modules
as plain source. The original binary project stays unchanged.

The interpreter, debugger, source importer and editor remain in the separate
[tiny-vba repository](https://github.com/matt123p/tiny-vba), pinned by full Git
revision in Cargo manifests/lockfile. No sibling checkout or copied interpreter
code is needed. WordCraft can use the same native protocol and revision with its
own document objects. Workspace-root egui patches select the shared reviewed fork.
The pin is `ea38c3f346b4010d6017b3a6d27e645ebad7ef79`; corpus/reference
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
| `vba.open` | `{source?: string, project?: bool}`; default authored FillCells demo; never executes |
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
| Application | Application, ThisWorkbook, ActiveWorkbook, ActiveSheet, Worksheets(index/name), Sheets(index/name), Range(address), Cells(row, column) |
| Workbook | Name, Application, ActiveSheet, Worksheets(index/name), Sheets(index/name), Range, Cells |
| Worksheet | Name, Range(local rectangular A1 address), Cells(1-based row, column) |
| Range | Get/Let Value, Value2, Formula; ClearContents; absolute Address |

Unqualified Application, ThisWorkbook, ActiveWorkbook, ActiveSheet, Worksheets,
Range and Cells are M5 imports. Object names/member names are case insensitive;
named arguments bind explicitly. ActiveWorkbook is confined to this editor's
workbook; the unqualified ActiveSheet root is captured when the editor opens.
Application.ActiveSheet and unqualified Range/Cells use the current active sheet
in that workbook. The profile does not allow cross-workbook effects.

Multi-cell reads return 1-based two-dimensional Variant arrays. Writes accept
matching two-dimensional arrays with arbitrary lower bounds or broadcast a
scalar. Shape/type conversion is validated before mutation. Values preserve
literal text (including text beginning `=`), Empty, booleans and supported CVErrs.
Worksheet numbers are IEEE doubles: Date/Currency become numeric values; Value
currently has the same behavior as Value2 (no number-format-based Date/Currency
inference). Integers outside the exact double range, Null/Missing/Nothing/objects,
unpaired surrogates, unknown CVErrs and unsupported cell error kinds are refused.
Text-format (`@`) cells support literal string and CVErr writes, and explicit
Formula writes calculate while preserving the format. Value/Value2 strings
remain literal; Formula strings beginning `=` use the explicit formula path.

Limits: 8,192 interned object handles; 65,536 cells per range; 8 MiB range payload;
tiny-vba's project/string/runtime budgets also apply. Unknown members fail with
VBA error 438. Invalid indices/ranges, type errors, stale references and trust
failures return explicit numbered errors and are visible to VBA handlers.

## Lifecycle and boundaries

File provenance requires explicit consent before **any** execution, including
host-free code. Every tick and every late-bound host invocation rechecks trust
and document identity. Script trust does not grant VBA consent or change file
provenance. There is no automatic open-event execution.

Each run uses a spanning engine undo group, retained through break/step/resume.
Completion, error, reset and close finalize successful writes as one undo entry.
Errors/stop do not roll back prior successful effects. No mutable workbook borrow
survives a host call or tick. Document close/switch or sheet topology changes cancel
the execution and invalidate the bridge. Reopen the editor to bind new references.
The host performs no asynchronous operations, so there are no outstanding wait
tokens to drain. The generic interpreter wait/callback contract remains available.

Worksheet model objects lack a session-stable identity; this profile rejects
observed topology changes rather than retargeting by ordinal. It does not provide
structural sheet/range mutations, cross-document objects, events/EnableEvents,
classes/UserForms, full collection/default-member syntax, Office
binary writeback/signatures or external IO/COM. Import rejects executable class/
document modules; inert document metadata remains in the preserved package.

## Spreadsheet workflows

1. Build with `--features vba`. Use the toolbar **VB** button or **Alt+F11** to
   open/toggle the editor. **Alt+F8**, the toolbar **Macros**, and View → Macros
   open the procedure picker. View → **Developer tab** exposes the persisted
   Developer ribbon with Visual Basic, Macros, Save Source to Workbook, Insert
   and Design Mode. Developer is initially hidden, as in Excel; the View location
   of its preference is GridCraft's own convention.
2. Edit standard modules in the shared editor. **Save Source to Workbook** applies
   source and registers public Functions; the workbook must then be saved normally.
   Source is chunked through existing hidden-name storage, not a new VBA binary.
   Opening/importing source never runs it. For file provenance, give explicit
   session permission before running macros or calculating VBA functions.
3. In Macros choose a public no-argument Sub and **Run**, **Edit**, **Assign Button**
   or **Add to Toolbar**. Developer → Insert → Button opens the same assignment
   picker. New buttons are placed at the active cell (or `vba.assign.at`).
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
This profile requires inputs as arguments: Range objects, Application.Caller,
Application.Volatile, workbook reads inside functions, persistent UDF state,
cross-workbook/add-in registration and explicit interior omitted arguments are
not supported. Reference arguments are values, not Excel Range objects. Consent
is rechecked/revocable; closing the editor project or invalidating its document
disables registered functions. Reopening stored source is inert; Apply or explicit
consent re-registers it. Replaced/undone source cannot execute a stale compiled
version: apply the desired source again. Sheet-name topology changes invalidate
the editor, which must be reopened; shape-owned assignments survive sheet renames.

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
