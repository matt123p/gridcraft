# Optional tiny-vba integration

Updated 2026-10-27. Profile: `GridCraft.Excel.v1`, over tiny-vba's M5 native
protocol and M8 shared IDE. This is a bounded first integration, not full Excel VBA.

Build the desktop with `cargo +1.99.1 run -p gridcraft --features vba` (or enable
`gridcraft-web/vba` for the browser). Default application builds keep VBA disabled.
View → Macros → VBA Editor opens the shared editor. Read VBA Project Source
decodes the original project without execution. No source change rewrites
`vbaProject.bin`; export edited modules as plain source instead.

The interpreter, debugger, source importer and editor remain in the separate
[tiny-vba repository](https://github.com/matt123p/tiny-vba), pinned by full Git
revision in Cargo manifests/lockfile. No sibling checkout or copied interpreter
code is needed. WordCraft can use the same native protocol and revision with its
own document objects. Workspace-root egui patches select the shared reviewed fork.
The pin is `4d80a9584b3daeb75d589c629eff272ecd5881fe`; corpus/reference
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
UDF registration, classes/UserForms, full collection/default-member syntax, Office
binary writeback/signatures or external IO/COM. Import rejects executable class/
document modules; inert document metadata remains in the preserved package.

## Validation and extension

`crates/vba/tests/workbooks.rs` owns real workbook acceptance, including undo,
recalc, errors, typed arrays, trust, lifecycle and package preservation.
`crates/ui-egui/tests/vba.rs` owns optional editor/control/accessibility acceptance.
Run both plus the ordinary workspace CI/WASM gates; default CI does not enable
optional features. Add any new Office member and its contract tests here, keeping
language semantics and editor code in tiny-vba.

The same adapter tests also use `wasm-bindgen-test` on wasm32-unknown-unknown.
Install `wasm-bindgen-cli` 0.2.129, set
`CARGO_TARGET_WASM32_UNKNOWN_UNKNOWN_RUNNER=wasm-bindgen-test-runner`, and run
`cargo test -p gridcraft-vba --target wasm32-unknown-unknown`. These execute in
JavaScript (Node), including GridCraft's browser clock imports. WASI is not a
supported GridCraft target. The optional CI job runs native, shared UI, strict
Clippy, browser compilation and JavaScript-hosted WASM acceptance explicitly.

The staged cross-repository evidence and resumption record lives in tiny-vba's
`docs/gridcraft-integration.md`.
