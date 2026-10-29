# M4 — Excel identity, values and member semantics

Updated 2026-10-29. Owner: GridCraft. Shared getter binding is tiny-vba ADR 0025,
published commit `22cc845e5ee74a1bcd15eec32cdb8a8f18d316f8` on
`feat/m4-getter-alias-dispatch`. Every tiny-vba dependency and lockfile entry uses
that immutable Git revision. The temporary development override is removed.
M3 and M4 are complete for the profiles documented here.

## Implemented profile

Source save/assignment remains restricted to the active code-owner workbook,
even when execution uses another document. The existing optional-window regression
caught and now guards this boundary.

M4.1: sheet identities are process-local Arc tokens preserved through COW,
rename, reorder and undo snapshots. A copied/imported sheet receives a fresh
identity. These tokens do not enter serialized files or content equality. Object
handles retain workbook UID plus sheet identity mapping; deleted sheets and closed
workbooks cannot retarget a replacement index. ThisWorkbook is the code owner;
ActiveWorkbook/ActiveSheet, Cells/Rows/Columns, Range and collection roots resolve
the current application document at each source access. Assigned object variables
capture concrete results. ActiveCell/Selection return actual current geometry.

Calls use the engine's existing scoped document targets and restore the visible
window. Every target enforces consent separately; code-owner consent cannot grant
access to another untrusted workbook. The code owner must remain open and trusted.
Each touched workbook gets its own bounded undo group; earlier successful effects
survive a later VBA error, and finish/reset closes groups in every workbook.

M4.2: Value reads infer Date or Currency from number formats using shared VBA
binding; Value2 returns Double. Date reads/writes respect the 1900/1904 systems.
Excel's fictitious 1900 leap day is explicitly refused as Date. Typed Value Date
and Currency writes atomically apply a default date/currency format only to
General cells; Currency Value storage rounds to cents. Value2 keeps numeric
storage without adding a Date/Currency format. Exact CVErr tags remain typed.

Reads are 1-based 2D Variants. Writes broadcast scalars or map a 2D array from its
own lower bounds, crop surplus elements and fill uncovered destination cells with
CVErr(2042). All conversions and format payloads validate before the native commit.
General-format strings use the engine's en-US typed-input conversion, including
numbers, dates, leading-apostrophe text and explicit formula strings. Text-formatted
cells retain literal Value/Value2 text. Formula explicitly parses leading `=` even
in Text cells. Formula validation failures leave contents/formats unchanged.

Address binds RowAbsolute, ColumnAbsolute, ReferenceStyle, External and RelativeTo
with named/omitted argument checks. It renders absolute/mixed/relative A1/R1C1,
whole rows/columns and disjoint areas without materializing cells. Relative R1C1
defaults to A1; supplied origins must be same-sheet Ranges. External references
carry the current workbook title and sheet name with conditional apostrophe
quoting/escaping. Saved filesystem path/closed-external-workbook addressing is not
in this profile.

M4.3: live NumberFormat, Font (Bold/Italic/Name/Size/Color) and Interior.Color use
bounded native style commands with protection, recording/replay and undo. Mixed
format reads return Null. Worksheet.Name uses the native rename implementation,
including formula/name adjustments and workbook structure protection. Calculation
supports Automatic/Manual/AutomaticExceptTables with real engine effects across
open workbooks and inheritance for newly opened workbooks. Application.Calculate
calculates all open workbooks; Workbook.Calculate targets its actual owner.

The catalogue advertises only these implemented signatures. ScreenUpdating and
EnableEvents are marked unsupported and fail with 438; no successful placeholder
settings are supplied. Event suppression belongs to M5's event contract. Further
Font/Interior members, Borders, Activate/Select and unselected settings remain
unsupported rather than implied by their object type.

## Lifetime and resources

M3 materialization/search/handle budgets are retained. Styles and per-cell formats
are bounded; typed content and format changes share one native atomic command.
External row/column/cell structural edits and undo/redo invalidate saved range,
area and formatting coordinates with 1004 instead of silently retargeting them.
Native VBA structural edits still preflight and rebase saved handles. Sheet rename
and reorder preserve handles; a deleted sheet gives 91. External cut/paste and
arbitrary third-party workbook replacement are not a separately verified live
reference-maintenance profile. Source/debugger edits retain existing VM contracts.

## Acceptance and evidence

`crates/vba/tests/identity_values.rs` contains ten independently authored unchanged
source workflows: exact addresses and error numbers, Date/Currency tags, 1904 date
writes, arrays/bounds/error fill, real sheet reorder/deletion, active-document
switching, closed index reuse, consent, protection, undo and calculation effects.
`crates/engine/tests/styles.rs` asserts native/envelope replay parity, undo and
malformed/protected refusal atomicity; M3's structural tests remain required.

Validation: `cargo test -p gridcraft-vba` passes 58 tests; the identical suite under
`wasm32-unknown-unknown` and `wasm-bindgen-test-runner` passes 58 actual JavaScript/
WASM tests using Node 22.23.4 and runner 0.2.129. Native structural/style tests pass.
Full `cargo xtask ci` passes all ten gates (2,593 workspace test/doctest results).
Strict optional-window Clippy and `cargo check -p gridcraft-web --target
wasm32-unknown-unknown --features vba` pass. The corrected optional-window suite passes all 1,150 tests. The final full CI
rerun passes all ten gates against the completed code. `CRAFT_FONTS_DIR=/tmp/gridcraft-acceptance-fonts` uses the
optional font checkout `aced4137c022f1f587624d48e81ec77e1914a9cc`; no font bytes
enter this repository. Initial missing-glyph and mismatched-runner setup failures
were corrected; they are not counted as passes. The final pinned native/JavaScript-WASM/CI/window checks run without Cargo path
overrides. Downloaded tools/fonts remain external development inputs. No fresh Office oracle run, imported upstream
macro execution, universal Excel compatibility or package-writeback claim is made.

## Public references

Microsoft documentation describes [Value](https://learn.microsoft.com/office/vba/api/excel.range.value),
[Value2](https://learn.microsoft.com/office/vba/api/excel.range.value2),
[Address](https://learn.microsoft.com/office/vba/api/excel.range.address),
[ThisWorkbook](https://learn.microsoft.com/office/vba/api/excel.application.thisworkbook),
[ActiveWorkbook](https://learn.microsoft.com/office/vba/api/excel.application.activeworkbook),
[Calculation](https://learn.microsoft.com/office/vba/api/excel.application.calculation)
and [Calculate](https://learn.microsoft.com/office/vba/api/excel.application.calculate).
These are public behavior references, not locally observed Office results.

## Final immutable-pin gate

PASS against published `22cc845`, without any local Cargo override:
`cargo test -p gridcraft-vba` (58 tests), actual JavaScript/WASM adapter tests
(58), `cargo xtask ci` (all ten gates; 2,593 test/doctest results), and
`cargo test -p gridcraft-ui-egui --features vba-window` (1,150 tests).
Logs are ignored under `target/m4-validation/gridcraft-m4-pinned-*.log`.
Only the 13 intended shared Git packages changed in Cargo.lock. M3/M4 acceptance
is closed; M5 document modules/events is the next implementation milestone.
