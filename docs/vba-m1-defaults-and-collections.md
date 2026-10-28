# M1.2/M1.3: defaults and collection acceptance

Implemented 2026-10-28 on `feat/m1-defaults-and-collections`. The separate tiny-vba
repository is fetched by immutable Git revision
`d80e0393b408e0421a8df7e16667dbbbc970783b`; it is not embedded or copied here.
No sibling path override is used. Main branches are untouched.

GridCraft owns the Excel types, imports, catalogue, object-store variants,
default-member selection, collection ordering and document mutations. tiny-vba
owns only generic IR 6/default/iteration machinery. Runtime coercion is reused
through the same pinned vba-runtime crate. The optional L6 bridge and native
command/undo/recalculation path are preserved.

Supported additions:

- Worksheets and Sheets are actual typed collection objects with Count, Item,
  default Item, Application/Parent, name/index lookup and bounded For Each.
- Only supported worksheet kinds are exposed; no synthetic chart/macro sheets.
- Range default Get/Let forwards zero indexes to Value and indexes to Item.
  Item supports relative one/two indexes (including column letters); Range loops
  enumerate individual cells in row order. Set writes fail explicitly.
- Property assignment operands retain named/indexed calls through `$value`.
  Writes still use the engine's commands, protection and macro undo group.
- Enumeration preflights the shared 8,192-handle store before adding a snapshot;
  there is no second registry, retained document borrow or iterator token.

Original `collections.rs` tests cover unchanged typed source, values/order,
relative/indexed access, Count/Item/defaults, error numbers, stale handles,
atomic handle-budget refusal, safe paused watches, consent revocation and undo.
Existing workbook/project/UDF tests remain required. Validation results are
recorded below after all gates complete.

This is an independently implemented, documented subset, not an Excel oracle
observation or universal unchanged-XLSM claim. Range materialization, lightweight
whole-sheet references, Cells/Rows/Columns, stable sheet identity, events, forms
and package writeback retain their separately planned scope.

Public references: [Range](https://learn.microsoft.com/office/vba/api/excel.range(object)),
[Item](https://learn.microsoft.com/office/vba/api/excel.range.item),
[Cells](https://learn.microsoft.com/office/vba/api/excel.range.cells),
[Worksheets](https://learn.microsoft.com/office/vba/api/excel.worksheets),
[For Each](https://learn.microsoft.com/office/vba/language/reference/user-interface-help/for-eachnext-statement).
All source/tests are original; no Office files/assets or unlicensed specimens
were added. Earlier audit/refusal reports are historical evidence.

Validation completed on Windows with Rust 1.99.1:

- `cargo test -p gridcraft-vba`: 40 tests pass.
- The same suite under `wasm32-unknown-unknown`, wasm-bindgen-test-runner and
  Node 24.22.1: 40 tests pass with a successful runner exit.
- `cargo xtask ci`: all 10 gates pass; 2,568 test/doctest results.
- `cargo clippy -p gridcraft-ui-egui --features vba-window --all-targets -- -D warnings`: pass.
- `cargo test -p gridcraft-ui-egui --features vba-window`: 1,150 tests pass.
- `cargo check -p gridcraft-web --features vba --target wasm32-unknown-unknown`: pass.

The older local Node 24.14.0 crashed during runner shutdown after reporting
passing tests; that attempt is not counted as a passing gate. The complete suite
was rerun successfully with the existing Node 24.22.1 runtime. Browser compilation
retains existing dead-code warnings; strict native Clippy passes. Logs are in
ignored `target/m1-*.log`. The pure spreadsheet UDF regression also proves that
For Each over a two-dimensional value array recalculates after precedent changes.
