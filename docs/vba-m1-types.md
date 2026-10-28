# M1.1 — Registered Excel nominal types

Updated 2026-10-28. Branch: `feat/m1-host-types`. This report follows tiny-vba's
new remaining-plan numbering. It does not claim a complete Excel object library
or universal unchanged-workbook compatibility.

## Ownership

tiny-vba remains an independent Git dependency, pinned by full revision in both
VBA/UI manifests and Cargo.lock. It owns generic nominal identity, reference
resolution, verified IR 5 metadata and runtime validation. It has no Excel or
GridCraft object definitions, document store, command adapter or reverse dependency.

GridCraft owns `crates/vba/src/types.rs`: the Excel reference descriptors and the
resolver for Application, Workbook, Worksheet and Range. These are the four real
object kinds already implemented in this adapter. The reference version is the
`GridCraft.Excel.v1` adapter profile, not an installed Office typelib version.

The bridge's existing object vector is now one bounded shared store. Type queries
read that same store; there is no shadow registry. Host/object/document generation
checks precede lookup. Query locks end before document commands or host calls.
Bridge close/drop invalidates queries; ordinary tick/close lifecycle checks remain
mandatory when the document, trust or topology changes. Workbook edits still use
native engine commands and the existing undo/recalculation policy.

## Behavior and evidence

Qualified/unqualified declarations such as `As Worksheet` and `As Excel.Range`,
typed parameters/results and Nothing now compile and execute against registered
objects. Wrong classes fail with VBA error 13; stale/foreign/closed handles fail
with 91. Object/Variant late binding remains. Unregistered types remain explicit
compile errors; no COM/DLL/type-library loading or automatic execution is added.

| Requirement | Original fixture / assertion |
|---|---|
| Unchanged Excel source executes real effects | `crates/vba/tests/nominal.rs`: Application/Workbook/Worksheet/Range declarations, ByVal/ByRef, typed Function result, actual A1 write of 42 and exact retained source |
| Wrong class does not retarget a handle | Same suite: Range assigned to Worksheet fails with 13; A1 remains empty and earlier B1 write remains committed/undoable |
| Paused edits obey declared type | Same suite: Worksheet cannot replace a Range local; original reference and resumed A1 write remain correct |
| Resolver sees actual shared store/lifetime | `types.rs` unit test: newly interned Range, wrong class, foreign host, stale generation and closed-store queries |
| Unsupported types still fail | `workbooks.rs`: unregistered Excel.Chart and missing library; collection/Cells/Address gaps remain separate refusals |
| Existing source/import/UDF/lifecycle behavior | All adapter workbook/project/worksheet suites; worksheet compilation uses the same reference snapshot while value-only UDF policy remains |

Final native and actual Node/wasm-bindgen runs pass all **33 adapter tests**.
Full GridCraft CI passes all ten steps (**2,561 tests** across its test/doctest
results); optional desktop-window tests (**1,150**) and strict all-target window
Clippy pass. The browser VBA feature compiles; its existing `Modified` dead-code
warning remains unrelated. Final dependency-pin verification is recorded below.

Final dependency: tiny-vba
[`98a88c37ac125cf4de65d42c0e9d9120e731e49d`](https://github.com/matt123p/tiny-vba/commit/98a88c37ac125cf4de65d42c0e9d9120e731e49d).
Both manifests and every tiny-vba Cargo.lock entry use that revision, with no
path override or embedded interpreter checkout. All gates above were rerun
against this exact pin. Final adapter tests also assert completion for qualified
Worksheet/Range variables; catalogue library identity is `Excel`, separate from
the adapter profile version. M1.1 fork acceptance is complete; next is M1.2.

Commands use Rust 1.99.1 and the optional licensed `target/craft-fonts` checkout:

```text
cargo +1.99.1 test -p gridcraft-vba
cargo +1.99.1 test -p gridcraft-vba --target wasm32-unknown-unknown
cargo +1.99.1 xtask ci
cargo +1.99.1 clippy -p gridcraft-ui-egui --features vba-window --all-targets -- -D warnings
cargo +1.99.1 test -p gridcraft-ui-egui --features vba-window
cargo +1.99.1 check -p gridcraft-web --features vba --target wasm32-unknown-unknown
```

Actual WASM tests use the matching wasm-bindgen-test-runner 0.2.129 via
`CARGO_TARGET_WASM32_UNKNOWN_UNKNOWN_RUNNER`; logs are ignored `target/m1-*.log`.
No installed Excel binary, vendor artwork, downloaded macro or new Office oracle
observation is used. Fixtures and adapter code are independently authored.

Public behavior references:

- [Worksheet object](https://learn.microsoft.com/en-us/office/vba/api/excel.worksheet)
- [Range object](https://learn.microsoft.com/en-us/office/vba/api/excel.range%28object%29)
- [ByRef type identity and temporary expressions](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/byref-argument-type-mismatch)
- [Nothing](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/nothing-keyword)

## Remaining scope

Default members and true collections remain new M1.2/M1.3. Broader Excel classes,
range semantics, document modules/events, Range-aware UDF context, forms and
unchanged package/writeback acceptance remain their later milestones. Standard
controls are not enabled by type metadata; third-party ActiveX/OLE remains out of
scope. The [compatibility audit](vba-excel-compatibility.md) retains these boundaries.
