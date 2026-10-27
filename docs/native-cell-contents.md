# Native cell-content command

Updated 2026-10-27. `range.setContents` writes explicit scalar values or formulas
while preserving each destination's format. It is independent of VBA and uses the
same engine command lifecycle as UI, CLI, MCP and control calls.

The native API accepts a row-major rectangle without JSON dispatch:

```rust
use gridcraft_engine::{CellContent, Session, core::{RangeRef, Value}};

fn write(session: &mut Session, destination: RangeRef) -> gridcraft_engine::Result<()> {
    session.set_range_contents(0, destination, &[
        CellContent::Value(Value::text("=literal text")),
        CellContent::Formula("=1+2".into()),
    ])?;
    Ok(())
}
```

Here `destination` must contain exactly two cells. `Value` constants preserve
Empty, Number, Text, Bool and Error. Text is stored literally, including leading
equals signs, apostrophes, numeric/date-looking strings and empty strings.
`Formula` explicitly parses a formula, with or without its initial `=`, using
GridCraft's formula parser, function argument validation and normal recalculation.
Neither a formula nor a cell error becomes text when the format is Text (`@`).
Cell/row/column formats are retained; this API does not infer dates or change
number formats. Cell arrays are rejected: rectangles use the flat content list.

The equivalent command envelope is replayable:

```json
{
  "sheet": 0,
  "range": "A1:B1",
  "contents": [
    {"value": {"t": "Text", "v": "=literal text"}},
    {"formula": "=1+2"}
  ]
}
```

Call `Session::execute("range.setContents", envelope)` or the existing control/CLI
command route. `sheet` accepts an index/name and defaults to the active sheet.
The command requires one explicit rectangle, and rejects a conflicting
sheet-qualified range. Every content object has exactly one `value` or `formula`
field. Values use core's tagged scalar schema: `{"t":"Empty"}`,
`{"t":"Number","v":1.5}`, `{"t":"Text","v":"..."}`,
`{"t":"Bool","v":true}` or `{"t":"Error","v":"Div0"}` (a `CellError`
enum name). Unknown content/scalar fields or tags, nested arrays and mismatched
payloads fail.

Native and JSON paths use the same mutation/validation implementation. Native
calls bypass JSON decoding; a successful native call creates a serialized
envelope only for the journal/recorder. Recording replay therefore reaches the
same command, with the same types and formula distinctions.

The common executor provides enabled checks, panic recovery, history/groups,
recording, journaling and revision updates. The common edit/commit path provides
protection, legacy-array checks, table-header synchronization and recalculation.
A failed range write commits no part of that range and adds no journal/history
entry. Earlier successful macro writes retain their existing partial-effect and
spanning-undo-group policy.

Limits apply before mutation: exactly matching shape; 65,536 cells; 8 MiB content
payload charged as 64 bytes per cell plus text/formula UTF-8 bytes or 16 bytes
for other scalars; finite numbers only; formula bodies contain 1–8,192 Unicode
characters (excluding an optional leading `=`). Input-derived allocations/work are bounded. The wire decoder checks
tags/shape/payload before cloning text or decoding error enums and never
deserializes recursively shaped cell arrays.

`range.setValues` retains its existing typed-input and `parse` behavior. Use
`range.setContents` for explicit host storage semantics. Tests live in
`crates/engine/tests/contents.rs`; the real VBA adapter tests run natively and as
JavaScript-hosted WASM.
