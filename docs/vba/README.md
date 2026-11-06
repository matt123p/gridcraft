# VBA documentation

[Getting started](../vba.md) explains how to enable VBA, run a macro and add a
worksheet function.

| Reference | What you will find |
|---|---|
| [Supported functions and features](supported.md) | VBA intrinsics, worksheet-function names, workbook features and an object-member index |
| [Unsupported functions and features](unsupported.md) | Missing functionality, restricted options, host requirements and resource limits |
| [Integration architecture](integration.md) | Tiny-vba wiring, startup/host calls, design rationale, source storage, events and lifetimes |
| [Native worksheet controls](worksheet-controls.md) | Nine Form Control families, VBA members, trusted assignments, bindings, native persistence and limits |
| [Reference data](reference/README.md) | Machine-readable source mappings, signatures, implementation evidence and their interpretation |

The [design decisions](integration.md#design-decisions-and-trade-offs) explain why
commands, object identity, callbacks, source storage and editor execution use
these boundaries, and how the integration evolved.

Native UserForm runtime now includes selected dynamic Controls, multiple selection
and worksheet ControlSource/RowSource bindings. See the [support](supported.md),
[limits](unsupported.md) and [ownership](integration.md) references for their current
scope; M7A.4 remains in progress.

Selected native keyboard callbacks and the debugger-aware default-input barrier
are integrated from shared ADR 0070. Full CI, strict optional VBA/UI checks,
four form UI cases, web compilation and all 21 native/21 Node-WASM package
cases pass at shared revision `9fde04ce9ed8eab15b7f0609160617c774c2fc1b`.
The [unsupported reference](unsupported.md) retains the remaining M7A.4 gaps.
