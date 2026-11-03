# Excel VBA object-model audit and Luna implementation plan

Reviewed 2026-11-03. Status: implementation in progress. Scope is the
**third surface only: Excel's VBA object
model**, including its properties, methods, events, collections and enums.
This is a documentation-led audit, not an audit of a particular workbook.

## Scope and the evidence to use

Use Microsoft's [Excel object-model reference](https://learn.microsoft.com/en-us/office/vba/api/overview/excel/object-model)
as the audit catalogue, but use GridCraft's current capabilities as the hard
implementation boundary. An engine command, worksheet formula or visible UI
feature does not by itself establish a working VBA member; it establishes only
that a corresponding projection may be eligible. A catalogue entry does not
prove every optional argument or Excel behavior works.

This project integrates VBA with GridCraft; it does not add unrelated GridCraft
features. A member is implementation work only when the underlying behavior is
already available through a GridCraft command, model, calculator function, file
path, renderer or host service. VBA-specific binding, dispatch, type conversion,
object identity, diagnostics and tests are in scope. Adding a missing spreadsheet,
UI, file, provider, rendering or collaboration capability is not.

Every row must record one of these eligibility outcomes before packet priority:

- `ELIGIBLE_EXISTING`: project the existing GridCraft behavior through VBA;
- `OUT_OF_SCOPE_GRIDCRAFT`: retain it in compatibility output, but do not build
  the underlying feature in this workstream;
- `POLICY_BLOCKED` or `SERVICE_UNAVAILABLE`: retain the exact reason and fail
  explicitly rather than providing a compatibility stub.

High/Medium/Low orders eligible projections only. Deliver eligible High slices
first, then selected eligible Medium subsets. **Low is deferred and will not be
implemented in this workstream at this time**, even when its backing capability
exists. OM IDs are stable family identifiers, not implementation sequence numbers.
The [delivery order](#work-packets-and-delivery-order) below is authoritative.
A High Excel member whose behavior GridCraft lacks is still out of scope. If GridCraft later implements
that behavior through its own roadmap, eligibility can be reviewed separately.

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
Some pages require reconciliation and an eligibility decision before becoming
implementation requirements.
Member signatures, defaults, return types, read/write modes and platform remarks
must be expanded per packet; the snapshot is not a completed semantic oracle.

Implementation baseline is the current GridCraft branch and the immutable
tiny-vba revision pinned by its Cargo manifests and lockfile. Check the actual
pin and current adapter before each implementation slice; this plan does not
freeze a historical source commit. The Microsoft documentation archive and
inventory provenance remain independently pinned as recorded above.

## Initial audit baseline (historical)

The tables below describe the initial audit, before the OM02 slices recorded at
the end of this plan; they are not the current missing-work queue. In particular,
Union/Intersect, fills, Replace, SpecialCells, CurrentRegion, linked paste and
grouped selection have since landed. Use current descriptors, tests and the
implementation progress below when selecting work.

The shipping adapter is principally `crates/vba/src/objects.rs`, with argument
binding in that file, type/enum registration in `types.rs`, completion metadata
in `lib.rs::catalog`, and specialized range/value/format/clipboard helpers.
`worksheet.rs` supplies a separate constrained UDF execution context. The
engine implements many more features than this adapter exposes.

| Surface | Current baseline | Catalogue gaps (eligibility still required) |
|---|---|---|
| Application | Dynamic workbook/sheet/cell/selection roots; calculation, EnableEvents, CutCopyMode; explicit same-workbook Range Goto | Workbooks/Windows access, Union/Intersect, general Run/Evaluate/WorksheetFunction and eligible Goto variants; screen/alert/status settings, scheduling, UI and platform APIs remain out of scope where GridCraft lacks them |
| Workbooks/Workbook | Code-owner and active identity, selected properties and worksheet collections | Workbooks collection; Add/Open/Close/Save/SaveAs/SaveCopyAs, windows, names, metadata, links, protection and the remaining documented members |
| Worksheets/Sheets/Worksheet | Worksheet-only collections, Count/Item/enumeration, name, ranges, UsedRange, narrow Select/Activate/Paste | Add/Copy/Move/Delete and grouping; visibility/CodeName/Index; chartsheets; protection, data/print/drawing objects and remaining members |
| Range | Selected construction, default access, values/formulas, geometry, areas, offset/resize, End/Find, contents clearing, bounded structural operations, Select/Copy/Cut and a native-backed PasteSpecial subset | Remaining Destination variants, PasteSpecial source-theme/conditional-format modes, fills, Replace, SpecialCells, CurrentRegion, merges, formatting breadth, formulas/local/R1C1/dynamic arrays, filtering/sorting and remaining methods/properties |
| Clipboard | Single-area Copy/Cut, Range.Copy/Cut Destination, Worksheet.Paste with or without Destination, bounded current-workbook moves/pastes, formula/style preservation, Change callback, CutCopyMode and PasteSpecial type/operation/blank/transpose subset | Cross-workbook operations/consent, Worksheet.PasteSpecial format variants, linked paste, multi-area rules, CopyPicture/system formats |
| Formatting | NumberFormat; selected Font properties; Interior.Color | Borders, alignment, themes, patterns, remaining font/fill members, Styles, conditional formatting, dimensions and page-facing formatting |
| Events | Explicit Workbook_Open and bounded Worksheet_Change delivery | Remaining Application/Workbook/Worksheet events, cancellation ByRef, ordering and host UI integration |
| UDF host | Range-aware arguments, restricted reads, Caller/Volatile context | General WorksheetFunction/Evaluate/Run are not implied by this path; preserve read-only UDF restrictions |
| All other object families | Engine/UI/package functionality may exist | No general VBA projection for Sort/SortFields, Names, ListObjects, Windows, PageSetup, charts/drawings, pivots, connections and most other documented objects |

Inventory status meanings:

| Status | Snapshot pages | Meaning |
|---|---:|---|
| `implemented-subset` | 104 | Positive catalogue/dispatch or event evidence; expand option/semantic gaps before claiming completion |
| `context-only` | 2 | Caller/Volatile on the UDF path, not general macro support |
| `object-subset` | 9 | An object-page alias for a type represented in the adapter |
| `not-exposed` | 4,602 | No matching member in the reviewed adapter projection |
| `platform-service-needed` | 2 | Clipboard image and external-format paste dispatch explicitly refuse; without corresponding GridCraft services they are out of this integration's scope |
| `object-not-exposed` | 335 | No general projection of this documented object |
| `formula-bridge-not-exposed` | 390 | WorksheetFunction member pages; bridge work only in this plan |
| `partial-definition` | 10 | Existing enum type registration; reconcile actual values/aliases |
| `definition-gap` | 259 | Enum definitions absent from the projection or awaiting reconciliation |
| `reference-triage` | 283 | Standalone/legacy/reference page; reconcile receiver, alias and platform before counting a feature |

No row is marked fully Excel-compatible. The groups overlap conceptually across
objects and members; do not turn these counts into a compatibility percentage.
The JSON packet field is an initial routing aid. The packet contract below takes
precedence where a member has multiple concerns, and OM20 must resolve the tail.

## Ownership and implementation rules for Luna

- GridCraft owns Excel object definitions, lifetimes, collections, enums,
  workbook/document state, UI, undo/recalc, file persistence and services.
  Reuse native engine commands and models. Adapter plumbing may expose existing
  semantics, but this project must not add an engine/product capability merely
  to satisfy an Excel member. When no underlying capability exists, classify
  the row `OUT_OF_SCOPE_GRIDCRAFT`. Do not implement a second workbook or
  duplicate an engine algorithm inside the VBA adapter.
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
- Unsupported, unavailable, deprecated, out-of-scope and platform-restricted are distinct
  outcomes. Record every documented row even where implementation needs a new
  service or policy contract. A missing GridCraft service is not VBA backlog.
  Do not hide difficult APIs by omitting them from the catalogue or declaring
  all of an object supported after one method works.
- Preserve current roadmap IDs and completed intrinsic work. OM IDs below are
  a new cross-repository object-model workstream, not replacements for M1-M8.

## Architecture review and low-friction delivery (2026-11-03)

This review inspected the Rust implementation, not just the reference catalogue.
Keep the existing crate boundary. No new crate, RPC service, plugin framework or
second workbook abstraction is needed. The changes below are implementation
instructions for subsequent slices, not claims that those changes have landed.

### Verified ownership and concrete review findings

| Concern | Existing Rust evidence | Delivery decision |
|---|---|---|
| Dependency boundary | `crates/vba/Cargo.toml`, `crates/engine/Cargo.toml`, `xtask/src/layers.rs`: adapter at L6 above engine; engine has no VBA dependency | Keep Excel names, errors, bindings and handles in `gridcraft-vba`; retain optional shell features. |
| Dispatch and metadata | `objects.rs::invoke_object` is a large receiver/member ladder; `lib.rs::member_descriptors` supplies signatures, but handlers separately pass argument-name arrays to `objects::arguments` | Extract touched families into private modules incrementally. Share reviewed argument specifications; do not require a wholesale dispatcher rewrite. |
| Identity and targeting | `types.rs::Objects`, `Bridge::sync_sheets`, `WorkbookHost::invoke`, `Session::with_target` already implement shared handles and document targeting | Reuse these mechanisms for new collections; never create another registry or capture document borrows in objects. |
| Mutations | `values.rs` uses native typed content operations; `clipboard.rs` and `search.rs` use engine commands, including Replace preview | Keep exact VBA values typed. Reuse command execution and its undo/recalc/protection path; do not write model cells directly. |
| Events | Review reproduced wrong-workbook ownership, duplicate/missing events and post-write capacity failures; now `prepare_changes` runs in the receiver context, materializes handles/arguments and reserves the complete batch | Remediated with native and WASM regression coverage recorded below. Imported handlers remain bound to their own workbook, including when another workbook uses identical code names. |
| Services | `WorkbookHost::resume` currently refuses and `cancel` is empty | Synchronous projections can land now. File/dialog operations needing a wait require a separate bounded continuation slice; engine file support alone does not establish that host contract. |
| Audit | `xtask/src/vba_audit.rs` already verifies the pinned archive and generates contracts, enum Rust and the audit report | Extend this exporter; do not build a competing generator. Current result states are derived from baseline status/metadata, not executed compatibility evidence. |

The source comment defending a single dispatch ladder describes the current
implementation, not a requirement to keep growing it. Leave `objects.rs` with
receiver/default-member routing, shared binding and collection mechanics; move
Application, lifecycle, formatting and data-tool handlers into private family
modules as those families are touched. Keep one `Bridge` and object store. A
small direct Rust match or shared helper is preferable to a generic kernel DSL
until several actual members demonstrate the same behavior.

### Keep the engine boundary useful, not obstructive

An adapter may need a narrow **host-neutral exposure of existing behavior**:
explicit document/sheet/range targeting, a read-only preview, a typed entry point
sharing the command executor, or an effect footprint. That is allowed plumbing,
as demonstrated by the existing Replace preview and native content API. Keep it
small, reusable by CLI/UI/MCP, and tested at the engine boundary. It must not add
Excel member names, VBA errors, foreign handles, or another mutation path below
L6. A new spreadsheet algorithm, dialog, provider or persistence capability
still belongs to GridCraft's separate roadmap.

Prefer explicit targets when the engine supports them (`data.sort` already
accepts sheet/range). Do not introduce temporary selection/activation merely to
reuse a UI-oriented default. Where unavoidable in an existing command path,
define and test the visible state on both success and failure. An undo group is
not an atomic transaction for a sequence of commands. Validate options, target,
protection and all event/handle capacity before the first write; do not roll back
successful earlier macro statements on a later error.

Adapter-only transient state is legitimate for Excel interaction protocols,
such as pending SortFields before Apply, Find options and event policy. Key it
by existing document/sheet identity, bound it and clear/invalidate it on the
appropriate lifecycle transition. It must not duplicate persisted workbook
state or add a new GridCraft feature.

### One source for each kind of information

Keep source-page provenance in the inventory and documented contracts in the
exported contracts file. Add reviewed implementation decisions to the inventory,
keyed by stable row ID: eligibility, engine symbol/command evidence, supported
options, family, priority and test IDs. Distinguish an unreviewed row from an
out-of-scope row; do not classify all absent projections as missing engine
capabilities. Review the selected slice first; full-catalogue classification is
an acceptance task, not a prerequisite to the next useful PR.

Keep supported runtime descriptors in Rust. Incrementally add structured
argument/access information used by binding and catalogue generation for touched
families. Do not parse signature strings at runtime. Pinned documentation has
unknown defaults/modes; it is input for review, not executable semantics.
Generate only reviewed boilerplate and metadata, with handwritten semantic
handlers and positive behavior tests. Registration or a generic implementation
symbol does not prove that a handler supports all advertised options.

### Tooling that pays for itself

Use the existing Rust exporter for durable checks and generation:

```sh
cargo +1.99.1 run -p xtask --features vba-audit -- vba-audit
# Only when changing inventory/contracts or generated metadata:
cargo +1.99.1 run -p xtask --features vba-audit -- vba-audit --write
cargo +1.99.1 run -p xtask --features vba-audit -- vba-audit
cargo +1.99.1 test -p gridcraft-vba --test collections --test workbooks
cargo +1.99.1 xtask layers
```

The audit needs `target/vba-docs-audit/reference.zip` with the documented hash;
ordinary builds/tests must not fetch it. Replace the example test targets with
the affected family suites. Use the full commit gates below before committing.

Next extend the same exporter to validate eligibility/evidence fields and emit
an eligible queue by family. Then add binding consistency checks and generated
negative cases for the first repeated family. Do not generate semantic expected
outputs from the same descriptors being tested. Keep generated Rust checked in;
no build-time generator or Python runtime dependency. Defer `vba-usage`, proc
macros, command-schema extraction and benchmarking infrastructure until a
measured recurring task justifies them. Command parameter documentation is not
an executable schema and cannot supply Excel defaults automatically.

Python's standard library is useful now for read-only JSON joins, duplicate and
page-count checks, packet counts and Markdown link checks. Use it for analysis
or mechanical migrations with reviewed diffs; promote recurring authoritative
checks into the existing Rust xtask. Never execute downloaded macros during
analysis. Do not hand-edit generated contracts, enums or the audit report.

### Smallest useful implementation sequence

1. Select one eligible synchronous OM03/OM04 slice, recording its engine backing
   and reachable root/collection path. Existing calculation is a baseline to
   verify, not work to recreate. `Run` needs generic invocation/reentrancy review;
   it is not automatically a simple engine-command wrapper.
2. Extract only the family being changed and share its repeated binding data.
   Land behavior plus metadata/tests together; avoid a foundation-only rewrite.
3. Add reachable read-only collections/properties and existing sheet operations;
   handle blocking Open/Save/Close and prompt continuations separately. Retain the
   target/event reservation regressions when broadening cross-document writes.
4. Take the smallest eligible formatting, Sort/SortFields or formula/name slice
   whose actual prerequisites exist. WorksheetFunction should use one proxy into
   the existing calculator, with explicit per-function eligibility and VBA
   argument/result/error conversion, not hundreds of wrappers.
5. Complete the eligible High table graph (OM18) before general Medium work.
   Deliver OM11 event behavior with each High slice; enforce protection on every
   mutation and project ordinary Protect/Unprotect under the High member rule.
   Then take the selected Medium subsets in the delivery order below. Low slices
   are deferred. Prerequisites are per slice: an existing single-workbook
   formatting property need not wait for all lifecycle, view and event APIs.

For each slice test exact results/errors, named/omitted arguments, receiver
identity and the relevant effects. Add undo/protection/events for writes,
persistence for persisted state, and pause/resume for waiting operations. Reuse
shared regression suites instead of repeating the entire matrix for each
property. UI/render inspection is required when the slice changes visible UI
or rendering; it is not a prerequisite for a metadata-only change. No reliable
calendar estimate follows from documentation page counts.

## Work packets and delivery order

**Do not implement packets in numeric order.** OM00-OM02 are accepted foundations;
the event remediation is also complete. The remaining delivery order follows
[the accepted priority report](vba-object-model-v1-priority-final.md), subject to
each member's existing-capability gate. Keep OM IDs and inventory routing stable:
renumbering them would break provenance and handoffs without simplifying delivery.
Detailed packet contracts below remain in numeric order for reference only.

| Stage | Delivery order | Acceptance boundary |
|---|---|---|
| Accepted foundation | OM00 → OM01 → OM02, plus completed event remediation | Retain the accepted bounded profiles and their regressions; do not restart them. |
| High: reachable roots | OM03 + OM04 synchronous roots, collections, identity and existing sheet operations | Establish the object paths needed by later slices. Review ordinary protection members from OM09 here where they support these operations. |
| High: formula and function access | OM07 | Reachable WorksheetFunction proxy, then eligible formula/name/Evaluate slices. Do not wait for blocking file lifecycle operations to finish. |
| High: everyday workbook operations | OM05 → OM06 → OM18 | Common formatting, then sort/filter/validation/data tools, then structured tables. OM18 comes before general Medium work. |
| High: remaining invocation/lifecycle work | Remaining eligible OM03/OM04 slices as their contracts are ready; complete before High sign-off | Run/reentrancy and Open/Save/Close waits are separate slices. Implement earlier when their actual dependencies are ready; they do not block unrelated synchronous High work. |
| Alongside every implemented slice | OM11 events; OM09 protection enforcement; OM00 metadata and OM20 reconciliation | Event ordering/cancellation and existing protection gates land with the operation, not in a later catch-up packet. Events follow their owning feature's tier. |
| Medium: first subsets | OM12 → OM13 | Reachable basic charts/drawings first, then existing local pivot behavior. No provider/OLAP expansion. |
| Medium: remaining selected subsets | OM08 → remaining OM09 → OM10 → OM19 → OM15 | Views, additional protection/permission projections, print/export, local annotations and interaction/controls, taking only slices whose dependencies and services exist. |
| Deferred Low | OM14, OM16, OM17, and Low members routed to any other packet | No runtime implementation now. Retain provenance, classification, existing package fidelity and explicit unsupported results. |
| Profile sign-off | OM20 | Close after High and the selected Medium subsets; Low implementation is not a prerequisite. Publish supported, partial, out-of-scope and deferred rows separately. |

Arrows give the default queue, not whole-packet barriers. A later slice may move
forward when its actual receiver/engine dependencies exist. A small Medium-owned
helper required by a High operation (for example a prompt continuation) may land
with that High slice; this does not authorize the rest of the Medium family.
Do not pull Low behavior into a High slice through a dependency. Refuse the Low
variant and retain the portable High path. High acceptance requires every
eligible required High obligation to be implemented/tested; unresolved High
requirements block High sign-off. Record them rather than silently downgrading
them to unblock Medium work.

Ordinary sheet/workbook Protect/Unprotect and protection status are High under
the member-level rules in the final report, even though OM09's family default is
Medium. Existing protection enforcement is required from the first mutation.
Likewise, OM11's family default is High, but events for deferred Low APIs remain
deferred, and Medium-feature events accompany their selected Medium slice.

### Packet index by delivery tier

Counts retain the original source-page routing. A family tier is a planning
default; classify mixed-owner members by behavior. Prerequisites below mean the
specific needed receiver/command contract, not completion of every listed packet.

| Packet | Reconciled pages | Tier / disposition | Deliverable and actual dependencies |
|---|---:|---|---|
| OM00 | 272 | Accepted / cross-cutting | Pinned inventory, signatures/enums and metadata; maintain with each slice. |
| OM01 | Cross-cutting | Accepted foundation | Execution policy, registry and diagnostics over OM00. |
| OM02 | 68 | Accepted High profile | Selection, clipboard/editing/ranges; extend only for selected High dependencies. |
| OM03 | 206 | High | Existing Application roots/calculation/invocation; synchronous roots first, Run only with its reviewed invocation contract. |
| OM04 | 278 | High | Existing workbook/sheet collections and lifecycle; needs relevant OM03 roots and explicit consent/identity. Waiting services are separate slices. |
| OM07 | 515 | High | Names/formula/Evaluate/function bridge; needs only the relevant OM03/OM04 roots and calculator context. |
| OM05 | 541 | High | Existing formatting/styles/conditional-format subsets; existing Range/Worksheet receivers suffice for simple properties. |
| OM06 | 180 | High | Existing sort/filter/validation/data tools; depends on OM02 range targeting and only the OM05 properties actually consumed. |
| OM18 | 134 | High | Existing structured tables; needs relevant OM04 receivers, OM05 styles, OM06 operations and OM07 structured references per slice. |
| OM11 | 123 | High by owning feature | Existing event hooks, cancellation/order and lifecycle semantics; implement alongside the corresponding operation. |
| OM12 | 1,220 | Medium | Existing basic charts/drawings; needs OM04 object ownership and relevant OM05 formatting; OM08 only for genuinely window-dependent members. |
| OM13 | 838 | Medium | Existing local pivots; relevant OM04/OM06 and table/formula contracts where used; provider/model behavior excluded. |
| OM08 | 163 | Medium | Existing windows/panes/views; relevant OM03/OM04 document/window identity. |
| OM09 | 96 | Medium default; ordinary protection members High | Existing protection/permission projections; relevant OM04 receivers. Mutation protection enforcement applies throughout. |
| OM10 | 133 | Medium | Existing page setup/print/export; relevant OM04/OM05 state, OM08 only for view-dependent calls, actual export/print service. |
| OM19 | 101 | Medium | Existing local notes/links/metadata; relevant OM04/OM05 receivers and event hooks if the operation needs them. |
| OM15 | 115 | Medium | Existing interactions/controls; actual host-service continuation and relevant OM03/OM08/OM11 contracts only. |
| OM14 | 305 | Low — deferred | Connections/query/import/XML/refresh automation: classify/report only; no projection work now, including otherwise eligible local variants. |
| OM16 | 138 | Low — deferred | Remote/collaboration/publishing: classify/report only; no projection/provider work now. |
| OM17 | 570 | Low — deferred | Legacy/Graph/platform/VBE interoperability: classify/report only; no projection/platform work now. |
| OM20 | Cross-cutting | Continuous / sign-off | Reconcile source rows and reachability against the selected profile; close without implementing deferred Low rows. |

`DEFERRED_LOW` is a planning disposition, separate from eligibility and runtime
result state. An eligible Low row is deferred, not `OUT_OF_SCOPE_GRIDCRAFT`; a row
lacking backing behavior remains out of scope and may also be Low. Runtime
support already present is retained; this decision creates no new Low runtime
work. Low triage records the family/tier, eligibility or known restriction and
source provenance; it does not require deep parameter review, new receiver
objects, enum/binding generation, per-member stubs or a new Low test programme.
Use existing unsupported dispatch/diagnostics rather than building deferred graphs.
The per-row priority/disposition schema migration remains planned, not
implemented by this documentation change. Do not reinterpret generated MISSING
counts as an instruction to implement all rows, or hand-edit generated reports.

Page routing is not a project estimate. Record genuine family reassignments in
the inventory, keeping the pinned source set; delivery ordering alone requires
no inventory reassignment or generated-file regeneration.

### OM00 — inventory and contracts first

Implement a Rust maintenance exporter/checker, preferably `cargo run -p xtask --features vba-audit -- vba-audit`,
that reads a locally cached pinned documentation snapshot without fetching at
build/test time. Reproduce the current page set and provenance. Preserve every
input page; recognize aliases, legacy Excel Graph topics and enum pages.
Maintain stable receiver-qualified IDs and retain source page IDs separately
when several pages map to the same member.

Expand rows with member kind, signature, argument modes/defaults, return type,
read/write support, event parameters/cancellation, enum values/aliases, platform
notes, dependencies, profile, implementation symbol, test IDs and result state.
Use states such as MISSING, PARTIAL, IMPLEMENTED_TESTED, OUT_OF_SCOPE_GRIDCRAFT,
PLATFORM_SERVICE_NEEDED, LEGACY_REVIEW and REFERENCE_ALIAS. A service-needed row
is implementation work only when the corresponding GridCraft service exists;
otherwise it is out of this integration's scope. Fully implemented states require
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
overall compatibility percentage. Every row must also have an explicit
eligibility outcome; out-of-scope rows are reported, not treated as unfinished.

### OM01 — cross-cutting execution and diagnostics

The shared IDE now accepts bounded execution limits used at normal launch and
idle document callback creation (`tiny-vba/crates/vba-ide/src/lib.rs`). GridCraft
sets a macro profile with a 12,000 host-call cap and 1 GiB payload cap; the VM
default remains 1,024 host calls for other consumers. Do not remove the VM limit
or increase every consumer's global defaults.
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

Add eligible missing global roots/collections and preserve unqualified root
behavior. Project calculation state, `Run`, environment/path values and settings
only where GridCraft already owns equivalent state or commands. `ScreenUpdating`,
`DisplayAlerts`, status/cursor settings, scheduling, custom lists and similar
members are `OUT_OF_SCOPE_GRIDCRAFT` when no matching product behavior exists;
do not create host state, queues or UI behavior for them in this packet.

Treat Evaluate and formula bridges as OM07. Route remote/OS/legacy APIs to
OM16/OM17 classification rather than faking them.
Acceptance must exercise separate documents, live state, callback ownership,
teardown and headless/browser availability. [Application reference](https://learn.microsoft.com/en-us/office/vba/api/excel.application(object)).

### OM04 — workbook and worksheet lifecycle

Project Workbooks collection/default/enumeration contracts and the workbook and
sheet lifecycle operations already provided by GridCraft. Include eligible
Add/Open/Activate/Close/Save/SaveAs/SaveCopyAs and sheet Add/Copy/Move/Delete/
visibility/index/CodeName/chart-sheet behavior. Include
optional Before/After, omitted destination/new workbook, count/template, naming
collisions, workbook structure protection and stale references only where the
underlying operation exists. Missing Excel lifecycle variants are out of scope.

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

Project eligible existing Font, Interior, Borders/Border, Styles/Style and range alignment,
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
of locale/Pinyin/Stroke order. Expose only supported collation and classify the
rest `OUT_OF_SCOPE_GRIDCRAFT`; do not accept SortMethod and ignore it.

Then project existing Range/Worksheet AutoFilter, Filters/Filter and criteria operators,
ShowAllData, AdvancedFilter, validation objects, RemoveDuplicates, TextToColumns,
Subtotal, outline/grouping, scenarios/consolidation/goal-seeking and the remaining
eligible data-tool members routed to this packet. Reuse native sort/filter/validation/
outline commands. Test blanks/errors/mixed types/headers/stability, descending
order, multiple keys, metadata/formula movement, protected/merged cells, saved
state, event footprints, cancellation and zero writes on invalid configuration.
[Sort](https://learn.microsoft.com/en-us/office/vba/api/excel.sort),
[SortFields.Add](https://learn.microsoft.com/en-us/office/vba/api/excel.sortfields.add).

### OM07 — names and formula-facing object-model APIs

Project existing workbook/sheet Names/Name behavior: scope, Add/Delete, RefersTo variants,
RefersToRange, visibility and identity. Complete Formula/FormulaR1C1/Formula2,
array/spill and local forms without confusing displayed text and formula source.
Add Range/Worksheet/Application Evaluate and Application.WorksheetFunction
through a typed context-sensitive bridge to existing engine functions. Preserve
Range arguments, arrays, errors and Excel return/error differences between
Evaluate, WorksheetFunction calls and worksheet formulas. Unsupported calculator
functions are `OUT_OF_SCOPE_GRIDCRAFT`; this packet does not implement their
mathematics. Inventory the 390 WorksheetFunction member pages as bridge
candidates, then apply function-registry eligibility.
Test reference ownership, dynamic arrays/spills, default members, errors versus
raised runtime errors, UDF read-only policy, locale and recursion/payload budgets.

### OM08 — windows, panes and navigation

Medium: follow the delivery table after High completion, except a narrowly
required helper for a High operation.

Project eligible Windows/Window/Panes/Pane and custom/sheet views onto real per-window
state: ActiveWindow, activation, selection, scrolling, freeze/split panes, Zoom,
gridlines/headings, display modes, bounds and window lifecycle. Worksheet-owned
state and window-owned state must remain distinct. Goto scrolling depends on
this packet. Test two windows on one workbook, two workbooks, focus changes,
saved views, closed handles, narrow/native/browser layouts and headless results.
[Window reference](https://learn.microsoft.com/en-us/office/vba/api/excel.window).

### OM09 — protection and permissions

Ordinary protection/status/Protect/Unprotect members are High under the
member-level rules; additional eligible permission/protected-view/encryption
projections are Medium. Existing protection checks apply to every mutation
throughout the workstream. Do not defer those checks until this packet.

Project existing workbook/sheet/range protection and AllowEditRanges, protection status,
password/encryption options, protected-view and document permissions through
existing protection models and services. Unsupported options are out of scope. Preserve option
semantics such as UI-only protection and allowed edits; do not treat a password
string as proof that encryption exists. Validate protection for every newly
exposed mutation. Test locked/unlocked cells, structure protection, failed
authentication, save/reopen, cancellation and retained trust boundaries.

### OM10 — printing and export

Project existing PageSetup, print areas/titles, margins, paper/orientation/scaling,
headers/footers, page breaks and Workbook/Worksheet/Range PrintOut/PrintPreview/
ExportAsFixedFormat. Delegate to existing print/PDF services; missing engine
options are `OUT_OF_SCOPE_GRIDCRAFT` rather than VBA implementation tasks.
Printer selection,
preview and native printing are service-backed; pure file/PDF export can have a
headless path. Test page geometry and selected-sheet/range scope by rendering
output and inspecting it, plus cancellation, unavailable printer and file errors.
[PageSetup](https://learn.microsoft.com/en-us/office/vba/api/excel.pagesetup).

### OM11 — event completeness

Deliver events with the owning operation: High lifecycle/edit/table events in
the High queue, selected Medium events with their feature, Low events deferred.
This packet is not a demand to implement all 123 event pages before High sign-off.

Inventory every documented event and ByRef/Cancel parameter. Extend delivery
only for lifecycle and mutation signals GridCraft already exposes. Events that
would require new product hooks are `OUT_OF_SCOPE_GRIDCRAFT`. Host commands, UI
edits and VBA writes need one
defined ordering and suppression model. EnableEvents must apply at the correct
scope; UDF execution must not accidentally enable mutating events.
Use cancellable precommit events where documented; do not deliver BeforeSave or
BeforeClose after irreversible effects. Test reentrancy/recursion limits, FIFO,
ByRef cancellation, cross-document callbacks, error policy, modal waits, pause/
resume, source edits, Reset/close and event subscriptions across saved projects.
Do not retain document borrows or mutable locks across callbacks.

### OM12 — charts and drawing objects

Project eligible existing ChartObjects/ChartObject, Charts/Chart,
ChartArea/PlotArea, Series/
SeriesCollection, axes/titles/legends/labels, data binding and chart settings;
then Shapes/Shape/ShapeRange, pictures, text frames, connectors, groups, geometry,
z-order, formatting, placement, OnAction, export and copy/paste. Include reachable
child/collection objects backed by existing GridCraft state. Reuse chart/drawing
models, commands and renderer. Engine feature gaps are out of scope for this
packet and remain on the normal GridCraft roadmap. Test object identity, data edits,
anchor transforms, grouping/ungrouping, protected sheets, undo, source deletion,
OOXML round trips and rendered native/browser/export output.
[Shapes](https://learn.microsoft.com/en-us/office/vba/api/excel.shapes).

### OM13 — pivots, slicers and data models

Project GridCraft's existing PivotCaches/PivotCache, PivotTables/PivotTable,
fields/items/layout, filters/calculated fields, refresh/drilldown and eligible
charts. Cube/model/relationship/measure, slicer or timeline members without
existing GridCraft behavior are out of scope. Use existing pivot models;
do not create superficial objects that cannot calculate or persist. Separate
local pivot operations from external OLAP/data-model services. Test cache reuse,
refresh, missing data, field identity, filter/layout results, saved files, events,
undo and service cancellation. Unsupported provider functionality remains
visible as `OUT_OF_SCOPE_GRIDCRAFT`, not an empty successful refresh or VBA backlog.

### OM14 — connections, import and refresh

Low — deferred; classification/reporting only at this time. Do not implement
members even if a local backing behavior exists. Native file/import features
and existing package preservation continue through their established owners.

Classify Connections/WorkbookConnection, QueryTables/QueryTable, queries,
parameters, text/ODBC/OLEDB/XML/data-feed connection objects, mappings and refresh
lifecycle. Retain local candidates for future reprioritization; no OM14 runtime
projection is scheduled now. Database
drivers, network/authentication, refresh providers and credential ownership are
not added by this project and are `OUT_OF_SCOPE_GRIDCRAFT`. Existing package-part
preservation remains a file-fidelity behavior, not an executable VBA provider.

### OM15 — interaction and built-in controls

Project documented Excel dialogs, Application.InputBox type/Cancel semantics,
file pickers, status/interaction surfaces, dialog-sheet/form-control objects and
their action/event wiring only where an equivalent GridCraft interaction already
exists. Scope built-in Office controls according to existing
tiny-vba form contracts; third-party ActiveX/OLE controls remain outside the
adopted policy and must be listed explicitly. No fake dialogs or success without
user-visible effects, and no new GridCraft dialogs or controls are created for
VBA parity. UI interaction needs suspend/resume with released borrows,
bounded callbacks and explicit headless/browser behavior. Test modal ownership,
Cancel and typed results, selection/focus, project teardown and accessibility.

### OM16 — remote, collaboration and publishing APIs

Low — deferred; classification/reporting only at this time.

Keep every sharing/review/mail/server/SharePoint/publishing/cloud-related row
visible. No new projection is scheduled, including for an existing GridCraft
feature/service.
There are currently no general host adapters for these APIs, so provider-backed
rows are `OUT_OF_SCOPE_GRIDCRAFT`. Do not reinterpret SendMail as a
no-op, automatically send messages, or silently replace a Microsoft service with
a different one. This plan does not authorize new providers, external sends or
remote state changes.

### OM17 — legacy and platform interoperability

Low — deferred; classification/reporting only at this time. Portable equivalents
remain deferred even if existing GridCraft behavior could back them.

Reconcile Excel Graph legacy topics, XLM/Excel4 methods, DDE, COM/XLL/add-ins,
OLE, OS-specific operations, VBProject/VBE and similar APIs. Classify aliases,
deprecated-but-documented contracts and true platform dependencies. Preserve
all rows even when a native portable implementation is inappropriate.
The existing Rust-only/no-ambient-COM/no-VM-IO and third-party-control boundaries
still apply. Retain eligibility evidence separately from the deferred priority;
use `OUT_OF_SCOPE_GRIDCRAFT`, `SERVICE_UNAVAILABLE` or `POLICY_BLOCKED` only where
those reasons actually apply. Do not change eligibility merely to express deferral.
Any policy expansion needs its own reviewed design, not an implicit exception in
an implementation PR. A complete portable profile and complete Windows Excel
interoperability are different acceptance claims; publish both denominators.

### OM18 — structured tables

High: deliver the selected table graph after its actual formatting/data/formula
prerequisites, ahead of general Medium work; its number is not its queue position.

Project existing ListObjects/ListObject, ListRows/ListRow, ListColumns/ListColumn,
table ranges/totals/header rows, resize/convert/add/delete, styles, filters and
structured references through actual table models. Include TableObject and
connection-backed variants with OM14; they are Low/deferred, and independently
out of scope without an existing provider. Test empty-table ranges, headers and
totals, calculated columns, row insertion/deletion, reference maintenance,
filter/sort interaction, protection, formula spill interaction and saved files.
[ListObject](https://learn.microsoft.com/en-us/office/vba/api/excel.listobject).

### OM19 — annotations and metadata

Project existing Comments/Comment, threaded comments/persons/replies where supported,
Hyperlinks/Hyperlink, custom and built-in document properties, custom worksheet
properties and annotation formatting/visibility. Distinguish local notes from
collaborative threaded state and external follow-link actions. Remote identity/
service-backed Low members remain deferred even if a provider becomes available. Use current
metadata/comment/link commands and package models. Test copy/move/delete,
structural rebase, author/thread identities, formatting, undo, file round trips
and unavailable remote/person services. Missing remote identity/thread services
are out of scope rather than new provider work.

### OM20 — finish the tail and acceptance

Every current page has a family packet; there are zero unrouted OM20 pages.
Reconcile every `reference-triage` row within its assigned family, including
names with unusual receiver spelling and legacy standalone pages. Ensure smart
tag actions/proofing (OM15), phonetics/characters (OM05), add-ins (OM17), cell
watches (OM08), file converters/recovery (OM04) and the smaller object families
are projected when eligible and selected for High/Medium delivery, or explicitly
classified as deferred Low/out of scope/unavailable rather than lost behind the
headline objects. Low and unselected Medium rows are not runtime implementation
requirements for this sign-off. New source pages must be routed to a precise
packet or a new explicit family packet. Do not
introduce an unexamined miscellaneous bucket at sign-off. Resolve duplicates/
aliases without deleting source-page provenance.

Close OM20 after High and the selected Medium subsets. Retain all Low rows in
the report as deferred; their implementation is not required for acceptance.
Do not label all of Medium complete when only selected subsets have landed.

Run a final object graph reachability audit: a member cannot be complete when
its receiver cannot be obtained, its return type has no usable projection, or
its enum/default/child collection is missing. Check parameter combinations,
events, lifecycle, UI, services and persistence independently. Generate the
supported/partial/missing/platform-restricted reports from the same inventory.

## Per-PR handoff and verification

Each implementation PR must demonstrate that the underlying GridCraft capability
already exists, then contain the specific inventory row IDs and Microsoft
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

## Recommended next Luna assignment

Take one synchronous High OM03/OM04 roots/collection slice with existing engine
backing, supported paths and identity tests. Verify the existing calculation
baseline rather than recreating it. Then take the reachable OM07 WorksheetFunction
proxy or formula/name slice; continue OM05 → OM06 → OM18 by actual prerequisites.
Keep OM11 events and OM09 protection obligations in each relevant slice. Return
to remaining eligible High OM03/OM04 lifecycle/Run work as its invocation/wait
contracts are ready, and finish it before High sign-off.

Do not start general Medium work while required High behavior remains incomplete.
The first Medium delivery is a basic OM12 chart/drawing slice, then OM13 local
pivots, followed by selected OM08/OM09/OM10/OM19/OM15 subsets. OM14/OM16/OM17 and
all other Low members receive classification/diagnostics only; no Low runtime
implementation is scheduled. Use the delivery table rather than numeric OM order.

## Audit validation performed in this planning task

- Downloaded Microsoft's nontruncated public Git tree and pinned archive; SHA256
  `285218108a034fc33acc8296d714395fd182397ada27c848b93c3d5423efd75e`.
- Reconciled every one of the 5,996 Excel API source paths, classified headings
  and matched the current catalogue/selected generic dispatch/context paths.
- Inspected GridCraft object registry/dispatch, clipboard/formatting helpers,
  native edit/sort/file commands and tiny-vba IDE/VM budget construction.
- Research scope excludes workbook specimens, intrinsic implementations and
  worksheet-function mathematics. No application code was changed during the
  initial audit; the subsequent implementation slices and their tests are listed
  in the progress section below.

The inventory is an auditable breadth baseline, with explicit uncertainty about
unreviewed parameter/platform semantics. It is not a claim that all 5,996 pages
have independently observed Excel behavior or completed implementation designs.

## Implementation progress

- **OM00a-1 — pinned source-page integrity and deterministic export:** implemented
  as `cargo run -p xtask --features vba-audit -- vba-audit [--write]`. It verifies
  the cached archive SHA256 and documentation commit, rejects duplicate paths,
  checks exact equality of all 5,996 Excel API pages, and deterministically
  generates `vba-object-model-contracts.json`, typed enum source and an unresolved
  row report. Current output has 272 enum types and 2,260 typed runtime constants.
  The source contract and generated runtime tables retain both same-value
  `Constants.xlManual` and `XlSortOrder.xlManual` identities. A generated ordered
  alias library supplies both module-qualified forms while deterministic
  unqualified lookup keeps the primary value; unit tests assert the numeric value
  for all three spellings. The
  schema 2 exporter records 4,296 argument occurrences: 2,879 required/optional
  flags and 2,880 types are extracted from parameter tables, with 172 explicit
  typed defaults extracted from descriptions. Argument mode is marked
  `NOT_SPECIFIED_IN_SOURCE` where Microsoft’s reference syntax provides no
  ByVal/ByRef mode; undocumented requiredness, types and defaults retain source
  provenance instead of being guessed. The
  audit passes and currently reports 5,321 MISSING, 390 PARTIAL, 283
  LEGACY_REVIEW and 2 PLATFORM_SERVICE_NEEDED implementation rows. Property
  access and platform notes are source-reviewed, legacy ownerless pages are
  explicitly routed to OM17, and the report deliberately publishes no aggregate
  compatibility percentage. OM00's inventory/contract acceptance is complete;
  those result states still require eligibility classification, and only eligible
  rows become work for their owning implementation packets.
- **OM01a — configurable bounded execution limits:** published in tiny-vba and
  pinned by GridCraft. GridCraft configures 25,000,000 instructions, 12,000 host
  calls and a 1 GiB payload ceiling for workbook macro sessions; limits carry
  across pause/resume and idle document-event sessions. The shared IDE embedding
  regression passed. Catalogue, completion/object-browser metadata and the audit
  now derive from `MemberDescriptor`; descriptor reconciliation rejects duplicate
  or unassigned canonical members. Unsupported dispatch fallbacks name receiver
  and member, while option/service/budget errors retain distinct VBA numbers and
  the VM supplies source locations. The nonexecuting JSON compatibility report
  classifies statically resolved calls as SUPPORTED, SERVICE_REQUIRED or MISSING
  and dynamic receivers as UNKNOWN. OM01 acceptance is complete.
- **OM02a-1 — clipboard and PasteSpecial subset:** implemented for one-area
  source and destination ranges in the current workbook, including cross-sheet copies.
  `Range.Cut` with or without Destination and `Application.CutCopyMode` stage same-workbook bounded moves;
  successful cuts queue source/destination Change targets and overlap refusal
  preserves source cells. Formula rebasing, styles, selection, cell bounds and
  callbacks use the native engine path. `Worksheet.Paste Destination` uses the
  same target behavior. `Range.PasteSpecial` supports engine-backed paste types,
  arithmetic, SkipBlanks and Transpose; source-theme and conditional-format merge
  modes refuse. Cross-workbook consent remains open. The 15-test clipboard suite
  passes, including callback ordering, live same-workbook `Worksheet.Paste
  Link:=True` formulas and the recorded 20-slot copy loop with one undo group.
  `Worksheet.PasteSpecial` and `Range.CopyPicture` validate their options and
  explicitly report unavailable host services. Because GridCraft does not
  currently supply those services, they are out of this integration's scope.
- **OM02a-2 — Range clearing subset:** `Range.Clear`, `ClearContents` and
  `ClearFormats` now route through native engine clear commands. Operations are
  capped at 65,536 selected cells; content-clearing calls preflight and queue the
  bounded Worksheet_Change target, while format-only clearing does not queue a
  value-change event. The range suite passes, including distinct content,
  format and full-clear behavior.

- **OM02a-3 — range activation and CurrentRegion:** `Range.Activate` now activates
  its owning worksheet, retains the selected range and sets its first cell as
  active. `Range.CurrentRegion` delegates to the engine's native region helper
  for a single-area range and refuses results above 65,536 cells. Both behaviors
  have focused VBA integration coverage.
- **OM02a-4 — range-scoped Replace:** the engine `edit.replace` command now accepts
  an optional sheet/range scope and a read-only preview that reports matching
  cells. `Range.Replace` validates its full documented argument list, limits the
  receiver to 65,536 cells, commits through the native replacement logic, returns
  a Boolean, and queues a single Worksheet_Change target containing only changed
  cells. Contiguous horizontal and vertical changed-cell runs are coalesced; more
  than 64 disjoint event areas, MatchByte, format search, and FormulaVersion
  service requests refuse before mutation. Tests cover range isolation, case and
  whole-cell options, Boolean results, event Target and atomic service refusals.
  The range suite now has 15 tests.
- **OM02a-5 — CurrentArray:** exposes the full legacy CSE array range from either
  its formula anchor or a spill child, using engine formula metadata. Calls on
  cells outside a legacy array return error 1004; dynamic array spills are not
  reported as legacy array groups. Covered by a focused integration test.
- **OM02a-6 — directional fill and AutoFill:** `FillDown`, `FillUp`, `FillLeft`
  and `FillRight` route bounded single-area receivers through the corresponding
  native engine commands, preserving values, relative formula rebasing and
  formats. `AutoFill` accepts a same-sheet destination that contains the source
  and extends it in one direction. It supports default, copy, series, formats,
  values, days, weekdays, months and years modes; trend and FlashFill modes
  refuse explicitly. Invalid, cross-sheet and non-containing destinations are
  rejected before mutation. The range suite now has 18 tests, including exact
  Worksheet_Change targeting for the filled destination cells.

- **OM02a-7 — Union/Intersect and Go To history:** bounded two-to-thirty argument
  `Application.Union`/`Intersect` coalesce overlaps, preserve multi-area identity,
  return Nothing for an empty intersection and reject cross-sheet inputs.
  `Application.Goto` now uses native `edit.goTo`, accepts Range or reference text,
  records/replays history, persists Scroll view state and preserves the first
  area as the active cell.
- **OM02a-8 — SpecialCells:** implements constants/formulas masks, blanks,
  comments, LastCell, visible cells, conditional-format membership and validation
  membership with bounded coalesced results and error 1004 for no matches.
- **OM02a-9 — DataSeries:** native-backed linear, growth and chronological series
  support Rowcol, Date, Step and Stop. Trend regression and xlAutoFill inference
  refuse before mutation instead of silently substituting another fill.
- **OM02a-10 — side-map clearing:** `ClearComments`, `ClearNotes` and
  `ClearHyperlinks` route through native commands. The engine now clears notes and
  links on otherwise empty cells before clipping ordinary cell work to UsedRange.
- **OM02a-11 — array/spill/navigation properties:** `HasArray`, `HasSpill` and
  `SpillParent` use live formula/spill metadata with True/False/Null and error
  behavior; `Next` and `Previous` return cells without changing Selection. The
  range suite now has 27 tests. Partial-cell Insert/Delete also use the native
  vertical Shift defaults when the optional Shift argument is omitted.
- **OM02a-12 — grouped selection and Go To source history:** `sheet.select`
  stores the selected worksheet-tab group in canonical engine/window/debug state;
  `Worksheet.Select Replace:=False` extends that group and Activate restores a
  singleton group. `Application.PreviousSelections` independently stores and
  returns the last four ranges displaced by recorded Go To operations, including
  its indexed form and bounded object-array form.
- **OM02 routing reconciliation:** 112 rows originally assigned by broad owner
  matching were moved to their actual family packets (for example Borders to
  OM05, formula navigation to OM07, Pivot members to OM13, external data to OM14,
  annotations to OM19 and sheet lifecycle to OM04). OM02 now contains 68 rows:
  64 implemented subsets, two reachable object subsets and two explicit
  platform-service results. It has no missing or legacy-review rows.

OM02's bounded profile is accepted. Cross-workbook copy consent belongs to the
OM04 multi-workbook lifecycle, trend/FlashFill and Justify/DataSeries extensions
belong to OM06, and CopyOrigin from-right/below remains an explicit unsupported
option rather than a silent substitution. `Range.CopyPicture` and
`Worksheet.PasteSpecial` retain explicit clipboard-image/external-format service
results, satisfying the packet's native/browser service-unavailable contract.

## Event remediation (2026-11-03)

The review risks were reproduced with five failing integration tests before the
fix. The adapter now prepares events at the actual write handlers, after scoped
receiver resolution, rather than partly before `Session::with_target`.

- `Bridge::prepare_changes` checks the entire event-chain/queue requirement,
  interns callback owner and exact Target handles and materializes callback
  arguments before the first engine effect. `queue_changes` consumes that
  preparation without fallible allocation. Two-sheet cuts reserve both events
  together; same-sheet cuts retain one combined source/destination Target.
- `Bridge::write_range` covers explicit Value/Value2/Formula, default and indexed
  assignments and direct Application/Worksheet range assignments. Clearing uses
  only its own handler, removing the duplicate ClearContents notification.
- Clipboard, Replace, clearing, fills and DataSeries use the same preparation.
  Failed operations drop their prepared batch without queueing callbacks. Earlier
  successful macro writes retain the existing undo-group policy.
- Handler lookup rejects other document UIDs even when code names match. Event
  delivery resolves the owner under the bound workbook's scoped target without
  leaving that workbook active; sheet synchronization remaps callback ownership
  only for the bound workbook.
- Reporting callbacks in older clipboard/range fixtures now explicitly suppress
  their own writes. A separate test verifies that an unsuppressed default
  assignment queues exactly one later callback and remains in the same undo step.

Regressions in `crates/vba/tests/events.rs` cover inactive-workbook writes and
callback writes through Me, duplicate code names in another workbook, exact
assignment/ClearContents callback counts, follow-up callback delivery, an event
batch exceeding the remaining queue capacity and handle exhaustion before copy.
Capacity refusals preserve workbook data, selection and clipboard. The existing
pause/resume, Reset, trust, event ordering, protection and undo suites remain
part of acceptance. All product changes are private to `gridcraft-vba`; engine
and tiny-vba contracts remain as reviewed.

### Remediation validation

All commands used Rust 1.99.1. Evidence logs are local under
`/tmp/gridcraft-vba-remediation-*.log`.

- `cargo +1.99.1 test -p gridcraft-vba --no-fail-fast`: **121 passed**.
- `PATH=/tmp/tiny-vba-m8-browser/node_modules/node/bin:$PATH CARGO_TARGET_WASM32_UNKNOWN_UNKNOWN_RUNNER=wasm-bindgen-test-runner cargo +1.99.1 test -p gridcraft-vba --target wasm32-unknown-unknown`: **119 passed**, using the existing cached Node 22.23.4 and runner 0.2.129. System Node 18 crashed in the runner; use Node 22 as CI does.
- `cargo +1.99.1 test -p gridcraft-ui-egui --features vba-window --test vba_events --test vba --test vba_forms --test vba_windows`: **32 passed**.
- `cargo +1.99.1 clippy -p gridcraft-vba --all-targets -- -D warnings`: passed. Workspace formatting and strict Clippy also passed through `cargo +1.99.1 xtask ci`.
- `cargo +1.99.1 run -p xtask --features vba-audit -- vba-audit`: passed; all 5,996 pinned source/semantic rows, 272 enum types and 2,260 constants unchanged.
- `cargo +1.99.1 check -p gridcraft-vba -p gridcraft-web --features gridcraft-web/vba --target wasm32-unknown-unknown`: passed.
- `target/debug/xtask layers`, `assets`, `parity --check` and `keytips --check`: passed. Local Markdown links/new section anchors and `git diff --check` passed.
- Full `cargo +1.99.1 xtask ci` **does not pass in this checkout**: UI suite `it` has 901 passing tests and two unrelated font failures, `interaction::column_autocomplete_folds_case_beyond_ascii` (missing U+1E9E) and `interaction::language_switch_renders_the_ribbon_and_persists` (missing U+1EBB). The build lacks CI's optional craft-fonts checkout; no font policy or glyph assertions were changed. The task stops at tests, so its later stages were not run as part of that invocation; individual layer/asset/generated checks and the VBA web build were run separately above.
- Fixed the existing unused `args` warning in `xtask::cmd_vba_audit` when its feature is disabled so the strict workspace Clippy gate can run. No dependency or CI workflow changed.

## Revision history

- 2026-11-03: Replaced numeric OM delivery order with High-first/selected-Medium stages; moved OM18 ahead of Medium, interleaved OM11/protection, and explicitly deferred all Low runtime work. IDs and source routing unchanged.

- 2026-11-03: Implemented event targeting and complete pre-mutation batch reservation; added regression evidence and corrected callback fixtures.
- 2026-11-03: Rust architecture review; retained the optional adapter boundary,
  specified incremental family extraction, permitted narrow host-neutral engine
  plumbing, recorded event/targeting risks, and replaced broad prerequisite and
  generator work with slice-sized delivery. Historical baseline labelled.

### Initial documentation-review validation (before remediation)

- `cargo +1.99.1 run -p xtask --features vba-audit -- vba-audit`: passed;
  5,996 source/semantic rows, 272 enum types and 2,260 constants.
- `cargo +1.99.1 xtask layers`: passed; 17 crates, no violations.
- Python standard-library checks: inventory/contract source paths match exactly;
  OM02 remains 64 implemented subsets, two object subsets and two service rows;
  local link targets in the three edited documents exist.
- `git diff --check`: passed. No runtime source or generated artifact changed;
  runtime regression suites and full CI were not run for this documentation edit.
- Local `plan/STATUS.md`, `plan/execution-plan.md` and `plan/architecture.md`
  were absent; review used committed docs and current Rust sources. Cargo emitted
  unused-patch warnings and the existing non-audit xtask unused-argument warning.
