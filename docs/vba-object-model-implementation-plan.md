# Excel VBA object-model audit and Luna implementation plan

Reviewed 2026-11-03. Status: selected High native profiles delivered; Medium OM12, OM13, OM08, OM09, OM10, OM19 and OM15 selected native profiles completed; OM20 bounded native profile reconciliation and sign-off completed. Deferred Low and unselected catalogue extensions remain visible. Scope is the
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
precedence where a member has multiple concerns, and OM20 records the reviewed member-level resolution in the generated contracts.

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
| High: invocation/lifecycle (delivered) | Remaining eligible OM03/OM04 slices as their contracts are ready; complete before High sign-off | Run/reentrancy and Open/Save/Close waits are separate slices. Implement earlier when their actual dependencies are ready; they do not block unrelated synchronous High work. |
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
| OM08 | 163 | Medium | Native window/selection/lifecycle, saved-view and watch-list profile completed; independent panes/views and geometry services remain explicit native gaps. See completion record below. |
| OM09 | 96 | Medium default; ordinary protection members High | Native protection permissions, selection rights and Allow Edit Ranges completed; IRM/ACL/protected-view/encryption-writer gaps remain explicit. Mutation protection enforcement applies throughout. |
| OM10 | 133 | Medium | Native page setup/header definitions, bounded pagination/breaks, PDF export and actual host print/download/preview continuations completed; exact printer/option/graphic gaps remain explicit. |
| OM19 | 101 | Medium | Existing local notes/links/metadata; relevant OM04/OM05 receivers and event hooks if the operation needs them. |
| OM15 | 115 | Medium | Selected native dialogs, filename pickers, FindFile, status visibility, column autocomplete and cell-checkbox controls completed; missing/unselected interaction services and Office control/settings graphs remain explicit. |
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
OM20 implements the per-row priority/disposition schema in generated contracts.
Outside-profile capability review is recorded separately from known native gaps;
High reference extensions keep their High priority and literal catalogue state.
Do not reinterpret generated MISSING counts as an instruction to implement all
rows, or hand-edit generated reports.

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

## After bounded profile sign-off

The selected High native obligations and selected Medium OM12, OM13, OM08,
OM09, OM10, OM19 and OM15 profiles are accepted by OM20. Maintain their existing
parameter, event, lifetime, service and persistence limits. Review any further
catalogue extension against an existing GridCraft capability before scheduling
adapter work; absence of a descriptor alone does not establish a product gap.
Low runtime work remains deferred. See the [sign-off report](vba-object-model-profile-report.md)
for exact remaining dispositions and source rows.

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

- 2026-11-03: Completed OM20 bounded native profile sign-off, source-kind/legacy reconciliation, live nominal graph checks and generated per-row dispositions; catalogue extensions remain explicit.

- 2026-11-03: Completed selected OM15 native interactions/controls with owned UI continuations, modal lifetime guards, native/WASM evidence and source-ledger reconciliation; OM20 sign-off is next.
- 2026-11-03: Completed selected OM19 local annotation/metadata projection with source acceptance; eligible OM15 is next.
- 2026-11-03: Started Medium OM12 with embedded chart graph/settings, native identity and routing, source-contract test evidence and native/WASM validation.

- 2026-11-03: Delivered remaining OM03/OM04 continuations and selected OM11 persistence events; reconciled High native-profile evidence and exact exclusions.

- 2026-11-03: Delivered native-profile OM07 → OM05 → OM06 → OM18 projections, interleaved ordinary sheet protection, and recorded exact limits and regression evidence. This does not close remaining High invocation/continuation work.

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


## OM03/OM04 implementation progress (2026-11-03)

Historical synchronous delivery; superseded by the [High closure record](#high-native-profile-closure-2026-11-03)
below. The old invocation/service gaps in this table describe the earlier pin,
not the current runtime. Current tiny-vba pin: `8f4bebc3650cd51054fa6786f83a1ac73e07aeaf`.

| Projection | Eligibility and native evidence | Delivered contract |
|---|---|---|
| Application.Workbooks; Workbooks Count/Item/default/enumeration/Parent/Application | `ELIGIBLE_EXISTING`: Session documents, stable document UIDs and object registry | Live, one-based and case-insensitive filename lookup, nominal Excel.Workbooks type; enumeration snapshots handles. Global collection has no document owner. |
| Workbooks.Add | `ELIGIBLE_EXISTING`: file.new and existing new-workbook sheet preference | Omitted template uses the preference; xlWBATWorksheet creates one worksheet. Other sheet types and file templates refuse. Maximum 64 documents. |
| Workbooks.Open | `ELIGIBLE_EXISTING`: engine file.open and native import/password decoding | Explicit Filename and optional Password; omitted/zero UpdateLinks; omitted/False ReadOnly. Other supplied options refuse. Retains a pristine code-owner workbook. Import retains FromFile provenance, grants no consent and runs no macros/events. Browser filesystem errors propagate. |
| Workbook Activate/Name/Path/FullName/CodeName/Saved | `ELIGIBLE_EXISTING`: native window activation, document path/dirty state and workbook code-name metadata | Activate survives temporary target restoration. Name follows the saved file name. CodeName is read-only and empty where native metadata is absent. Saved is read-only live dirty-state projection. |
| Workbook Save/SaveAs/SaveCopyAs/Close | `ELIGIBLE_EXISTING`: file.save/saveAs/close, existing serializer/link rebase and file writer | Explicit supported OOXML destinations; matching supplied FileFormat only. Unsupported persistence options refuse. Save requires an existing path. SaveCopyAs preserves original path/dirty/saved state. Dirty Close requires explicit SaveChanges; failed saves keep the document open. No save/close prompts or cancellation callbacks are simulated. Closing the code owner ends its host through existing teardown. |
| Worksheets/Sheets.Add | `ELIGIBLE_EXISTING`: home.insertSheet | Before/After worksheet objects, mutually exclusive; omitted destination inserts before active sheet; Count 1..255; worksheet type only. Returns first inserted sheet, activates it and remaps existing handles before interning new ones. |
| Worksheet Move/Copy/Delete/Index/CodeName/Visible | `ELIGIBLE_EXISTING`: sheet.move, home.deleteSheet, sheet identities and Visibility | Same-workbook Before/After moves/copies only, native naming collisions and index remapping; deleted handles stay invalid even after subsequent insertion. Copy of a sheet with CodeName, tables or preserved package parts refuses because the native copy path cannot transfer those features without loss. No-destination/new-workbook and cross-workbook transfers refuse. Visible supports Visible/Hidden/VeryHidden through native protection and last-visible-sheet gates. Hidden-sheet deletion leaves the last visible sheet intact. |
| Workbook Protect/Unprotect/ProtectStructure (ordinary OM09 High dependency) | `ELIGIBLE_EXISTING`: review.protectWorkbook and native password verification | Structure:=True subset, optional password; Windows protection refuses. Failed password verification keeps protection. Existing mutation protection remains enforced. |
| Application Name/Version/OperatingSystem | `ELIGIBLE_EXISTING`: GridCraft build/target metadata | Reports GridCraft identity/version and the compilation target OS; does not impersonate an Excel installation. |
| Application CalculateFull/CalculateFullRebuild/Iteration/MaxIterations/MaxChange | `ELIGIBLE_EXISTING`: native recalc_all rebuild and formulas.calculationOptions | Full rebuild uses existing calculator; settings read active document state and setters affect currently open documents. All participants require consent and pre-reserved undo groups before mutation. Invalid bounds refuse before writes. Newly opened documents retain their native iteration settings. Existing Calculation/Calculate now apply the same all-participant consent/group preflight. |
| Application.Run | `SERVICE_UNAVAILABLE`: pinned foreign Reply supports Complete/Error/Wait, without nested invocation request/result routing | Explicit runtime error 1004 and SERVICE_REQUIRED descriptor. A fresh interpreter would lose live project globals, execution budgets and reentrancy semantics; no substitute interpreter is used. This remains required eligible OM03 work once a reviewed generic invocation contract exists. |

The native seams live in `Session::open_workbook`,
`Session::new_single_sheet_workbook` and `Session::set_worksheet_visibility`.
Their command envelopes remain recordable/replayable (`file.open` keepBlank,
`file.new` sheets, `sheet.visibility`). `file.saveCopy` uses existing serialization,
link rebasing and IO without replacing the document's saved snapshot. No provider,
chartsheet, template, read-only document mode, interactive continuation, event hook
or VBA binary project editor was added. Existing Workbook_Open/Worksheet_Change
ownership and consent regressions remain authoritative; lifecycle cancellation
and ordering must accompany future host continuation support.

The shared descriptors, nominal types, host globals and generated semantic contracts
include these members. The historical inventory's baseline statuses are retained;
contract dispatch evidence is refreshed by the existing Rust vba-audit exporter.
Regression evidence is in `crates/vba/tests/workbooks.rs`: live collections,
code-owner/active separation, stale handles, native file round trips, failed
save/close, imported-document consent, same-workbook transfers, explicit service
refusals, protection passwords and app-wide grouped calculation.

Validation for this delivery:

- Engine and VBA test suites passed; final workbook suite: 29 tests. Native
  lifecycle regressions and all existing event/clipboard regressions passed.
- Strict Clippy passed for engine, VBA and the audit-enabled xtask; audit-enabled
  xtask tests: 28 passed. Formatting and `git diff --check` passed.
- Native structural checks passed: assets, layers (17 crates), wasm, parity and
  keytips. VBA including its tests also compiled for wasm32-unknown-unknown.
- Regeneration and subsequent read-only vba-audit check passed: 5,996 source and
  semantic rows, 272 enum types, 2,260 constants. Current descriptor evidence
  updates result states without rewriting the historical inventory baseline.
- Full workspace CI passed formatting/Clippy and stopped on two pre-existing UI
  font failures: `column_autocomplete_folds_case_beyond_ascii` lacks `ẞ`, and
  `language_switch_renders_the_ribbon_and_persists` lacks `ẻ`. Both also failed
  before this delivery; no UI/font code was changed to suppress them.


### Font-enabled CI retest (2026-11-03)

The CI-pinned font checkout was already installed at `target/craft-fonts`, revision
`b076417b5880bd4154bd22f00695b1d28c6bea83`; the earlier test builds had not enabled
it. Rebuilding with the required font input resolves both missing-glyph failures.
Both exact regression tests pass, the full UI integration suite passes all 903
tests, and `cargo xtask ci` passes all ten stages, including tests, doctests and
WASM. No font policy, glyph assertion or runtime code changed for this retest.

Run from the GridCraft checkout:

```sh
CRAFT_FONTS_DIR="$PWD/target/craft-fonts" CRAFT_FONTS_REQUIRED=1 cargo xtask ci
```

This supersedes the no-font CI limitation above for builds using this font input.
Font assets remain outside tracked GridCraft source.

## OM07 → OM05 → OM06 → OM18 native-profile delivery (2026-11-03)

The four High families now have native-backed projections in the order above.
This delivers the bounded `GridCraft.Excel.v1` behavior below, rather than
certifying every source page or optional Excel behavior in those families.
The shared semantic exporter continues to label projected Microsoft contracts
`PARTIAL`; supported execution and complete Excel parity are separate claims.
The historical inventory and Low deferrals are unchanged. The later High closure
record below delivers the remaining OM03/OM04 invocation/continuation obligations.

| Packet | Delivered native projection | Deliberate limits |
|---|---|---|
| OM07 | Application/Worksheet/Range Evaluate; live workbook/sheet Names and identity-qualified Name objects; scoped names, visibility, comments, RefersTo A1/R1C1/local variants and RefersToRange; Formula2, R1C1, local and legacy FormulaArray; generic WorksheetFunction dispatch using the existing calculator registry, typed scalar/array/reference arguments, CVErr versus runtime-error behavior | No mathematical implementations or UDF mutation privilege added. Missing native functions are `OUT_OF_SCOPE_GRIDCRAFT`. Formula2 uses the native dynamic-array evaluator; this does not implement an additional Excel implicit-intersection dialect. Formula locales follow the native English/Spanish entry tables. Qualified `Excel.Name` declarations now bind through the shared contextual type parser and the native Name resolver. Range.Evaluate is an explicitly requested GridCraft extension, without a corresponding pinned Excel Range reference page. |
| OM05 | Native font/fill/theme/tint and mixed reads; alignment, wrap/shrink/indent/orientation, protection and number formats; row/column dimensions, Hidden and AutoFit; merge geometry; Borders/Border; Styles catalog, Add from a native range, and Range.Style application; live FormatConditions, cell/expression rules, scales/bars/icons/top/unique/average rules, criteria/color objects, AppliesTo, priorities, StopIfTrue and native rule modifications | Native style model has no indexed palette, independent RGB tint, localized number-format language or separate named-style editing/deletion command. Non-English NumberFormatLocal and unavailable options refuse. Named-style reads require an exact native style match. Native row/column metrics remain the product's own metrics. Conditional formatting exposes model-backed fields, not all Excel differential-style, icon-override or data-bar UI variants. |
| OM06 | Persistent worksheet/table Sort/SortFields/SortField configuration; explicit keys, headers, stable native ordering, multiple levels, custom order, case and orientation; live sheet/table AutoFilter/Filters/Filter and native criteria; Validation; RemoveDuplicates; delimited/fixed-width TextToColumns; Subtotal; row/column grouping, Outline and ShowLevels; AdvancedFilter, Consolidate, GoalSeek, snapshot data tables and Scenarios | SortMethod Pinyin/Stroke is `OUT_OF_SCOPE_GRIDCRAFT`; native codepoint ordering is not those collations. Excel color-priority/icon/DataOption sorting, per-column dropdown visibility, typed field-conversion options, linked consolidation, selected-copy header lists, replacement of existing subtotal rows, scenario locking/visibility and other unrepresented options refuse. Data tables retain the native static snapshot contract, not live TABLE formulas. Nested Array(Array(...)) is rejected by tiny-vba before dispatch; FieldInfo accepts a two-column VBA array. |
| OM18 | Local range-backed ListObjects/ListObject, ListRows/ListRow, ListColumns/ListColumn; headers/data/totals ranges, names, style flags, resize, insert/delete, Unlist, calculated range writes, native total calculations, table Sort/AutoFilter and structured-reference resolution | Provider-backed sources, refresh/publishing/connection variants remain Low/deferred. ShowHeaders reads existing state; a separate header-visibility toggle is unavailable. Local native table styles are used. No provider or new calculation behavior added. |

Private family modules own Excel binding rules and descriptors. They use one
existing document-qualified handle store; no second workbook/style/table model
was introduced. Native opaque identities survive COW/undo for defined names,
conditional formats, tables and columns, so delete/recreate does not resurrect
stale handles. Scenario records retain opaque IDs in the existing native hidden
name storage. Native commands continue to own protection, history, recalc,
recording and XLSX serialization. A macro's edits share the existing undo group.

Narrow engine seams expose existing behavior: reference-preserving formula
results, targeted table row/column compositions, partial conditional-rule updates,
typed legacy-array placement and read-only output planning for data writers.
Text splitting, advanced filtering, consolidation and subtotals reserve callback
destinations from the native plan before writes; failed preflight queues no
callback. Ordinary Worksheet Protect/Unprotect/ProtectContents bindings accompany
the mutations, including existing permission flags and password verification.
Sort configuration is adapter invocation state until Apply, not saved workbook
content. Saved range coordinates follow known structural edits and invalidate
on untracked external structural epochs, as in the accepted foundation.

Matrix boundaries now convert explicitly between the calculator's row-major
storage and tiny-vba's column-major arrays, including worksheet UDF arguments
and returns. This fixes subscript errors that a read/write round trip alone could
hide. The UI passes its native formula-entry locale to the bridge; headless
hosts can set that locale directly.

Regression evidence: `crates/vba/tests/high.rs` covers typed functions/references,
rectangular arrays and subscripts, local/scoped formulas, legacy arrays, stale
names/rules/tables, native styles/geometry, rule-specific children, inactive-sheet
sort/filter/outline operations, validation, scenarios, GoalSeek, snapshot tables,
writer plans/events, protected/invalid zero-write paths, grouped undo and XLSX
round trips. Existing range, event, workbook, UDF and macro regression suites
remain required. Final validation results are recorded with this delivery below.

Final validation for the native-profile delivery:

- `CRAFT_FONTS_DIR=/home/coder/src/gridcraft/target/craft-fonts CRAFT_FONTS_REQUIRED=1 cargo xtask ci`:
  all 10 stages passed (i18n-report, formatting, strict workspace Clippy,
  workspace tests, doctests, assets, layers, WASM, parity and keytips). The fonts
  are the existing external pinned checkout, not new repository assets.
- `crates/vba/tests/high.rs`: **17 passed**, including stale table Sort/Filter
  views after native table-ID reuse and exact native writer event destinations.
  `crates/vba/tests/worksheet.rs`: **17 passed**, including independent UDF
  rectangular-array argument and return subscript checks.
- Optional `gridcraft-ui-egui --features vba-window` suites `vba`, `vba_events`,
  `vba_forms`, `vba_windows`: **32 passed**.
- `cargo run -p xtask --features vba-audit -- vba-audit --write`, followed by
  the non-writing audit: passed; **5,996 source paths/semantic rows, 272 enum
  types and 2,260 typed constants** retained. Audit unit tests: **3 passed**.
  Historical inventory statuses were not rewritten. Source rows with current
  implemented descriptors: OM07 373, OM05 208, OM06 100, OM18 51; these are
  bounded profile entries, not a parity percentage or a full-packet certificate.
- Optional VBA web build (`gridcraft-vba`, `gridcraft-web/vba`) and VBA test
  compilation on `wasm32-unknown-unknown`: passed. The optional web build still
  emits the pre-existing UI `Modified::Unknown` dead-code warning.
- `git diff --check`: passed. No dependency pin or CI workflow changed.

## High native profile closure (2026-11-03)

Status: **OM03, OM04 and selected High OM11 contracts delivered**, completing the
selected High native profiles after OM07 → OM05 → OM06 → OM18. The machine-readable
[High acceptance ledger](vba-object-model-high-acceptance.json) lists 13 required
obligation groups, exact native evidence, source IDs and authored Rust tests.
`vba-audit` rejects missing test functions/native evidence, duplicate obligation IDs,
unknown source IDs and unresolved required obligations. Exclusions remain separate.
This closes implementation against the accepted existing-capability profile; it
does not certify every option on a Microsoft page or migrate all historical inventory
rows into v1 priority/disposition fields. Source contracts remain PARTIAL and the
full unresolved source catalogue is retained. General Medium work can follow; Low
runtime work remains deferred.

Current dependency: published tiny-vba commit
`8f4bebc3650cd51054fa6786f83a1ac73e07aeaf` on `feat/suspended-host-project-calls`,
with ADR 0035 and generic native/WASI regressions. All GridCraft tiny-vba packages
use that immutable revision. No local path override or sibling checkout is required.
The additive IDE/IR8 capability preserves older consumer targets and the existing
foreign protocol; WordCraft's independently pinned existing API needs no migration.

- **OM03 Run:** public standard procedures in the running project, optional module
  qualifier and matching workbook qualifier; up to 30 typed arguments with Missing,
  optional defaults and ParamArray. Nested results resume the parked caller in the same VM. Arguments are positional.
  Lookup/signature failures are catchable, while an unhandled child runtime error
  stops the entire invocation, even when the caller has On Error Resume Next. Globals, classes, debugging, undo ownership and
  lifetime budgets survive suspension. Run argument snapshots do not write callback
  ByRef outputs into the original call's argument slots. Ambiguous/private/missing
  targets and another project's macros refuse explicitly; there is no native shared
  project registry to authorize and route a different project.
- **OM04 waits:** native open/save pickers and Save Changes decisions are explicitly
  injected services. The bridge owns one bounded operation/token through prompt,
  callbacks and commit; idle prompt frames do not consume VM host-poll budget.
  Named/omitted arguments and unsupported options are checked before callback or
  picker effects. Wrong/duplicate/stale tokens cannot replay an operation. Reset,
  source staleness, trust revocation and close retain existing cancellation policy.
  SaveCopyAs retains the native path/saved-state contract. Imported code stays inert.
- **OM11 persistence:** Workbook_BeforeSave/BeforeClose precede native effects,
  then connected Application WorkbookBeforeSave/WorkbookBeforeClose sinks run in
  connection order with the current ByRef Cancel. Signatures and complete handler
  snapshots are validated before callbacks. The shell collects a lifecycle batch's
  ByRef outputs before ordinary queued Change callbacks can replace them. Cancellation prevents persistence/close;
  VBA errors propagate without pretending success. AfterSave receives actual native
  success/failure after IO. SaveCopyAs does not manufacture Save events. Existing
  EnableEvents suppression, FIFO mutation events, owner binding and undo survive.
- **Native shell:** save/Save As and final document close hold their engine command
  while callbacks run/pause/wait. The existing Save Changes dialog resumes the
  approved close once, including asynchronous Save; closing one of multiple views
  does not fire Workbook.BeforeClose. Native quit and root close honor cancellation
  for the explicitly attached project, including clean documents. Other imported
  packages do not acquire executable projects during shutdown. Control callers see
  an explicit pending lifecycle result and can inspect `vba.status.lifecycle`.
  Deferred commands and Save Changes decisions retain document IDs across activation;
  a callback or user window switch cannot redirect the final save/close.
- **Platform/scope:** headless callers configure prompt services or use explicit
  arguments. Browser builds retain no native path picker/filesystem authority.
  Lossless cross-project sheet transfers, chartsheets/templates/read-only mode,
  providers and edited binary VBA writeback remain out of scope. Additional Excel
  event pages without existing native event hooks are not inferred from polling;
  Medium window/chart and Low provider/remote events retain their feature priority.

The earlier OM03/OM04 progress table and its service-unavailable Run/prompt statements
are historical. Current reference/diagnostic metadata marks Application.Run implemented;
the audit records all six persistence event pages and their acceptance test IDs.
Qualified `Dim item As Excel.Name` is covered by a native/WASM nominal regression.

Validation results are recorded after the final native/font, optional-VBA and WASM
gates below; no fresh Excel oracle comparison is claimed by these authored tests.

Final closure validation (published pin `8f4bebc3650cd51054fa6786f83a1ac73e07aeaf`):

- `CRAFT_FONTS_DIR=/home/coder/src/gridcraft/target/craft-fonts CRAFT_FONTS_REQUIRED=1 cargo xtask ci`:
  all 10 stages passed, including 163 native adapter tests.
- Strict Clippy for the VBA adapter, `ui-egui/vba-window` and audit-enabled xtask:
  passed with `-D warnings`. The five optional VBA UI suites passed **42 tests**,
  including 10 lifecycle tests for cancellation, Reset, save ordering, failed IO,
  native quit, multiple views, pickers and document identity across activation.
- Node 22.23.4 with wasm-bindgen-test-runner 0.2.129:
  `cargo test -p gridcraft-vba --target wasm32-unknown-unknown` passed **156 tests**.
  Use the cached Node 22 path as in the earlier remediation command; system Node
  18 is unsuitable for this runner. Native filesystem-only tests remain native.
- `cargo run -p xtask --features vba-audit -- vba-audit --write`, followed by its
  non-writing check: passed for 5,996 source pages, 272 enum types, 2,260 typed
  constants and all 13 required High obligation groups. Audit-enabled xtask:
  **28 tests passed**. ScreenUpdating and unsupported external-format/image
  clipboard services are classified out of scope, rather than unfinished High.
- `cargo check -p gridcraft-web --features vba --target wasm32-unknown-unknown`:
  passed. The reused native Save Changes dialog was rendered offscreen and visually
  inspected; no assets or new application UI were introduced.
- tiny-vba at the published revision: full CI, WASM and WASI gates passed; all
  **7** generic suspended-call tests also passed under WASI, including fatal policy,
  wrong/stale tokens, pause/Reset, optional/ParamArray values and shared budgets.
  Earlier compiler/default capability and contextual type cases also passed WASI.
- All 13 tiny-vba lockfile packages use the published immutable Git revision;
  no local path patch is part of this checkout. Fresh Office oracle comparisons,
  general Medium delivery and Low runtime implementation were not performed.

## Medium kickoff — OM12 embedded chart slice (2026-11-03)

Historical kickoff status (superseded by the completion record below): **first OM12 native chart slice implemented; Medium remains in progress**.
The order remains OM12 → OM13 → selected OM08/OM09/OM10/OM19/OM15. This slice
projects existing embedded charts; it does not mark the 1,220-page OM12 family,
all drawings or the selected Medium stage complete.

| Surface | Native-backed contract |
|---|---|
| Worksheet.ChartObjects / ChartObjects | Live Count, one-based/default Item, case-insensitive native generated names, For Each, Parent and Application. Existing charts are resolved by document/sheet/object identity, not collection ordinal. |
| ChartObject | Chart/Parent/Application, generated read-only Name, Delete, point-based Left/Top/Width/Height, TopLeftCell and Placement. Geometry uses native grid measurement and chart.set; dimensions are limited to the native 20..10,000 point bounds. |
| Chart | ChartType for represented native families, ChartStyle 1..16, PlotBy, SetSourceData with a same-workbook rectangular Range and optional xlRows/xlColumns, HasTitle/ChartTitle and HasLegend/Legend. Source rebinding reuses chart.selectData and preserves the explicit chart owner on inactive sheets. |
| ChartTitle / Legend | Title Text/Caption/Delete and Legend.Position; live Parent/Application. Missing title/legend access refuses; mutations retain native drawing protection and grouped undo. Native empty title text means no visible title; setting HasLegend=True preserves an existing position. |

Implementation lives in private `crates/vba/src/chart_api.rs`, using the shared
Feature registry, nominal resolver and descriptor catalogue. A process-local
Chart identity follows the existing table/sheet identity contract: it survives
COW and undo, is excluded from serialization/content equality, and distinguishes
a replacement chart even when native numeric IDs are reused. VBA property calls
snapshot only bounded scalar metadata, never the chart's series payloads.

The native chart/anchor commands now accept explicit sheet targeting. Their
existing default still targets the active chart owner, including when Select Data
names a different source sheet. This is routing plumbing; no chart calculation,
renderer, file format or product feature was added. All chart mutations use native
commands and enforce the current sheet drawing lock before effects.

`ChartObjects.Add` is explicitly OUT_OF_SCOPE_GRIDCRAFT: the existing insert.chart
command builds a data-backed chart and has no blank-chart creation contract. The
adapter does not invent dummy cells/series to pretend to create an empty chart.
Native generated chart names are exposed read-only; arbitrary Excel drawing-name
storage/rename is not added. Unsupported 3-D/legacy subtype options refuse before
effects. Series/axes/data-label child graphs, broader Shapes/ShapeRange, text
frames, pictures, grouping, z-order and export remain subsequent OM12 slices;
chart-sheet, window-only and provider behavior retain their established scope.
No new Excel event is inferred by polling chart state. Existing cell-change events
and native chart recalculation/rendering continue through their owning paths.

Authored acceptance is in `crates/vba/tests/charts.rs`: typed traversal and
inactive-sheet geometry/settings, enumeration/default/name lookup, native source
rebinding, single undo groups, numeric-ID reuse during a debugger pause, rejected
options/protection, deletion/stale children, XLSX round trips and shared renderer
primitives/data. The six test IDs are linked into generated source contracts and
validated by vba-audit. These tests are native/WASM contract evidence, not new
Excel oracle observations.

Next smallest OM12 slice: native SeriesCollection/Series data binding and formatting,
then the native Shapes graph. Keep OM13 queued behind the selected basic OM12 work.

OM12 kickoff validation:

- Font-enabled `cargo xtask ci`: all 10 stages passed, including **169** native
  adapter tests and the existing chart/model/XLSX/engine/UI suites.
- **6** new chart tests passed natively and under Node 22 / wasm-bindgen; the
  full VBA WASM suite passed **162** tests, including shared chart renderer output.
- Strict Clippy passed for the adapter and audit-enabled xtask; **28** xtask tests
  passed. Audit regeneration and non-writing verification retained all **5,996**
  pinned source pages and link six real test IDs to each of the **35** affected
  source contracts (34 native projections, one explicit out-of-scope variant).
- Optional `ui-egui/vba-window` and `gridcraft-web/vba` WASM builds passed.
  The shared tiny-vba revision remains unchanged. No corpus macro or new Office
  oracle execution was performed; no asset or workflow was changed.


## OM12 native profile completion (2026-11-03)

Status: **OM12 complete for the accepted existing basic chart/drawing profile**.
Medium remains in progress; **OM13 local pivots is next**. This supersedes the
kickoff's deferred series/drawing list. It does not certify all 1,220 Excel
reference pages. The machine-checked [OM12 acceptance ledger](vba-object-model-om12-acceptance.json)
records required native contracts, bounds, test functions and product exclusions.
The pinned source catalogue retains PARTIAL rows for unsupported Excel variants.

| Native capability | Delivered projection and boundary |
|---|---|
| Embedded charts | ChartObjects live Count/Item/default/enumeration/Delete; ChartObject ownership, generated names, Index, anchor cells, geometry/placement, ShapeRange and native z-order; Chart source/type/style/orientation/title/legend, Delete, Refresh and Location to an existing worksheet. Native chart.move keeps its destination `sheet` distinct from the optional `sourceSheet` owner. |
| Series | Typed SeriesCollection/Series, same-workbook Range or bounded array bindings, name/values/categories/bubble sizes, native Add with inferred labels, Delete, PlotOrder, ChartType, AxisGroup, Smooth and inclusive boxplot quartiles. Reordering preserves lifetime identity; source replacement invalidates removed elements. Format exposes the native shared series fill/line color. Blank NewSeries, label overrides and writable SERIES formula strings refuse. |
| Trendlines | Native families, name/automatic name, order/period, Forward2/Backward2, intercept, equation/R-squared, indexed Add/edit/Delete; native line color, 0.25..12 point weight and represented dash styles. Collection budget is 16 per series, matching the renderer. |
| Axes and chart parts | Primary Cartesian axes plus represented secondary value axes, primary axis titles and value gridlines. Auto scales and ChartArea/PlotArea bounds come from the native renderer layout; manual scales/layout and independent child styling refuse. Native value labels are chart-wide; per-series enable/disable is accepted only for a one-series chart. Label text uses the renderer's source number format. |
| Shapes and pictures | Live native shapes/images/charts in native paint order, typed Shape wrappers, AddShape for represented presets, AddTextbox, AddLine without unsupported flips, embedded AddPicture through the existing native file service, geometry/placement, native shape OnAction, colors, anchor cells, Delete and native layer-relative z-order. Names remain native generated read-only names. Browser file reads fail explicitly; no picker is opened. |
| ShapeRange | At most 256 identity-qualified children counted against the shared object budget. Snapshot Item/enumeration/Count, bounding geometry, common-origin translations, same-layer alignment and deletion. All children, bounds and protection are checked before bulk effects. Multi-object collective scaling has no native transform and refuses. |
| Text and formatting | Textbox Characters slices count UTF-16 code units; invalid surrogate writes refuse. Whole-body native Font/ShapeFormat, paragraph alignment/reading order and Office TextFrame2/TextRange2/Font2 expose stored text fill/outline and vertical anchor. Partial character font runs, margins, autosize, gradients and effects have no native state. |
| Ownership, persistence and rendering | One document-qualified Feature registry; nonserialized lifetime tokens for charts, series, trendlines, shapes and images survive COW/undo. ShapeRange children follow sheet remapping; parked chart descendants follow native chart transfers, including external command/undo ownership changes. Mutations use native commands with explicit owner sheet and protection preflight, and retain the macro's single undo group. Native/XLSX round trips and shared renderer/raster output are covered. |

No persistent group or connector model exists. `arrange.group` only aligns
anchors, so VBA Group/Ungroup does not report success. Chart sheets, blank charts,
chart-only export, object clipboard/duplication, nested chart drawings and
independent per-point/legend-entry/effect state remain product gaps. Standalone
sparklines do not supply the missing SparklineGroup/axis/point graph; its Range
root belongs to OM06 and remains outside that packet's accepted native profile.
These boundaries are recorded in the ledger, rather than implemented as inert
objects or dummy data. No chart algorithm, renderer feature, external service,
workflow or vendor asset was introduced.

Acceptance uses the six chart tests and the drawing/series tests in
`crates/vba/tests/drawings.rs`, including native file embedding and browser service
refusal, identity reuse, sheet moves, stale children, protection, source rebinding,
UTF-16 edits, undo and XLSX/shared raster parity. Office child contracts and enum
facts are verified against the same pinned archive, separately from the unchanged
5,996-page Excel inventory.

Final validation (2026-11-03):

- Font-enabled `cargo xtask ci`: all 10 stages passed, including **179** native
  VBA adapter tests and the workspace chart/model/XLSX/engine/UI suites.
- Full Node 22 / wasm-bindgen VBA suite: **171** tests passed, including **15**
  OM12 chart/drawing cases. Native coverage has one additional picture-file test.
- Strict Clippy passed for the adapter and audit-enabled xtask; **28** xtask tests
  passed. Audit regeneration and non-writing verification passed for **239**
  required native source contracts, plus **14** sourced Office child extensions,
  while retaining all **5,996** pinned Excel pages and the existing High ledger.
- Optional native `ui-egui/vba-window` and WASM `gridcraft-web/vba` builds passed.
  The shared tiny-vba revision is unchanged; no new Office oracle execution,
  vendor asset, workflow change or external service was used.


## OM13 native local pivot profile completion (2026-11-03)

Status: **OM13 complete for the existing local native profile**. Medium continues
with the selected OM08/OM09/OM10/OM19/OM15 work; OM14 remains deferred.
The [OM13 acceptance ledger](vba-object-model-om13-acceptance.json) records 153
required native contracts, their real Rust tests and the native capability limits.
This is a native profile completion, not a claim that the entire Excel pivot,
OLAP, model and timeline reference is implemented.

| Surface | Native-backed contract and boundary |
|---|---|
| Worksheet.PivotTables / PivotTable | Live one-based/default/name lookup and enumeration, Parent/Application, explicit xlDatabase Add from a same-workbook bounded source to a receiver-sheet cell. Explicit creation keeps the current view. Native source, cache index, report/body ranges, aggregation/layout/header/totals/style settings, AddFields/AddDataField, values-axis placement, ClearTable/ClearAllFilters, Update and RefreshTable use engine commands. Names are read-only. Native overlap/protection failures remain errors. |
| PivotCaches / PivotCache | Workbook enumeration and same-report canonical views, source text/type, index, live native record count, OLAP=False and synchronous native refresh. Native XLSX writes one cache per report; two reports with the same source stay independent. Unattached/shared cache creation, cache reassignment, provider/connection properties, retained cache records and async refresh/cancellation have no native model and refuse. These views do not simulate shared Excel caches. |
| PivotFields / PivotField | Source-qualified fields plus explicit value/calculated-field lifetimes; row/column/page/data/hidden collections, placement, native functions/show-as modes/number formats, data captions, source names, formulas and native label sorting. The native Σ Values axis supports its stored row/column and row-position options. Source/calculated names are read-only; data-field captions can change. Reordering retains native definition identity. Source changes invalidate prior field/item views. |
| PivotItems / PivotItem | Bounded native item labels, visible/hidden collections, record counts, manual visibility/report filters and stored axis collapse. At least one pivot item must remain visible. Item names are read-only. There is no independent cached-label rename or calculated-item engine. Native date grouping remains usable by the product; Excel's separate grouping hierarchy/field metadata is not fabricated. |
| CalculatedFields | Add, lookup/enumeration, formula updates and deletion through the existing native aggregate formula engine. English/default formula mode only; unsupported localized parsing refuses. Deleted/recreated definitions cannot retarget parked field arguments or data fields. |
| Range / PivotCell | PivotTable/PivotCell traversal, native report-cell roles and True-only drilldown for one native value cell. EnableDrilldown and workbook-structure protection retain native behavior. Details sheets contain actual source records and remain in native undo. Non-pivot Range.ShowDetail still uses the existing outline implementation. Report-derived geometry refuses stale report/source combinations, using the same native validation as drilldown. No field/line/axis object metadata or GETPIVOTDATA Range-returning service is invented. |
| PivotCharts / PivotLayout | Existing linked charts expose the actual PivotTable and Chart parent. Native refresh rebinds their source through the existing engine; linked drawing protection is preflighted. Native/XLSX reports, chart values and shared raster output are tested. |
| Table slicers | Existing native widgets expose SlicerCaches/SlicerCache/Slicers/Slicer/SlicerItems/SlicerItem, local ListObject binding, selection/filter clearing, deletion, native caption/column count and shape/geometry. Shape traversal uses native slicer identity and anchor commands. Cache views are per widget; standalone/shared caches, PivotTable slicer connections, OLAP levels and timelines remain native gaps. **Native XLSX has no slicer codec.** Widget state is covered by model snapshot/undo serialization, not claimed as an XLSX round trip. No new codec is added by this packet. |

Private `pivot_api`, `pivot_fields_api` and `slicer_api` use the shared qualified
Feature registry and nominal resolver. PivotTables, value fields, calculated fields
and slicers have nonserialized lifetime tokens; source epochs prevent stale source
views from silently rebinding. Scalar names/labels and input-sized collections are
bounded before projection. Native commands gained explicit pivot/slicer sheet
routing and narrow existing-state parameter plumbing for atomic area edits and
slicer formatting. No pivot algorithm, external service, new event framework,
vendor asset or workflow was introduced.

Tests in `crates/vba/tests/pivots.rs` cover typed parent identity, explicit and
inactive destinations, atomic layout/refusals, field reordering, source rebinding,
numeric ID reuse, deleted calculated-field arguments, protection, native details
records, source events, independent cache views, PivotChart refresh, XLSX/native
raster parity and slicer model snapshot limits. Existing source-cell Change events
remain intact. Internal report rewrites do not masquerade as explicit cell edits;
there is no native PivotTableUpdate/provider event source to project. Unsupported
provider/model/timeline families are classified as OUT_OF_SCOPE_GRIDCRAFT rather
than offered empty successful refreshes.

Final validation (2026-11-03):

- Font-enabled `cargo xtask ci`: **all 10 stages passed**, including **189**
  native VBA tests and the workspace engine/model/chart/XLSX/UI suites.
- Full Node 22 / wasm-bindgen VBA suite: **181 tests passed**, including all
  **10** OM13 tests. Native and WASM use the same aggregation and renderer paths.
- Strict adapter and audit-enabled xtask Clippy passed; **28** xtask tests passed.
  Audit regeneration and non-writing verification passed for **153** required
  OM13 source contracts while retaining the **5,996** pinned Excel pages and the
  existing High/OM12 ledgers. **435** unbacked OM13 provider/model/timeline and
  related source rows are explicitly OUT_OF_SCOPE_GRIDCRAFT.
- Optional native `ui-egui/vba-window` and WASM `gridcraft-web/vba` builds passed.
  The tiny-vba dependency revision is unchanged. No new Office oracle execution,
  XLSX slicer codec, external service, asset or workflow was introduced.


## OM08 native window profile completion (2026-11-03)

Status: **OM08 complete for its existing native profile**. The selected Medium
sequence continues with remaining eligible OM09, then OM10, OM19 and OM15.
The [OM08 acceptance ledger](vba-object-model-om08-acceptance.json) links **50**
required source contracts to executable Rust tests and records the exact limits.

- `Application.Windows`, `Workbook.Windows`, `ActiveWindow`, `Window.NewWindow`
  and `Workbook.NewWindow` project real native session IDs. Collections support
  one-based/default/caption lookup and bounded enumeration. Parent/Application,
  read-only Caption, Index/WindowNumber, xlWorkbook Type, SelectedSheets,
  ActiveSheet, ActiveCell, Selection and RangeSelection use native ownership.
- Reads of parked windows preserve user focus, activation recency and the live
  view of inactive workbooks. Dynamic roots follow the view currently shown by
  the UI. Explicit activation and new windows retain their chosen focus;
  next/previous cycle through native visible-list order. Global collections
  survive closing their originating workbook. Closed window handles fail 91.
- Native parked views now follow existing sheet identity across rename/reorder.
  Visible, hide/unhide and non-last Close use engine commands. Non-last Close
  ignores save options as Excel specifies. Last-window Close reuses workbook
  callbacks, save consent and prompt continuations; it returns Boolean success
  or False on callback/user cancellation. Destination workbook trust still applies.
- CustomViews expose the existing sheet-ordinal/selection/zoom/freeze snapshot.
  Add requires explicit `PrintSettings:=False, RowColSettings:=False`; unsupported
  defaults/options refuse before edits. Metadata, geometry bounds and protected
  selections are validated before Show. Names are read-only. Native XLSX retains
  the stash. Saved sheet ordinals are not persisted worksheet identities, so
  snapshots should be recreated after changing workbook sheet topology.
- Application.Watches aggregates existing native per-document lists. Add accepts
  one qualified cell via a Range; Source returns that Range, and deletion uses
  existing commands and permission checks. Native text references do not track
  worksheet rename/deletion. Ambiguous legacy relative watch entries refuse
  Source instead of rebinding to another sheet.
- Native saved-view and watch collections use their hidden defined-name stash's
  lifetime. Any edit invalidates prior child handles in that collection/document;
  reacquire them. Collections, metadata and returned handles are bounded before
  mutations. These are narrower lifetime semantics than Excel's individual objects.

Native zoom, freeze, gridlines/headings and display modes are **worksheet-owned**;
scroll persistence is also per sheet. They are not projected as independent Window
or Pane state. Native Split aliases freeze, so there is no fabricated Panes graph.
Single-window Goto Scroll=True retains existing saved-top-left behavior; with two
windows on one workbook it refuses before navigation effects. Scroll=False keeps
native per-window selection behavior. WorksheetView/SheetViews, dialog/module views,
TreeviewControl, ScrollArea and side-by-side synchronization have no eligible backend.
Window geometry, VisibleRange, pixel conversion and Arrange require an actual
window/viewport service on native and browser hosts; headless never substitutes
inferred dimensions. No window activation event source is invented.

Final validation (2026-11-03):

- Font-enabled `cargo xtask ci`: **all 10 stages passed**, including **207**
  native VBA tests and the workspace engine/model/XLSX/UI suites.
- Full Node 22 / wasm-bindgen VBA suite: **199 tests passed**, including all
  **18** OM08 regressions. Native and browser hosts share the same adapter.
- Strict audit-enabled xtask Clippy passed; **29** xtask tests passed.
- Desktop `gridcraft-ui-egui/vba-window` and browser `gridcraft-web/vba` checks passed.
- Pinned audit regeneration and no-write verification passed for all **5,996**
  source pages, preserving High/OM12/OM13 evidence and **50** OM08 required contracts.

The final close regression verifies that a BeforeClose callback creating a sibling
window cannot accidentally close the shared workbook. No UI layout, vendor assets,
workflow or tiny-vba dependency revision changed.


## OM09 native permission profile completion (2026-11-03)

Status: **OM09 complete for the existing native profile**. Medium continues with
eligible OM10 printing/export, then OM19 and OM15. The
[OM09 acceptance ledger](vba-object-model-om09-acceptance.json) links **33** required
source contracts to executable tests and retains ordinary High protection evidence.

Worksheet.Protection exposes the native stored formatting/insertion/deletion/sort/
filter flags. AllowInsertingHyperlinks and AllowUsingPivotTables remain False;
requests to grant them or UI-only/independent object/content/scenario protection
still refuse before effects. ProtectionMode is False; drawing/scenario status
reflects the existing worksheet guard. EnableSelection maps native flags only on
an already protected sheet. Its native selectionOnly command parameter preserves
the password and every other right, and refuses mixed password/permission changes.
The native locked-only combination has no Excel enum and refuses the getter;
there is no dormant preference to set while the sheet is unprotected.

Typed AllowEditRanges and AllowEditRange provide one-based/default/title lookup,
bounded enumeration, Add from a receiver-owned Range, Title, read-only Range,
ChangePassword, Unprotect and Delete. All mutations use native commands and
preserve native protected-sheet restrictions. Nonserialized definition lifetimes
survive COW/undo and rename; deleted/recreated definitions cannot rebind old VBA
handles. Definition changes and owner-sheet deletion revoke matching session
unlocks; undo never resurrects revoked authentication. Passwords are hashed by
the existing engine and are absent from handles/journals. New titles are nonempty,
trimmed and at most 255 UTF-8 bytes; password inputs are bounded to 1,024 bytes.

Range.AllowEdit uses the same native content-edit policy as actual writes,
including unlocked cells and authenticated/password-free edit ranges. Every cell
must be allowed; reads are bounded to 65,536 cells. It does not grant formatting
or structure-edit rights. Protection options, edit ranges and password hashes
round-trip through native XLSX; session unlocks do not. Collections are bounded to
8,192 entries and lightweight range views to 64 areas.

Native securityDescriptor bytes remain opaque round-trip metadata: there is no
user/principal resolver or ACL enforcement to project as Users/UserAccessList.
Code-execution consent is not IRM or ProtectedView. No ProtectedViewWindow graph,
independent chart protection, named-style protection-inclusion flag, signatures,
write reservation/sharing or retained workbook encryption-settings model exists.
The existing encrypted-file reader remains reachable through Workbooks.Open's
Password argument. Workbook.Password/HasPassword and encryption-provider setters
are not fabricated; no encryption writer is added by OM09.

Final validation (2026-11-03):

- Font-enabled `cargo xtask ci`: **all 10 stages passed**, including **216**
  native VBA tests and the workspace engine/model/XLSX/UI suites.
- Full Node 22 / wasm-bindgen VBA suite: **208 tests passed**, including all
  **9** OM09 regressions, with no separate browser permission implementation.
- Strict adapter and audit-enabled xtask Clippy passed; **30** xtask tests passed.
- Pinned audit regeneration and no-write verification passed for all **5,996**
  source pages, preserving prior High/OM12/OM13/OM08 evidence and **33** required
  OM09 source contracts.

OM08 was committed separately as `bc8751f4` before this packet began. OM09 adds
native-state plumbing and lifetime/authentication checks; it adds no vendor assets,
workflow changes, encryption algorithm/writer or tiny-vba dependency revision.


## OM10 native printing/export profile completion (2026-11-03)

Status: **OM10 complete for the existing native profile**. Medium continues with
eligible OM19 annotations and then OM15 interactions. The
[OM10 acceptance ledger](vba-object-model-om10-acceptance.json) links **86** required
source contracts to native, WASM and actual UI-service integration tests. Printer
presentation and platform capability limits remain explicit; PDF bytes alone are
never described as a submitted job or a displayed preview.

PageSetup is a live worksheet-owned view of native PrintSettings: supported paper
IDs, orientation, point/inch margins, a rectangular print area, row/column titles,
scale/fit counts, centering, gridlines/headings and raw header/footer sections.
Supported paper mappings are Letter, Tabloid, Ledger, Legal, Executive, A3/A4/A5,
B4/B5 JIS and 11x17. Other native papers without a supported published Excel mapping
refuse. Margins are 0..1440 points; Zoom is 10..400 or False; set Zoom=False before
fit counts because native has no dormant fit-count storage. Zero/False fit counts
mean unlimited, up to the native 32767 limit. Order reads native xlDownThenOver;
there is no independent ordering setting. Oversized headers and unsupported options
refuse before native truncation or clamping.

FirstPage/EvenPage and their HeaderFooter.Text properties edit the existing native
variant definitions, preserving the native combined 255-character limit. Pages
Count/Item/enumeration use native per-sheet pagination, capped at 2000 pages.
Numbered Page/HeaderFooter nodes are read-only document-revision snapshots;
reacquire them after native revision changes. Definition views remain live. Manual
H/VPageBreaks project stored vectors; automatic boundaries come from the same native
layout. Add, Delete, manual Type, horizontal Location, single-row/column Range.PageBreak
and ResetAllPageBreaks use pageLayout.pageSetup. Manual markers are bounded to 1026
per axis. Break handles also expire on native revision changes, preventing stale
coordinates from rebinding after deletion/recreation. Partial regions, automatic
extent, vertical Location assignment and DragOff have no eligible native model.

Workbook PDF exports all visible sheets; Worksheet exports its explicit receiver;
Range exports one rectangle and intersects PrintArea unless IgnorePrintAreas=True.
The native service gained explicit sheet/selected-group and print-area override
plumbing while retaining its existing pagination/rendering algorithms. Native files
use the existing atomic writer and retain the workbook's source path/state. Missing
Filename uses the actual native save picker when connected; cancellation returns
1004 without writing. Browser PDF export hands owned bytes to the real download
callback, using the explicit/default PDF name; it never claims a browser filesystem
write. Standard-quality PDF with native document metadata is supported. XPS, minimal
quality, metadata suppression, From/To intervals, post-export viewer opening and
external extension-class pointers refuse before effects.

PrintOut delegates to the existing native printer worker or browser print callback.
Native copies are 1..999; native printer names are validated against the real
configured/discovered printer list. ActivePrinter reads/sets that same UI state.
Workbook/worksheet collections print visible sheets; Window/SelectedSheets print
the actual native selected group. Native worker receipts/errors complete/fail the
suspended VBA call; browser completion acknowledges callback handoff, not physical
printing. Browser printer/copy overrides and native unsupported collation/page-range/
print-to-file options refuse. An unavailable printer never submits a job.

PrintPreview opens actual Backstage for workbook/single-sheet scope and suspends
VBA until the preview closes. Range/group-specific preview and EnableChanges=False
have no corresponding product UI representation. Missing host services return 1004
without a fake success. Owned print requests share the existing host continuation
manager: no document borrow survives a picker, preview or worker wait; document
liveness/trust is rechecked; reset cancels unsubmitted requests without replay.
An already submitted spooler job cannot be retracted by resetting the VBA editor.

Independent chart page setup/export, per-object print inclusion, header/footer
Graphic/Picture graphs, persistent monochrome/draft/first-page-number/comment/note/
error options, printer-resolution state and PrintCommunication batching have no
eligible native state. Existing embedded charts still render through worksheet PDF.
CustomView.PrintSettings stays False; native saved views do not snapshot page setup.

Final validation (2026-11-03):

- Font-enabled `cargo xtask ci`: **all 10 stages passed**, including **232**
  native VBA tests and the workspace engine/model/PDF/XLSX/UI suites.
- Full Node 22 / wasm-bindgen VBA suite: **224 tests passed**, including all
  **16** browser-capable OM10 tests (download/print transport use owned mock services).
- Optional `vba-window` editor suites: **27 tests passed**, including **5** real
  UI-service integration regressions for native worker settings/errors, unavailable
  printers, actual preview completion and PDF picker success/cancellation.
- Strict optional-editor and audit-enabled xtask Clippy passed; **31** xtask tests passed.
  Browser `gridcraft-web/vba` check passed.
- Pinned audit regeneration and no-write verification passed for all **5,996**
  source pages, preserving earlier packet evidence and **86** required OM10 contracts.
- Exported native PDF was rasterized and inspected: A4 landscape **841.89 × 595.28**
  points, correct worksheet/range content, header/footer and page numbering. Actual
  Backstage capture was inspected after fixing the editor presentation: preview
  temporarily returns to the workbook and restores prior editor visibility afterward.
  Generated validation images/PDFs remain outside the repository.

OM09 was committed separately as `9d8e1b5d` before this packet began. OM10 remains
an existing-native-capability projection, with no vendor assets, workflow changes,
new renderer/codec/print controls or tiny-vba dependency revision.


## OM19 local annotation/metadata profile completion (2026-11-03)

Status: **OM19 complete for the selected existing local native profile**. Medium
continues with eligible OM15 interactions and built-in controls. The
[OM19 acceptance ledger](vba-object-model-om19-acceptance.json) links **53** source
contracts to real VBA integration tests. This is a bounded native projection;
it does not establish the full Excel annotation family or complete Medium.

Delivered behavior:

- Classic Comments/Comment, Range.Comment/AddComment/NoteText, same-sheet
  navigation, author labels, idempotent visibility, deletion and enumeration.
  Comment.Text uses UTF-16 positions; edits that split surrogate pairs fail.
  NoteText reads/writes chunks of at most 255 UTF-16 units.
- Local CommentsThreaded/CommentThreaded, AddCommentThreaded, root text editing,
  AddReply, typed Replies collections and navigation within each thread.
  Author.Name reads the existing local display label. It does not synthesize
  provider or user IDs. Individual reply editing/deletion has no native command
  and fails explicitly, as do remote identity and local timestamp requests.
- Worksheet/Range Hyperlinks, cell-only Add, Item/Count/enumeration, native
  Address/SubAddress, ScreenTip and TextToDisplay updates, Range and deletion.
  Internal Follow invokes existing link navigation. External Follow and browser
  request/history/new-window options fail before queuing an external action;
  the UI's OpenUrl queue is not an acknowledged VBA navigation continuation.
- BuiltinDocumentProperties adapts all six native text fields: Title, Subject,
  Author, Keywords, Comments (description), Company. Office.DocumentProperties
  and DocumentProperty retain native parents and exact Excel indexes
  1/2/3/4/5/21; other fields fail instead of returning fabricated defaults.
- All mutations use existing engine commands and macro undo groups. Annotation
  maps are separate from sparse cell content, so the adapter adds the missing
  protected-sheet guard before any note/link mutation. Existing consent and
  document/sheet routing apply. Structural VBA edits rebase cell-owned handles;
  explicit OM19 Delete invalidates parked annotation/reply/author handles.

Native limits remain visible: no custom workbook/worksheet-property store,
rich-text note Shape graph, shape hyperlink anchors, remote author identities,
sound-note service, favorites/linked-document creator or Excel calculation
thread configuration. Cell-map handles do not track a native cut to another
cell; a delete/recreate performed wholly outside these VBA members may rebind
an old cell handle. External topology epochs invalidate feature ranges. No
independent COM object lifetime is claimed. Same-sheet navigation uses Nothing
at boundaries in this native profile.

Tests exercise protected/untrusted receivers, UTF-16 editing, duplicate refusal,
collection defaults and nominal types, root/reply navigation, copy/cut payloads,
structural rebase, stale deletion handles, internal/external navigation, six
built-in metadata fields, macro undo and existing note/thread/link XLSX codecs.
The unchanged 5,996-page pinned source inventory and shared Office source archive
remain the provenance for regenerated contracts. No intrinsic-language, formula,
provider, drawing, storage or collaboration feature was added.

Validation:

- `CRAFT_FONTS_DIR=/tmp/gridcraft-acceptance-fonts CRAFT_FONTS_REQUIRED=1 cargo xtask ci`:
  all **10** steps passed, including **3,014** native tests. The first run without
  the optional font input failed two existing UI glyph tests (ẞ/ẻ); using the
  cached approved acceptance-font bundle resolved both without source changes.
- Full native VBA suite: **239** passed before the final hostile-input regression;
  repository CI subsequently passed all **8** final OM19 tests with the final code.
- Full Node 22 / wasm-bindgen VBA suite: **231** passed. After final allocation
  hardening and the added hostile-input regression, the final OM19 WASM suite
  passed all **8** tests.
- Final strict adapter/audit clippy, generated-source audit write/check and
  `git diff --check` passed. Audit retains **5,996** pinned Excel pages, **53**
  OM19 required contracts and **10** supplemental Office metadata members.

Eligible OM15 interaction/control projection is next; Low runtime work remains deferred.


## OM15 selected native interaction/control profile completion (2026-11-03)

Status: **OM15 complete for the selected native profile**. The selected Medium
profiles have landed; **OM20 profile reconciliation/sign-off is next**. This does
not declare all Excel interaction APIs or every Medium member implemented. The
[OM15 acceptance ledger](vba-object-model-om15-acceptance.json) records **14**
canonical source contracts, two inline CellControl members and three inline
constants on the pinned `Range.CellControl` page.

Delivered projections:

- Application.Dialogs/Dialogs/Dialog expose seven existing native dialogs using
  their exact Excel IDs: Column Width 47, Row Height 127, Zoom 256, Insert Cells
  55, Delete Cells 54, Define Name 61 and Hyperlink 596. Count/enumeration cover
  this native set; Item uses Excel IDs. Show accepts omitted arguments only,
  waits for the actual dialog and returns real OK/Cancel results. Invalid form
  submissions keep the dialog open. Both keyboard and control-channel completion
  use the actual native command; no compatibility-only modal is created.
- GetOpenFilename/GetSaveAsFilename return native picker paths or Boolean False
  on Cancel without opening/saving. Native service filters/title/buttons are
  fixed, single selection only; unsupported options refuse before invoking the
  service. Save suggestions use InitialFilename or the native workbook title.
  FindFile additionally invokes the existing workbook loader after a successful
  native pick. Browser filename-based open and absent pickers fail explicitly.
- DisplayStatusBar reads/changes actual UiState visibility through the native
  UI command `view.statusBarVisible`; there is no fake headless setting or
  new custom-status text model.
- Range.AutoComplete shares the existing editor algorithm through a small
  toolkit-free engine helper. Unicode case folding, contiguous-column matching,
  ambiguity/no-match behavior and the 5000-row scan in each direction are kept.
  VBA adds bounded input and aggregate work; it does not depend on the editor's
  AutoComplete preference or change selection.
- Range.CellControl exposes the existing cell-checkbox format. Type returns
  xlTypeNone=0, xlTypeCheckbox=2 or Empty for mixed formats; SetCheckbox calls
  the native insert.checkbox command and preserves existing values/formulas.
  Empty cells become False. Worksheet_Change targets are reserved before edits
  and include only cells whose values change, once. Formatting-only repeated
  calls produce no value event. Protection, undo, structural rebasing and XLSX
  persistence use the existing paths.

All UI work crosses owned tokenized continuations. No document borrow or store
lock survives a wait. Native document/window/sheet/selection/range-epoch and
consent checks run before modal effects. Invalid or oversized selections fail
before requests are delivered. Responses are delivered once and type-checked;
duplicate and reset/close-stale tokens refuse. Owner changes or revoked consent
cancel the pending work. The editor yields to the existing modal and its prior
visibility returns on completion/reset. No extra provider, Office control
runtime, new dialog or tiny-vba intrinsic implementation was added.

Explicit profile limits remain: InputBox has no generic typed/range input
service and fails 1004, Office.FileDialog has no equivalent configurable object
graph, and dialog presets/custom picker filters/multiselect are unsupported.
Range.CheckSpelling and FunctionWizard have existing tools but no selected
range-owned VBA completion contract; they remain visible as service-needed
rows outside this selected native profile. Independent ScreenUpdating,
DisplayAlerts, custom StatusBar, Cursor/OnKey overrides, AutoCorrect and
ErrorCheckingOptions state, dialog sheets, smart tags, speech/research providers
and arbitrary Office/ActiveX/OLE controls are not fabricated. Existing shape
OnAction/assigned-macro wiring remains OM12/OM11 evidence, and generic user forms
remain the shared tiny-vba contract. Low runtime work stays deferred.

Validation:

- `CRAFT_FONTS_DIR=/tmp/gridcraft-acceptance-fonts CRAFT_FONTS_REQUIRED=1 cargo xtask ci`:
  all **10** steps passed, including **3,023** native tests and all **9** OM15
  adapter tests. The existing approved font input was used; no fonts/assets or CI
  workflow changes were made.
- Full Node 22 / wasm-bindgen VBA suite: **240** tests passed, including **8**
  OM15 tests. The successful native FindFile filesystem test is native-only;
  browser/headless absence and owned service transport are tested separately.
- Optional VBA UI lifecycle/forms/events/printing/interactions regression run:
  **49** tests passed, including **6** new interaction tests. The separate
  `vba-window` regression run passed **4** further tests.
- Strict clippy with optional `gridcraft-ui-egui/vba` and `xtask/vba-audit`,
  final formatting, generated-source audit write/check and whitespace checks
  passed. All seven dialog ID mappings reached their existing native dialog.
- Rendered and visually inspected the actual paused VBA-owned Column Width
  dialog (`/tmp/gridcraft-om15-dialog.png`). The native field and OK/Cancel are
  visible, the editor yields to the modal, and keyboard/accessibility harness
  checks pass. The screenshot remains a local artifact.
- The unchanged **5,996** pinned Excel pages remain reconciled. OM15 has **14**
  required canonical callable contracts, **2** inline CellControl members and
  **3** inline constants. Its **115** pages classify as **15 PARTIAL**,
  **3 PLATFORM_SERVICE_NEEDED** and **97 OUT_OF_SCOPE_GRIDCRAFT**; page counts
  are not feature counts or a compatibility percentage.

OM20 profile reconciliation/sign-off is next. Low runtime work remains deferred.


## OM20 bounded native profile completion (2026-11-03)

Status: **accepted for the declared bounded native profile**. The
[OM20 specification](vba-object-model-om20-acceptance.json) joins the eight
accepted High/selected Medium ledgers. The [generated report](vba-object-model-profile-report.md)
and [machine-readable witnesses](vba-object-model-profile.json) resolve **673**
required source IDs, retain all **5,996** pinned source pages and validate
**124** nominal object types reachable from Application, **45** collection
Count/Item/default contracts and **47** consumed enum types. Ten reviewed dynamic
return edges retain linked native regression evidence. This is static graph
validation plus the owning regression suites; it does not prove every Excel
parameter combination or every possible runtime path.

All **283** legacy Graph pages retain a separate source namespace. **Seven**
case-folded ID collisions keep method/property/event kinds and original paths;
activation methods no longer falsely claim activation event support. **Thirteen**
member-level routing corrections preserve original inventory packets. Four
deprecated enum pages lack numeric source tables and remain explicit exceptions;
they cannot be consumed by implemented typed signatures. Eleven forecast constants
are recovered from headerless source tables. Range.Replace's FormulaVersion
metadata now uses the documented Variant; supplied unsupported versions still
raise the existing error.

Every source row has independent priority, profile disposition, eligibility
review boundary and literal result state. **885** High reference-extension pages
retain High priority outside the accepted native obligations; their full
GridCraft capability review is **not claimed complete**. **1,106** unselected
Medium pages and **1,078** deferred Low pages are retained. Missing descriptors
alone do not prove absent native features. Exact out-of-scope/service rows remain
in the contracts. Acceptance closes the named native obligations, not all High
source members, all Medium families or general Excel compatibility.

No new UI or service behavior is introduced by reconciliation. Existing owning
ledgers retain UI, lifecycle, permission, persistence and native/WASM evidence.
Fresh Excel observation: **NOT_RUN**; the new checks use the pinned public
reference and synthetic regressions. tiny-vba remains on immutable revision
`8f4bebc3650cd51054fa6786f83a1ac73e07aeaf`.

Validation (exact commands in the OM20 specification):

- `cargo xtask ci` with the approved font bundle: all **10** gates passed,
  including **3,025** native tests.
- `cargo test -p gridcraft-vba`: **251** passed.
- WASM VBA suite via `wasm-bindgen-test-runner`: **242** passed.
- Optional audit-tool suite: **36** passed; affected-target strict Clippy passed.
- Offline audit write/check reproduced contracts, constants, reports and graph
  witnesses. Original source inventory and tiny-vba pin remain unchanged.
