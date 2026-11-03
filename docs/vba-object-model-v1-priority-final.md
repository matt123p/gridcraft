# Excel VBA object model v1 priority — final corpus-informed report

Reviewed and accepted 2026-11-03. This report turns the full 5,996-page Excel VBA inventory
into a first-version product profile. It follows the provisional judgment in
[the interim report](vba-object-model-v1-priority-interim.md), compares that
judgment with the available test corpora, and recommends a faster implementation
method.

## Executive decision

Do not implement every documented Excel member for v1.

The v1 promise should be: **VBA can reliably automate the workbook model that
GridCraft itself supports**. That requires a broad and coherent `Application` →
`Workbook` → `Worksheet` → `Range` path, formula/name/function access, everyday
formatting and data tools, structured tables, core events, exact errors, undo and
persistence. It does not require Excel's historical role as a COM automation
host, database front end, SharePoint client, XLL container or Windows desktop
shell.

Final priorities:

- **High — must have:** core workbook/range graph, formula and function bridge,
  common formatting/data/table operations, core events and execution controls.
- **Medium — selected subsets after High:** coherent basic chart/pivot/view/print/
  additional protection/interaction/annotation subsets.
- **Low — deferred; no implementation at this time:** connection-provider
  automation, remote publishing/collaboration and legacy/platform interoperability.
  This also applies to Low members inside otherwise High/Medium families.

This is a portable GridCraft profile, not a statement that low-priority Excel
features are unimportant to every Excel user.

## Binding scope rule

The VBA project is an integration layer, not a vehicle for expanding GridCraft.
A documented member is eligible only when its underlying behavior is already
available through a GridCraft command, model, calculator function, file path,
renderer or host service. VBA work may add typed bindings, object projection,
dispatch, diagnostics and integration tests; it may not add the underlying
spreadsheet or application capability.

Apply this gate before the priority tier:

- `ELIGIBLE_EXISTING`: GridCraft already supports the behavior; prioritize its
  VBA projection using the High/Medium/Low findings below.
- `OUT_OF_SCOPE_GRIDCRAFT`: GridCraft does not support the behavior today. Keep
  the row visible in compatibility output, but do not implement it in this
  project and do not count it as unfinished VBA work.
- `POLICY_BLOCKED` or `SERVICE_UNAVAILABLE`: the behavior conflicts with policy
  or depends on a GridCraft service that does not exist on the active host.

If GridCraft later gains a capability through its own roadmap, its VBA member
can be reclassified in a separate reviewed inventory change. Priority never
overrides eligibility. Eligibility does not override deferral either: an eligible
Low member remains deferred. Keep `DEFERRED_LOW` as a planning disposition separate
from `OUT_OF_SCOPE_GRIDCRAFT`, service availability and runtime support evidence.

## Complete family classification

This table classifies every implementation packet and therefore routes every
inventory page. OM IDs are stable family identifiers, not delivery sequence
numbers. The implementation plan's [priority-based delivery order](vba-object-model-implementation-plan.md#work-packets-and-delivery-order)
is authoritative. OM00 is metadata used by all tiers; enums inherit the tier of
their consumer rather than receiving an artificial standalone priority.

| Inventory family | Pages | Priority | v1 treatment |
|---|---:|---|---|
| OM00 — inventory, signatures, enums/constants | 272 | Cross-cutting | Keep complete metadata; generate only runtime constants required by reachable members, while retaining all documented rows in compatibility reports. |
| OM01 — host execution, registry and diagnostics | — | Accepted foundation | Maintain bounded execution and shared descriptors with each slice. |
| OM02 — selection, clipboard, editing and ranges | 68 | **High** | Already accepted for the bounded profile; extend only where corpus-backed High dependencies require it. |
| OM03 — Application state, calculation and dispatch | 206 | **High** | Project roots, `Run`, calculation and settings that already have real GridCraft state. Unsupported Excel application settings are out of scope, not requests to add host state. |
| OM04 — workbook and worksheet lifecycle | 278 | **High** | Implement Workbooks, open/save/close, sheet lifecycle, visibility/index and stale/cross-workbook ownership. |
| OM07 — names, formulas, Evaluate and WorksheetFunction | 515 | **High** | Project existing names/formula behavior and add one generic typed proxy over registered functions. Do not hand-code hundreds of function methods or add missing mathematics. |
| OM05 — formatting and conditional formatting | 541 | **High** | Implement common geometry, font/interior/borders, alignment, merge, number formats and a useful conditional-format subset through renderer/file state. |
| OM06 — sorting, filtering, validation and data tools | 180 | **High** | Project existing sort/filter/validation and common transformations first; specialist analysis behavior absent from GridCraft is out of scope. |
| OM18 — structured tables | 134 | **High** | Project GridCraft's existing table graph and structured references; connection-backed variants are Low/deferred with OM14, and independently out of scope where providers are absent. |
| OM11 — events | 123 | **High by owning feature** | Deliver High-operation events with their lifecycle/mutation slices. Medium events accompany selected Medium features; Low events remain deferred. Missing GridCraft hooks stay out of scope. |
| OM12 — charts, drawings and images | 1,220 | **Medium** | Project reachable charts/shapes and formatting already supported by GridCraft; omit unsupported Excel leaf behavior. |
| OM13 — pivots, slicers and data models | 838 | **Medium** | Project existing local pivot behavior only. Missing slicer, OLAP, cube or model features are outside this integration. |
| OM08 — windows, panes, views and navigation | 163 | **Medium** | Project the coherent desktop/view subset already represented by GridCraft; absent window/view behavior stays out of scope. |
| OM09 — protection and permissions | 96 | **Medium default; ordinary protection High** | Enforce existing protection on every mutation. Project ordinary Protect/Unprotect/status with High workbook/sheet work; take additional eligible permission/protection projections as selected Medium subsets. |
| OM10 — page setup, printing and export | 133 | **Medium** | Adapt existing deterministic page setup, PDF/export and print services only; missing options or services are not added by VBA work. |
| OM19 — notes/comments, hyperlinks and metadata | 101 | **Medium** | Local notes, links and properties first; remote identity/service-backed behavior remains Low/deferred, and out of scope where services are absent. |
| OM15 — dialogs, interaction and built-in controls | 115 | **Medium** | Adapt only interactions backed by an existing GridCraft service and explicit headless behavior; do not create new dialogs or controls. |
| OM14 — connections, query/import/XML and refresh | 305 | **Low — deferred** | Classification/diagnostics only; no new runtime projection, including eligible local variants. Existing file/package fidelity remains with GridCraft. |
| OM16 — sharing, collaboration and publishing | 138 | **Low — deferred** | No runtime projection now, even if a backing service exists. Missing providers/identity/authentication remain independently out of scope. |
| OM17 — legacy/Graph/OS/COM/XLL/VBE interoperability | 570 | **Low — deferred** | No new runtime projection/platform layer now, including portable equivalents. Preserve provenance and explicit refusal/policy/service results. |
| OM20 — reconciliation and profile acceptance | — | Cross-cutting | Run throughout; sign off after High and selected Medium subsets. Deferred Low implementation is not required. |

The packet-level page totals are **2,045 High**, **2,666 Medium**, **1,013 Low**,
plus 272 cross-cutting OM00 pages. They are pre-eligibility upper bounds: rows
whose behavior GridCraft lacks become `OUT_OF_SCOPE_GRIDCRAFT`. These totals
retain family defaults; they are not exact member-level tiers (notably OM09's
ordinary protection exception and events following their owning feature). OM01
and OM20 add no source pages. Page counts
measure documentation breadth, not Rust work items. Generic collections, shared
properties, generated metadata and the function proxy collapse many pages into
one implementation.

## Object-level rules within mixed owners

Large owners cross priority boundaries. Classify the member by what it does,
not merely by the word before the dot.

### Application

High members include eligible workbook/sheet roots, selection, calculation,
events, `Run`, `Evaluate`, `WorksheetFunction` and practical environment/path
values. Settings such as `ScreenUpdating`, `DisplayAlerts` or scheduling qualify
only if matching GridCraft state/services already exist; otherwise they are
`OUT_OF_SCOPE_GRIDCRAFT`.
Dialogs and window interaction are Medium. DDE, COM add-ins, XLLs, VBE,
SharePoint/mail/publishing and OS process control are Low.

### Workbook and Worksheet

Identity, collections, lifecycle, activation, visibility, names, paths, tables,
normal protection and supported events are High. Window/print/dialog and local
annotation projections are Medium. Provider connections, remote collaboration,
VBE and legacy automation are Low. A High owner does not promote all of its
provider-specific children.

### Range

Values, formulas, addresses, traversal, selection, editing, copy/paste, fill,
find/replace, structural operations, ordinary formatting, validation, sort/
filter, names and arrays/spills are High. Printing, annotations and UI-only
navigation are Medium. `QueryTable`, linked-data providers, OLAP-specific
children, subscription/collaboration and platform legacy members are Low.

### WorksheetFunction

`WorksheetFunction` is High, but it is an adapter—not 390 separate algorithms.
The implementation should:

1. expose `Application.WorksheetFunction` as one reachable proxy object;
2. resolve the invoked member through the existing GridCraft function registry;
3. use generated documentation metadata for aliases and argument shape;
4. pass typed VBA scalars, arrays, errors and range-backed values into a bounded,
   read-only formula context;
5. convert the result back without JSON/f64 flattening; and
6. translate worksheet errors into the documented raised VBA error behavior.

Unsupported calculator functions are outside the VBA integration scope and
remain explicit function-registry gaps. They do not require algorithms or empty
object-model stubs here.

## Corpus comparison

### Evidence sets and safety

No downloaded macro was executed.

The comparison used:

- 10 GridCraft VBA adapter integration-test files;
- three reviewed reflex-grid VBA case manifests;
- 16 MIT vba-project-starter `.bas`/`.cls` files;
- the MIT VBA-JSON source as a useful negative control;
- GridCraft/tiny-vba authored examples;
- the tree-sitter Excel member-access example;
- static source from 12 of 29 available macro-enabled corpus workbooks, decoded
  with tiny-vba's bounded read-only `vba-project` extractor; and
- package-part names from 27 XLSM workbooks to detect charts, pivots, tables,
  controls and connections without reading provider credentials or running code.

Seventeen XLSM projects were not source-scanned because the strict extractor
refused designer/UserForm data or inconsistent/unsupported project metadata.
That is a limitation of the evidence, not evidence that those workbooks use no
object-model members. The broad tree-sitter third-party source set was measured
separately because it mixes Excel, Access, web, Windows and library projects; it
is useful for capability discovery but not a representative GridCraft v1 demand
sample. Rights-pending teaching workbooks were used only for local static counts;
no source is copied into this report or the product.

The scanner was receiver-aware where declarations or roots made the receiver
clear, but it is still static lexical evidence. Late-bound objects, aliases,
`With` inference and inaccessible VBA projects make the counts lower bounds.

### Aggregate result

Across the relevant authored/reviewed/static sets, the scan inferred **181
distinct documented owner/member pairs**:

| Proposed tier | Distinct observed pairs | Share |
|---|---:|---:|
| High | 163 | 90.1% |
| Medium | 15 | 8.3% |
| Low | 3 | 1.7% |

The packet concentration was equally clear: OM02 60 pairs, OM04 30, OM05 28,
OM07 22, OM06 15 and OM03 8. Those six High families account for 163 observed
pairs before counting core-event and structured-table reachability that lexical
receiver inference did not fully capture.

Current status among the 181 observed pairs was 83 `implemented-subset`, one
`context-only`, 87 `not-exposed`, eight `formula-bridge-not-exposed`, and two
explicit platform-service results. This is a prioritization sample, not a
compatibility percentage; occurrence does not prove all optional arguments or
behavior.

### Strongest evidence

- Range/cell usage dominates every relevant Excel-oriented set. The reflex-grid
  cases alone contained 773 `Range` tokens across all three case files.
- Workbook/sheet ownership appears throughout the adapter tests, MIT project
  source and teaching workbooks. `Application.Workbooks`, worksheet visibility,
  workbook paths and worksheet creation are repeated unsupported gaps.
- Formatting is not cosmetic edge behavior: reflex-grid exercises Font,
  Interior, Borders, NumberFormat and conditional formatting extensively.
- Sort/filter/validation and tables appear in reviewed case/source sets, backing
  their High classification even though they are less frequent than Range.
- Basic charts deserve early Medium treatment: chart parts appeared in four of
  27 statically inspected XLSM packages, and extracted source referenced chart
  objects in teaching workbooks.
- Pivot support remains Medium: one inspected workbook contained pivot parts,
  with no comparable repeated source-level evidence in the successfully decoded
  projects.

### WorksheetFunction result

Corpus evidence supports the generic-proxy design. Eight distinct documented
functions were observed across three relevant sets:

`CountA`, `CountBlank`, `VLookup`, `Sum`, `RoundDown`, `Transpose`, `Trim` and
`Proper`.

This variety is enough to require scalar/range/array/error correctness, but far
too small to justify hundreds of handwritten wrappers. One proxy plus generated
metadata gives broader coverage with less code and a smaller defect surface.

### Connections and legacy APIs

No OM14 connection member pair was inferred in the relevant successfully
extracted source sets. Connection/query package parts did occur in specialist
workbooks, notably a database-read sample and a real-time-currency sample. This
changes the recommendation only in one respect:

- **package preservation and an explicit unavailable-refresh result matter;**
- **the database/query VBA automation object graph remains Low for v1.**

The broad third-party syntax corpus contained ADODB/DAO/Recordset and connection
tokens, but those were concentrated in mixed Access/web/platform projects. It
also contained substantial VBE, `CreateObject`, OS and add-in usage. Implementing
those dependencies would turn GridCraft into a Windows/COM compatibility project
and is not justified by the first-version workbook goal.

`Workbook.VBProject` appeared in two MIT project-starter files. It remains Low
and policy-sensitive: observing self-inspection or self-modification does not
make VBE automation a portable workbook requirement.

## Corpus-backed eligible High implementation queue

The highest-value missing VBA projections are listed below. Each individual
member must first pass the existing-capability gate; wording here does not
authorize missing GridCraft behavior. This list ranks value clusters, not strict
whole-packet prerequisites; use the delivery stages in the implementation plan.

1. **Workbook roots and lifecycle:** `Application.Workbooks`, `Workbooks.Open`,
   eligible close/path/visibility/index/add/copy/delete operations and collection
   semantics already supported by GridCraft.
2. **Practical Application state:** `Run`, calculation and environment values,
   plus only those settings with existing GridCraft state. `ScreenUpdating` and
   `DisplayAlerts` are excluded unless independently implemented in GridCraft.
3. **Formula/name bridge:** `Application.WorksheetFunction`, the generic function
   proxy, `FormulaR1C1`, `Text`, `HasFormula`, array formula forms, workbook/range
   names and Evaluate variants.
4. **Everyday formatting:** Borders, ColumnWidth/RowHeight/AutoFit, Hidden,
   alignment/wrap/indent, merge state, ColorIndex/Pattern and reachable
   FormatConditions.
5. **Data tools:** Range/Worksheet Sort, SortFields, AutoFilter/ShowAllData,
   Validation and the common RemoveDuplicates/TextToColumns paths.
6. **Structured tables and events:** reachable ListObjects graph, structured
   ranges/references, and events tied to the supported mutations/lifecycle.

Basic ChartObject/Chart access is the first Medium slice after those clusters
because GridCraft already has chart models and commands. Existing local pivot
behavior follows after its workbook/table/sort/filter foundations; unsupported
pivot or slicer behavior does not enter this project.

## Architecture review amendment (2026-11-03)

The [implementation plan's Rust architecture review](vba-object-model-implementation-plan.md#architecture-review-and-low-friction-delivery-2026-11-03)
refines the proposals below. Keep the current optional adapter crate and shared
object registry. Extract private family modules as they are touched; generate
reviewed metadata incrementally through the existing Rust exporter. A new kernel
DSL, whole-dispatch rewrite, corpus tool or full-inventory migration is not a
prerequisite to useful projections. Small host-neutral engine entry points may
expose existing behavior without expanding GridCraft's product capabilities.

Select synchronous reachable slices first, and review their eligibility just
before implementation. Whole packet dependencies are not barriers. Separate
waiting lifecycle services from ordinary properties/collections. The
[completed event remediation](vba-object-model-implementation-plan.md#event-remediation-2026-11-03)
fixes target-document event ownership, complete event-batch reservation and
missing/duplicate assignment notifications; preserve those regressions while
extending cross-document mutations. Python is suitable for offline inventory analysis;
maintained generation and validation stay in Rust. The remaining module/generation recommendations are planned changes; the linked
event remediation is implemented.

## A faster implementation model

The current documentation-led audit is valuable, but a member-at-a-time dispatch
ladder will be too slow. The faster model is **metadata-generated families plus
small handwritten semantic kernels**.

### 1. Add a versioned profile field to the inventory

Record `v1_priority`, `v1_disposition` and `implementation_family` per canonical
row. Suggested dispositions:

- `REQUIRED`
- `REQUIRED_GENERIC_PROXY`
- `MEDIUM_SUBSET`
- `SERVICE_REQUIRED`
- `DEFERRED_LOW` (including eligible Low rows; no runtime work now)
- `DEFERRED_PROVIDER`
- `POLICY_BLOCKED`
- `LEGACY_OUT_OF_PROFILE`
- `OUT_OF_SCOPE_GRIDCRAFT`

Generate the High/Medium/Low report from these fields so the product profile and
compatibility report cannot drift apart. The reports in this change are the
design input; the inventory migration should be a separate reviewed mechanical
change.

### 2. Generate checked-in Rust metadata with Rust tooling

Extend `xtask vba-audit` to generate checked-in Rust/JSON from the pinned
contracts. Python can remain a disposable analysis aid, but it should not enter
the build. Generated data should cover:

- canonical owner/member and aliases;
- invocation kind and property access;
- argument names, optionality and documented defaults;
- result owner/type;
- enum domains;
- implementation family and availability; and
- generated negative binding cases.

Generation should fail on duplicate/unassigned descriptors and stale outputs,
as the current audit already does for its contracts.

### 3. Replace per-member plumbing with implementation families

Most pages fit a small number of kernels:

- root/parent/application identity;
- indexed/default/enumerable collection;
- typed state property;
- range transform returning another Range;
- engine command mutation/query;
- formula-function proxy;
- format projection over a range;
- file/window/dialog service call;
- explicit unavailable provider/policy result.

A descriptor selects a kernel and supplies typed parameters. Handwritten Rust is
then reserved for semantic operations such as copy/paste, sort, lifecycle,
formula evaluation and events. This reduces boilerplate without turning subtle
Excel behavior into generic success.

### 4. Reuse engine command schemas

Where an engine command already owns validation, undo, recalc, protection and
persistence, map VBA metadata to that command rather than reimplementing the
operation. Add a small typed argument/result adapter and a preflight hook for VBA
events/errors. Generate straightforward mappings; hand-review commands whose VBA
defaults or event footprints differ. If no underlying GridCraft capability
exists, classify the member out of scope; do not add an engine command here.

### 5. Build the object graph reachability-first

Implement parents/roots and collection identity before leaf methods. A perfect
`PivotField` method is valueless if no macro can obtain a stable PivotField.
For each vertical slice, require:

`root → collection → item → child/property → mutation/query → save/reopen`.

This naturally prioritizes Workbooks, Worksheets, Range, Names, tables and basic
charts and avoids isolated catalogue stubs.

### 6. Turn corpus scanning into a repeatable planning tool

Add an optional, non-CI-by-default `xtask vba-usage` that accepts explicitly
reviewed source directories or read-only extracted project sources and emits
only aggregate owner/member counts. It must never execute macros or fetch
corpora. Join the result to inventory IDs to produce:

- observed unsupported pairs;
- observed service/policy dependencies;
- High members with no test evidence; and
- corpus scenarios to minimize into original regression tests.

The disposable scan used for this report demonstrates the value, but its loose
lexical inference should be replaced by tiny-vba's parser/binder plus the shared
member descriptors before becoming a maintained tool.

### 7. Generate tests, not semantic claims

For simple descriptor families, generate tests for named/positional/omitted
arguments, invocation kind, wrong receiver, invalid enums, stale objects and
unsupported availability. Keep authored integration tests for state changes,
events, undo, recalc, protection, persistence and service behavior. Generated
tests accelerate breadth; they cannot certify Excel parity by themselves.

## Recommended delivery sequence

Follow the implementation plan's [delivery stages and packet index](vba-object-model-implementation-plan.md#work-packets-and-delivery-order).
Keep OM IDs unchanged; numeric order is only a catalogue lookup aid.

1. Retain accepted OM00/OM01/OM02 and event-remediation foundations. Classify the
   selected slice's eligibility and engine evidence just before implementation.
2. **High:** establish the needed synchronous OM03/OM04 roots/collections, then
   OM07's generic WorksheetFunction proxy and eligible formula/name/Evaluate
   behavior; proceed through common OM05 formatting → OM06 data tools → OM18
   structured tables by actual per-slice prerequisites.
3. **High, interleaved:** deliver OM11 events with each relevant operation,
   existing protection enforcement throughout, and OM09 ordinary protection
   projections under the member-level High rule. Complete remaining eligible
   OM03/OM04 lifecycle/Run slices as their wait/invocation contracts are ready,
   before High sign-off. No whole-file-lifecycle barrier delays simple properties.
4. **Medium, after High:** basic OM12 charts/drawings → OM13 local pivots, then
   selected OM08 views → remaining OM09 protection/permissions → OM10 print/export
   → OM19 local annotations → OM15 interaction/control subsets. A narrow helper
   needed by a High operation may accompany it; broader Medium work waits.
5. **Low, deferred:** OM14/OM16/OM17 and Low members in all other packets receive
   classification/diagnostics only. Existing runtime support and package fidelity
   are retained; no new Low runtime projection is scheduled. Moving a Low feature
   to another packet does not make it High.
6. **OM20 throughout and at sign-off:** publish completion against High and the
   explicitly selected Medium subsets. Retain deferred Low and unselected Medium
   rows visibly; do not require their implementation to accept this profile or
   describe them as completed. Full source-page reconciliation remains required.

The member's eligibility, priority and required profile are separate decisions.
Out-of-scope behavior, eligible-but-deferred Low behavior, and partial implemented
support must remain distinguishable in planning and compatibility output. The
per-row priority/disposition migration is still planned; this documentation
revision changes delivery instructions, not the generated audit schema or counts.

## Revision history

- 2026-11-03: Aligned OM delivery to High-first/selected-Medium stages, explicitly deferred Low implementation, and clarified event/protection member exceptions without renumbering packets.

- 2026-11-03: Recorded completed event targeting, batch reservation and assignment notification remediation.

- 2026-11-03: Added Rust architecture review amendment and incremental delivery constraints; retained corpus priority findings.
