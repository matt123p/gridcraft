# Excel VBA object model v1 priority — interim assessment

Reviewed and accepted 2026-11-03. This is a product-priority assessment of the pinned
Microsoft Excel VBA object-model inventory, not a compatibility claim. It is the
pre-corpus judgment requested before measuring the repository test material.

## Accepted scope boundary

This integration exposes **only capabilities that GridCraft already supports**.
It may add VBA binding metadata, typed argument/result adapters, object identity,
dispatch, errors and tests around an existing GridCraft command, model or service.
It must not add a spreadsheet, UI, file, provider, rendering or collaboration
feature to GridCraft in order to satisfy an Excel VBA member.

Priority is evaluated only after this eligibility gate:

1. If the underlying behavior exists in GridCraft, the member may be High,
   Medium or Low for VBA projection.
2. If GridCraft does not currently support the behavior, the member is
   `OUT_OF_SCOPE_GRIDCRAFT`, irrespective of its Excel popularity or tier.
3. A future GridCraft feature may make a member eligible, but that product work
   is planned and delivered outside this VBA integration project.

The complete Microsoft inventory remains useful for compatibility diagnostics;
an out-of-scope row is not an unfinished VBA implementation requirement.

## Recommendation

GridCraft v1 should project the existing object model needed to read, transform,
calculate, format, validate, chart and save ordinary workbooks. It should not
treat every documented Excel automation API as equally valuable. The first
version should have an explicit portable profile and return a clear unsupported
member, unavailable service or policy error outside it.

The priority labels mean:

- **High — must have:** absence prevents ordinary workbook macros from using an
  already-supported GridCraft capability, or blocks access to a large amount of
  otherwise useful existing functionality.
- **Medium — nice to have:** materially useful and common in a narrower class of
  workbook, but not required for the basic editing/calculation workflow.
- **Low — unlikely for v1:** legacy interoperability, remote-provider automation,
  or specialist enterprise integration whose implementation cost and platform
  coupling are disproportionate to first-version value.

Enums and constants are not usefully prioritized in isolation. A constant has
the priority of the High/Medium/Low member that consumes it. Likewise,
`Application`, `Workbook`, `Worksheet` and `Range` are not all-or-nothing: their
core members are High while provider-, UI- or legacy-specific members follow the
owning capability below.

## High — must have

| Capability set | Principal owners | v1 judgment |
|---|---|---|
| Workbook graph and lifecycle | `Application`, `Workbook`, `Workbooks`, `Worksheet`, `Worksheets`, `Sheet`, `Sheets` | Essential roots, identity, active/owning context and adapters for lifecycle operations GridCraft already exposes; unsupported Excel lifecycle variants remain out of scope. |
| Cell and range model | `Range`, `Areas`, core collections | Essential addressing, values, formulas, selection, find/replace, copy/paste, fill, clear, structural edits and multi-area behavior. |
| Formula-facing APIs | `Range`, `Name`, `Names`, `Application` | Formula variants, names, evaluate/reference handling, arrays/spills and calculation control are needed by ordinary macros. |
| Worksheet functions | `WorksheetFunction` | High capability, low marginal implementation cost: expose one typed dynamic proxy to the existing function registry. Do not hand-implement 390 member wrappers. Preserve the documented difference between worksheet errors and raised VBA errors. |
| Everyday formatting | `Font`, `Interior`, `Borders`, `Border`, `Style`, `Styles`, basic `FormatConditions` | Required for macros that produce usable reports. Prioritize values that the renderer and file writer already preserve. |
| Common data operations | `Sort`, `SortFields`, `AutoFilter`, `Filter`, `Validation`, outline and basic data tools | Sorting, filtering and validation are normal workbook operations and already have substantial engine support. |
| Structured tables | `ListObject`, `ListObjects`, `ListRow`, `ListRows`, `ListColumn`, `ListColumns` | Tables are a modern core Excel structure and a dependency of many formulas and data macros. |
| Core events | `Application`, `Workbook`, `Worksheet` events | Implement only events tied to actions and precommit/cancellation points already represented by GridCraft. Do not add new engine lifecycle hooks solely for Excel event parity. |
| Execution controls and diagnostics | cross-cutting | Trust, bounded execution, cancellation, exact errors, undo grouping and compatibility reporting are release requirements rather than optional object families. |

The currently reconciled packet counts give a rough, deliberately conservative
upper bound of **2,045 High pages** across OM02–OM07, OM11 and OM18. This is not
2,045 independent implementations: collections, property families, generic
dispatch and `WorksheetFunction` account for many pages at once.

## Medium — nice to have

| Capability set | Principal owners | v1 judgment |
|---|---|---|
| Charts and drawing objects | `Chart`, `ChartObject`, `Series`, `Axis`, `Shape`, `Shapes`, formatting children | Valuable for reporting workbooks, but a coherent basic subset is more important than every chart/shape formatting leaf. |
| Pivot tables and slicers | `PivotTable`, `PivotCache`, `PivotField`, `PivotItem`, `Slicer` | High value for analytical workbooks but complex, stateful and not required for ordinary cell automation. Local pivot behavior should precede external OLAP/model providers. |
| Windows, panes and views | `Window`, `Windows`, `Pane`, `Panes`, view objects | Important for desktop macro fidelity; headless calculation/editing can work without full window parity. |
| Printing and export | `PageSetup`, page-break collections, print/export methods | PDF/export and page setup are useful; printer selection and preview are host services and can remain unavailable on some targets. |
| Protection and permissions | protection objects and range/sheet/workbook members | Important for authored templates and safe editing, but advanced encryption/protected-view policy is not needed to unlock core macro execution. |
| Dialogs, controls and user interaction | dialogs, `InputBox`, form/control objects | Eligible only where the matching GridCraft interaction/service already exists and can be safely adapted. Do not build new dialogs or controls for VBA parity. |
| Comments, notes, hyperlinks and metadata | comment/thread/person, hyperlink and document-property objects | Useful for review workflows, but secondary to values, formulas and formatting. Local notes/links should precede remote threaded identity. |

These packets currently contain **2,666 Medium pages**. The number is inflated by
the very broad chart/drawing and pivot object graphs; v1 should choose reachable
basic subsets rather than implement every formatting child merely because it has
a documentation page.

## Low — unlikely to be needed in v1

| Capability set | Principal owners | v1 judgment |
|---|---|---|
| External database/query connections | `QueryTable`, `WorkbookConnection`, `OLEDBConnection`, `ODBCConnection`, text/data-feed connections and provider parameters | Defer. These are external data-provider automation surfaces, not prerequisites for VBA workbook logic. Preserve imported package parts where practical, but do not execute or refresh them in v1. Local CSV/text import may be exposed through GridCraft's own file commands independently. |
| Collaboration, publishing and server APIs | sharing, mail, SharePoint, publishing, remote review and server objects | Out of scope while GridCraft lacks the corresponding provider, authentication and consent model. This integration does not create them and never reinterprets remote actions as no-ops. |
| Legacy/platform interoperability | DDE, COM/OLE Automation, XLL/add-ins, Excel4/XLM, `VBProject`/VBE, OS APIs and Excel Graph | Exclude from the portable v1 profile. Retain visible `SERVICE_REQUIRED`, `POLICY_BLOCKED` or `UNSUPPORTED_PROFILE` results rather than compatibility stubs. |
| External OLAP/data-model providers | provider-backed portions of pivot/cube/model APIs | Defer provider behavior even if local pivot tables are Medium. |

The directly assigned Low packets contain **1,013 pages**. Deferral removes a
large implementation burden without weakening the central promise that VBA can
automate a GridCraft workbook.

## Important object-specific judgments

### Range, Workbook and Worksheet

These are the centre of the v1 profile. Their eligible High subset should be
coherent, including ownership, collections, optional arguments, errors,
protection, undo, events and persistence already supported by GridCraft. A
catalogue entry alone is not enough and cannot authorize a new product feature.
Provider-specific or unrelated leaf members on these owners should still be
classified by capability; for example `Range.QueryTable` is Low even though
`Range.Value2` is High.

### Application

`Application` is a routing root, not a reason to reproduce the Excel desktop
process. High members include workbook/sheet roots, selection, calculation,
events, bounded `Run`, names/evaluation and practical settings used to suppress
painting or alerts. OS paths, DDE, COM add-ins, remote services and Excel UI
customization are Low or Medium according to the tables above.

### WorksheetFunction

Implement a single dynamic bridge:

1. Resolve a member name against GridCraft's existing formula-function registry.
2. Convert VBA scalars, arrays, errors and `Range` references into the formula
   engine's typed call context without flattening values to JSON.
3. Execute under the UDF/read-only and resource-budget policy.
4. Convert the result back to VBA values.
5. Raise the documented VBA error where `WorksheetFunction` differs from a cell
   formula or `Application.Evaluate`.

The documentation pages remain useful for aliases, arity and compatibility
reporting, but they should generate metadata/tests rather than handwritten Rust
methods.

## Provisional implementation order

1. Classify every candidate through the existing-GridCraft-capability gate.
2. Finish the reachable eligible High object graph: Application roots → Workbooks →
   Workbook → Worksheets/Sheets → Worksheet → Range/Areas.
3. Add the generic `WorksheetFunction`/Evaluate bridge for functions and formula
   behavior already implemented by GridCraft, and complete eligible formula/name
   projections.
4. Complete everyday formatting, data operations, structured tables and core
   events through existing engine commands.
5. Select coherent Medium projections only from existing chart, pivot, view,
   export/protection and annotation capabilities.
6. Publish Low and unsupported-GridCraft families as intentionally outside the
   v1 profile; preserve files
   containing their package parts where possible without executing them.

## Questions for corpus comparison

The final assessment should test these assumptions rather than merely count
member names:

- Which owner/member pairs occur in GridCraft-authored VBA integration tests?
- Which occur in reviewed MIT corpus source or expected-case manifests?
- Which occur only in unreviewed or rights-pending specimens?
- How many observed calls are covered by one generic family implementation?
- Do any Low families appear frequently enough to promote a narrow portable
  subset, without promoting the whole provider object graph?
- Which High owners are missing from the corpus, indicating a test-coverage gap
  rather than low product value?

