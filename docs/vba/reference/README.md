# VBA reference data

These files support the [feature reference](../supported.md),
[restriction reference](../unsupported.md) and [integration architecture](../integration.md).
Most readers can use those pages without opening this directory.

| File | Purpose |
|---|---|
| [Source inventory](vba-object-model-inventory.json) | Original pinned public Excel source paths and provenance; retained independently of current implementation |
| [Source contracts](vba-object-model-contracts.json) | Extracted syntax/arguments, member kind, native capability decisions, implementations and test links for each source page |
| [Source catalogue](vba-object-model-audit-report.md) | Searchable generated view of the contract rows |
| [Native profile](vba-object-model-profile.json) and [summary](vba-object-model-profile-report.md) | Reachable type/collection witnesses, exact required contracts and evidence boundaries |
| [Native feature evidence](vba-object-model-high-acceptance.json) | Core application/workbook/range/formula/data/formatting/table obligations and their regression tests |
| [Charts/drawings](vba-object-model-om12-acceptance.json), [pivots](vba-object-model-om13-acceptance.json), [windows](vba-object-model-om08-acceptance.json) | Native feature signatures, options, identity and persistence restrictions |
| [Protection](vba-object-model-om09-acceptance.json), [printing](vba-object-model-om10-acceptance.json), [annotations](vba-object-model-om19-acceptance.json), [interaction](vba-object-model-om15-acceptance.json) | Feature-specific contracts and host/service requirements |
| [Profile specification](vba-object-model-om20-acceptance.json) | Declared runtime scope, reviewed dynamic return types and dependency/source pins |
| [Unexposed capability records](vba-object-model-high-extension-review.json) and [summary](vba-object-model-high-extension-review.md) | Native backing, incomplete contracts and explicit exclusions for source members outside the declared adapter |

The historical `OM` labels in filenames are stable identifiers used by the
exporter; they are not instructions to implement every Excel page. The JSON
retains verification provenance without requiring milestone/process reports.

`PARTIAL` means a native subset or declaration has evidence, not that every Excel
option works. `MISSING` means an unexposed source contract. Eligibility/service
fields explain whether native backing exists. A deferred or unselected row is not
an implemented feature. Same-name methods/events and legacy Graph topics keep
separate source identities. No source-page total is an Excel compatibility score.

The [offline exporter](../integration.md#keeping-reference-data-current) owns the
generated files. Update input evidence when behavior changes, then regenerate and
check the output rather than hand-editing generated rows.
