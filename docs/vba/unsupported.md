# Unsupported VBA functions and features

M7A status (2026-11-04): native worksheet Button metadata decoding has started.
Interactive worksheet Form Controls, native `.frm`/`.frx` or embedded UserForm
import, edited native form persistence and `.tvform` removal are still open.
The metadata reader does not authorize or execute macros.
[Native forms ledger](../../../tiny-vba/docs/m7a-inventory.md).

Updated 2026-11-04: edited-source binary VBA project writeback and cryptographic
signature verification/re-signing remain unsupported. Save Source exports raw
source separately; document save retains the original binary project. Replacing
a captured binary project while reusing its original signature/data graph refuses.
Package-wide signatures are invalidated by package rewrite. Native worksheet
Form Controls and binary UserForm import remain deferred. See the
[writeback/signature contract](../../../tiny-vba/docs/decisions/0037-m7-package-preservation-and-writeback.md).

GridCraft exposes its existing workbook capabilities through a bounded VBA
adapter. Some Excel features have no native equivalent; others exist in
GridCraft but do not yet have a VBA binding. This reference distinguishes those
cases from host-dependent services and restrictions on supported members.

Use the [supported reference](supported.md) for available functions and features.
The [source catalogue](reference/vba-object-model-contracts.json) contains exact
source/member rows and capability-review reasons. A source row marked MISSING is
not, by itself, proof that GridCraft lacks the underlying feature.

## Runtime functions that need a service

The shared runtime recognizes more function names than this embedding supplies
services for. A recognized name can therefore refuse at runtime. GridCraft
configures invariant/Gregorian/UTF-16 context and binary string comparison; it
does not configure the following ambient services.

| Functions or mode | Why unavailable or restricted |
|---|---|
| `Date`, `Date$`, `Time`, `Time$`, `Now`, `Timer`; unseeded `Randomize` | No runtime clock service is injected; the interpreter does not read the operating-system clock implicitly |
| `Format`, `Format$`, `FormatCurrency`, `FormatDateTime`, `FormatNumber`, `FormatPercent`, `Error`, `Error$` | Require a locale/calendar formatting or localized error-description provider |
| `DateValue`, `TimeValue`, `MonthName`, `WeekdayName`; localized string-to-date/number conversions | Require date parsing, locale names or numeric parsing services; numeric Gregorian date operations have a separate supported path |
| Text comparison in `StrComp`, `InStr`, `InStrRev`, `Replace`, `Split` or `Filter`, including `Option Compare Text` | The configured comparison provider supports Binary mode; locale Text collation is not injected |
| `MsgBox`, `InputBox`, `InputBox$`, `DoEvents` | No generic language-intrinsic UI/scheduler service is connected; native object-model dialogs are a different API |
| `Command`, `Command$`, `Environ`, `Environ$` | Ambient process arguments and environment variables are not exposed |
| `Dir`, `Dir$`, `CurDir`, `CurDir$`, `EOF`, `FileAttr`, `FileDateTime`, `FileLen`, `FreeFile`, `GetAttr`, `Input`, `Input$`, `Loc`, `LOF`, `Seek` | No general VBA filesystem, file-handle or directory-enumeration service is injected; Workbooks.Open/Save use the bounded native file service instead |
| `GetSetting`, `GetAllSettings` | No operating-system settings/registry provider is injected |
| `Shell`, `CreateObject`, `GetObject`, `IMEStatus`, `MacID`, `MacScript` | Restricted process, COM, platform or scripting facilities; no ambient Office/OS automation |

Byte-string and `StrConv` operations are limited to the configured representation
and conversion modes. Code-page-dependent characters or locale-specific
conversions may be refused. `CDec` does not invent floating-point/Date rounding;
text conversion requires an appropriate provider. A numeric value supported by
`CDate` or `IsNumeric` does not imply localized text parsing is available.

## Worksheet functions

| Functions | Reason |
|---|---|
| `WEBSERVICE`, `FILTERXML` | Calculator registrations deliberately return errors; there is no network-fetch or XPath service |
| `CUBEKPIMEMBER`, `CUBEMEMBER`, `CUBEMEMBERPROPERTY`, `CUBERANKEDMEMBER`, `CUBESET`, `CUBESETCOUNT`, `CUBEVALUE` | No cube/provider backend; registering a formula name is not an external-data implementation |
| `WorksheetFunction.RTD` | No real-time data server/COM provider |
| Excel function names outside the calculator's registered set | No formula implementation to call; the VBA adapter does not supply replacement mathematics |

The [supported reference](supported.md#additional-functions-through-generic-dispatch)
also identifies callable functions that lack individual typed member metadata.
That is a binding/catalogue limitation, distinct from an unimplemented calculator
function. VBA intrinsics and worksheet functions with similar names are not
interchangeable in arguments, coercion or error behavior.

## Excel features without a native backend

| Area | Examples and reason |
|---|---|
| COM, add-ins and platform interoperability | COM/XLL/VBE automation, arbitrary DLL/type-library loading, DDE/OLE links and legacy Excel Graph graphs have no native host service |
| External data and data models | Provider-backed queries/connections/refresh, Power Query, XML maps/import/export and OLAP/model graphs are not supplied by local tables or pivots |
| Sharing and Office services | Coauthoring, legacy shared-workbook change tracking, workflows, publishing, mail, cloud users and external identity are unavailable |
| Office security services | IRM, resolved user ACLs, Protected View, independent write reservation, digital signing and encryption writing have no matching backend; code consent is a separate policy |
| Recovery and converter graphs | Excel AutoRecover scheduling and plugin FileConverter objects are not native file Open/Save or GridCraft's host AutoSave preference |
| Forms and controls | ActiveX, Office binary UserForm resources, arbitrary Office forms-control graphs, WordArt and SmartArt are not provided by native cell checkboxes/plain-source forms |
| Advanced rendering/state | Editable gradient-stop collections, arbitrary chart effects, indexed Office palettes and independent settings not represented in the native model cannot be added as inert VBA flags |
| Search and proofing options | Persistent Excel FindFormat/ReplaceFormat criteria, AutoCorrect graphs and full Office dictionary/options behavior are broader than native search/basic spelling |

## Existing features with incomplete VBA projection

The following have related native capabilities, but their full Excel contracts
are unavailable. These are not additional supported members merely because a
similar GridCraft operation exists.

| Area | Current boundary |
|---|---|
| Sheet state and preferences | Examples include Worksheet.Next/Previous/Tab/StandardHeight/Type, Application.DisplayFormulaBar/UserName/SheetsInNewWorkbook and unit-conversion helpers: native backing exists, but these members lack declared adapter bindings |
| Range and border extensions | Range.Text/BorderAround/InsertIndent and Border.ThemeColor/TintAndShade have native backing but no declared projection |
| Application/history/settings | Undo/Repeat registration, custom lists, caller contexts and process-wide preferences need exact scope/ownership; stored workbook state is not a matching Excel application setting |
| Phonetics and rich text | XLSX phonetic preservation and whole-text operations do not provide a complete editable per-cell phonetic/rich-run graph or reading generator |
| Effective formatting | Rendered conditional styles do not yet provide a complete read-only DisplayFormat graph |
| Sparklines and table styles | Native sparklines and style selection do not imply complete SparklineGroups/axes/points or editable TableStyleElements collections |
| Error and audit objects | Cell errors and precedent/dependent queries do not implement all Excel Error/Errors categories, ignored-error flags or persistent tracer-arrow navigation |
| Native host preferences | AutoSave, animations, recent files and window geometry need an explicit host seam and scope; a UI preference is not automatically a per-workbook VBA property |

The detailed [capability records](reference/vba-object-model-high-extension-review.json)
retain the precise native evidence and limits for these source contracts.

## Restrictions on supported features

| Feature | Unsupported variants and reason |
|---|---|
| Source persistence | Source edits do not rewrite `vbaProject.bin`. Standard modules use GridCraft storage; class/form projects must be exported separately where workbook storage refuses them |
| Project invocation | No arbitrary unchanged-XLSM compatibility promise, automatic execution on open, cross-project/add-in Run or automatic trust of another workbook |
| Sheet transfers | Cross-workbook/new-workbook transfers and copies requiring document-module, table or preserved-package transfer services refuse explicitly |
| Clipboard | No system clipboard, cross-workbook or multi-area clipboard graph; unsupported linked/theme/conditional-format-merge paste modes refuse |
| Calculation | No inferred selective Range/Worksheet.Calculate or Excel row-major-order contract from native whole-workbook calculation |
| Protection | UI-only protection, dormant selection settings before Protect and unsupported rights cannot be emulated by consent flags |
| Charts/drawings | Chart sheets, blank-chart creation, arbitrary new series, persistent shape groups/connectors, drawing-object clipboard and chart-only export need missing native services |
| Pivots | Standalone/shared external caches, provider connection lifecycles, arbitrary model calculations, unsupported aggregation and missing slicer persistence are not implied by local report output |
| Windows | Independent Excel panes, split/freeze/view/gridline/heading/zoom/scroll-position properties and several geometry queries are not all represented as per-window native state |
| PDF/printing | XPS, unsupported quality/metadata/viewer/page-interval options, range/group preview, header pictures and independent chart print configuration are unavailable |
| Annotations | Arbitrary custom-property editing, cloud identities, external hyperlink launch and non-cell hyperlink graphs require missing services |
| Events | Only the listed connected native signals are delivered. Additional Excel events, event polling substitutes and automatic Workbook_Open execution are unavailable |
| Worksheet UDFs | No workbook mutation, prompts, asynchronous waits, cross-workbook/add-in registration or arbitrary interior omitted arguments; built-in function names retain precedence |

## Host-dependent features

Native file pickers, save decisions, printing, print preview, dialogs and picture
file loading use services explicitly supplied by the host. A headless host can
use supported explicit paths/arguments where available; it cannot display a
picker or silently claim that printing succeeded. Browser builds use connected
upload/download/print callbacks and refuse unavailable filesystem operations.
`Application.InputBox`, `Range.PrintPreview` and unsupported window-geometry
services are not made available merely by enabling the editor.

## Resource and value limits

| Resource | Boundary |
|---|---|
| Object handles | 8,192 interned handles |
| Disjoint range areas | 64 |
| Materialized range | 65,536 cells per read/write, with an 8 MiB value payload budget |
| Search | 65,536 stored cells per End/Find scan; cumulative scan/wildcard budgets also apply |
| Macro host calls | 12,000, cumulative across debugger pauses/resumes |
| Event chains | 256 queued/delivered events; queue/identity capacity is checked before supported VBA mutations |
| Worksheet UDF | 100,000 instructions, depth 64 and 8 MiB payload per call; one million instructions per calculation pass |
| Applied source/functions | 64 modules, 4 MiB encoded source storage, 256 Functions, 64 UDF arguments; compiled-size limits also apply |
| Native print snapshots | At most 2,000 pages and 1,026 break markers per axis; snapshots expire when native revisions change |

Invalid/stale/foreign handles, deleted owners, protection failures, unavailable
services and exhausted budgets return explicit errors rather than success-shaped
placeholders. Unknown object members normally raise VBA error 438; wrong object
types raise 13, stale references 91, native range/option failures commonly 1004,
and consent failures 70. A handled error can be caught in VBA. Earlier successful
macro writes remain committed and grouped for undo; a later error does not
roll them back automatically.

M7A revision — 2026-11-04: initial native worksheet metadata decoder and explicit
remaining native form/control gaps; no interactive control or edited-save claim.
