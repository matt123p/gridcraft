# Unsupported VBA functions and features

M7A.3 (2026-11-04): the nine standard worksheet Form Controls now have an
[implemented bounded profile](worksheet-controls.md). Worksheet ActiveX, label
text binding, editable combo styles, external/multicolumn bindings, broader native
styling and unlisted control members remain outside it. Protected controls refuse
edits/input. Selected native UserForm import remains separate; broader UserForm
semantics, edited resources and .tvform removal are still M7A.4–M7A.7.

Updated 2026-11-06: edited-source binary VBA project writeback and cryptographic
signature verification/re-signing remain unsupported. Save Source exports raw
source separately; document save retains the original binary project. Replacing
a captured binary project while reusing its original signature/data graph refuses.
Package-wide signatures are invalidated by package rewrite. Native worksheet ActiveX remains unsupported; native UserForm import has
selected control families and runtime packets; full UserForm behavior remains
in progress. See the
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
| Forms and controls | Worksheet ActiveX, UserForm resources/control types outside the declared ten-family flat import/runtime profile, arbitrary Office forms-control graphs, WordArt and SmartArt remain unsupported |
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

M7A.1/M7A.2 revision — 2026-11-04: selected native UserForm import/attachment
through shared designer/runtime/debugger; source/resource import remains inert.
Native edited form saving remains unsupported.

M7A.3 revision — 2026-11-04: worksheet Form Controls no longer require a future
interactive-support stage. Profile limits and independent authored evidence remain
explicit; no ActiveX, UserForm encoder or fresh Office acceptance is implied.

M7A.4 increment — 2026-11-04: selected native boolean controls, instance routing,
properties and lifecycle are implemented. Collection enumeration/type binding,
full focus/validation and list/container/scroll/image/binding/platform packets
remain unsupported; native edited saving and .tvform removal stay later tasks.

M7A.4 runtime follow-up — 2026-11-04: selected native SpinButton/ScrollBar and
TextBox selection/input behavior are implemented. Further ComboBox/ListBox array-shape/loading cases, multiple selection/bindings, Frame/pages/
tabs, Image/RefEdit, broader properties/bindings/input and worksheet ActiveX remain
unsupported or incomplete; full M7A.4 acceptance is still open.

M7A.4 list array follow-up — 2026-11-04: selected 2D List/Column loading and
indexed/current-row Column access are implemented. Further array-shape parity, empty/1D
loading, column-vector assignment, further coercion/list/binding behavior and the
broader milestone remain incomplete. Native edited save/.tvform removal remain
later M7A tasks.

M7A.4 readback/metadata follow-up — 2026-11-04: selected array readback and native
Tag/ControlTipText/HelpContextID are implemented. Help-file association/F1 routing,
further list/binding semantics, containers/pages/images/RefEdit and applicable
platform controls remain incomplete, along with native edited persistence and
.tvform removal. No full M7A sign-off is claimed.

M7A.4 native Frame follow-up (2026-11-05): the shared runtime now supports
selected nested native Frame storages, immediate child Controls/Parent identity,
scoped tab/option groups, nested ActiveControl and cancellable ancestor focus.
Shared widgets use live relative geometry, clipping and accessible group ancestry.
The package regression exercises typed Frame access and trusted child text input,
then verifies exact original project resources after save/reopen. This does not
implement native edited-source/resource saving. MultiPage/TabStrip/Image/RefEdit,
further properties/bindings/platform/lifecycle semantics and .tvform removal
remain open; fresh Office observations remain NOT_RUN under M8.5.

M7A.4 native TabStrip follow-up — 2026-11-05: selected native records now have
separate Tab/Tabs identity, indexed/named lookup, enumeration, Value/SelectedItem,
metadata and authorized pointer/arrow input in the shared runtime. Nearby controls
stay on their original form/Frame. Dynamic tab mutation, writable Index/Name,
further presentation/events and the remaining M7A.4 scope remain open. Native
edited saving and .tvform removal remain separate milestones.

M7A.4 native MultiPage/Page follow-up — 2026-11-05: shared revision
`18e412caa8d9dfd1a1ba984ea49c42bc490ef8da` adds bounded native f/o/x joins,
canonical Page/Pages identity and child Controls/Parent, indexed/named lookup,
enumeration, Value/SelectedItem and selected-page input/focus/rendering. Internal
tabs remain private. Physical selection respects dirty-field cancellation and
Stop/resume. TabStrip/MultiPage Click now carries its Long index and selecting
the current page/tab raises no selection event. Native source/resources remain
inert and separate; original project bytes survive package save/reopen. Dynamic
Pages/Tabs mutation, further page/Frame/property/resource/event/lifecycle cases,
Image/RefEdit, bindings and applicable platform controls remain M7A.4 requirements.
Native edited saving/.tvform removal remain M7A.5/.6; no full milestone sign-off
or fresh Office observations are claimed.

M7A.4 native Image/picture follow-up — 2026-11-05: shared revision
`4455df4082430d3d0d56412df73defd92c26c1df` adds bounded inert BMP/GIF/JPEG/ICO
resources, ordinary Image/stdole picture identity and typed assignment/clearing,
clip/stretch/zoom, alignment/tiling and selected AutoSize/border/effect state.
Image clicks preserve dirty text focus and ordinary callback/Stop continuations.
Canonical state stays in VM instances, with immutable shared pixel assets and
bounded derivative texture caches. Original source/resources remain separate.
Vector/platform picture services and the full image event/presentation matrix,
RefEdit, bindings, dynamic collections, platform controls and remaining lifecycle/
property semantics stay in M7A.4. Native edited saving/.tvform removal remain
M7A.5/.6; fresh Office observations remain NOT_RUN under M8.5.

M7A.4 native RefEdit follow-up — 2026-11-05: the shared implementation adds
bounded inert class-table/CLSID resolution, ordinary RefEdit reference metadata,
canonical text/drop state and selected focus/validation/Change/DropButtonClick
continuations. Original source/resources remain separate. The Excel host owns
worksheet range selection; no DLL/OCX/COM activation is introduced. Broader
RefEdit/platform/property/event semantics, bindings, dynamic collections and the
remaining M7A.4 ledger stay open. Native edited saving/removal stay M7A.5/.6;
fresh Office observations remain NOT_RUN under M8.5. Integration checks PASS;
see the RefEdit validation record in [integration](integration.md).

M7A.4 native Tabs/Pages mutation — 2026-11-05 (shared revision
`422c5c2338a05276e353b27018a531bf23231af4`): native collections support
Add/Remove/Clear, writable Integer Index and runtime-created member Name.
Reordering preserves aliases and selected identity; collection changes commit
before ordinary Change/Stop continuations. Added empty Pages and reordered
native children use the shared renderer; removed Page children reject widget
input. Original project/source/resource bytes remain separate. Fresh Office
comparison remains NOT_RUN under M8.5; Controls.Add/Remove, multiple selection,
bindings and the remaining M7A.4 semantic ledger stay open. Native edited saving
and .tvform removal remain M7A.5/.6. Integration validation PASS: all 13 native and 13 actual Node/WASM package
cases, full ten-step CI, strict optional VBA/UI Clippy, three form widget cases
and the optional VBA web check. The new case checks both Tabs and Pages,
authorization, retained aliases and unchanged XLSM VBA project bytes.


M7A.4 runtime controls/selection/bindings — 2026-11-05: selected native runtime
Controls mutation, multiple selection and single-cell/row bindings are now
implemented. Bound rows refuse manual mutation; RowSource currently supports one
contiguous range of at most 1,024 rows and ten columns. Nonempty sources require
an explicit embedding Range capability. Broader binding value/update timing,
property/event/lifecycle/platform cases remain open. The approved static
permission table holds at most 4,096 sites; executed host calls remain limited to
1,024 per session. Native edited saving/.tvform removal and fresh Office evidence
remain separate requirements; this does not close M7A.4.


Runtime type inspection update (2026-11-06): foreign object TypeName no longer
always falls back to Object when the embedding supplies validated native type
metadata. VarType now reads one compatible zero-argument default via normal
authorized execution (Range.Value and Application.Name in this host). Unadapted
objects retain type 9; additional default/arity cases need M8.5 comparisons.
Further M7A.4 event/lifecycle/property/
platform semantics and fresh M8.5 comparisons remain open.


As New array update (2026-11-06): the generic project-class/native-form array
restriction is removed under ADR 0068/IR 14. External activation and object-module
Public arrays retain their existing refusals. Further alias/reentrant mutation
comparisons and broader typed-event/lifecycle/control/platform work remain open.


Empty-project collection update (2026-11-06): UserForms access no longer
requires a form component. Remaining collection iteration/mutation and lifecycle
cases are tracked in the shared M7A.4 completion checklist; the milestone is open.

Native typed keyboard callbacks are partially implemented under shared ADR 0070.
Remaining physical key/platform mappings, control-character shortcuts, root-window
focus, IME/locale/focus transitions and accelerators remain open, along with mouse
events and the other M7A.4 checklist rows. This does not complete M7A.4.

UTF-16 string storage does not supply an ANSI code page. Asc/Chr and connected
ANSI KeyPress handlers support ASCII in this host context; non-ASCII ANSI
conversion requires an explicit configured code page. Unicode text without a
connected KeyPress handler retains its existing text input path.

ADR 0071 implements selected root keyboard focus, Tab/SetFocus and non-text Space
ordering. Cross-window activation/lifecycle, other keyboard/platform/IME cases,
mouse, remaining accelerator cases and the remaining native control contracts stay open.

ADR 0072 implements selected native Label/button/boolean/Tab/Page accelerators.
Duplicate matches are explicitly unsupported; non-ASCII/platform shortcut and
further key/focus/event-order cases remain open. This does not close K03 or M7A.4.

ADR 0073 implements typed mouse callbacks and selected simple-control capture/
default processing. Text/list/numeric/drop/picker and indexed Tab/Page defaults,
further capture/geometry/window/event-order and platform cases remain open.
M01 and M7A.4 stay partial; event-signature metadata compaction raises no limits.

ADR 0074 implements selected native TextBox caret/drag/word pointer selection.
Editable ComboBox/RefEdit, list/numeric/drop/picker/indexed-tab pointer defaults,
selection margins, drag/drop, auto-scroll, further platform/word/focus cases and
all other open shared M7A.4 checklist rows remain unsupported or partial.

Shared ADR 0075 adds selected composite/list/numeric/tab-header pointer defaults.
Further mouse/platform ordering, text/list drag/wheel and mutation/binding cases,
tab/page background hits, lifecycle/focus/keyboard/IME and the remaining property,
resource and host checklist remain open. M7A.4 is not fully accepted.

ADR 0076 covers selected nested modal raw-input continuations and stale release
suppression. Full Activate/Deactivate, document/modeless focus, reentrant binding
and other remaining lifecycle/platform cases still require implementation and
acceptance. This does not finish L02 or M7A.4.
