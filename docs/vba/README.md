# VBA documentation

M7A.6 native-only files (2026-11-08): the shared custom form codec/APIs and
`.tvform` file routes/examples are removed under
[ADR 0099](../../../tiny-vba/docs/decisions/0099-native-only-form-workflows.md). New forms use
inert native templates. Native `.frm`/`.frx` Save/Export and edited macro-enabled
copy export retain the M7A.5 snapshot/signature/IO contract. Useful legacy-era
host callback tests now import native pairs; there is no legacy importer or
converter. Shared revision `d9431ce4992ee9580acfe2fe0f7276de229bedaf`. All 10 CI steps,
strict optional VBA/UI Clippy, six form UI cases, the VBA-enabled web build and
34 native/34 Node 22 WASM package cases pass. The migrated native import retains
its real workbook callback assertions. The host canvas was inspected; 13 shared
lockfile packages changed by revision only. M7A.7 acceptance and
M8.5 fresh Office comparisons remain separate; older codec/support entries below
are historical and superseded.

M7A.5 persistence (2026-11-08): native `.frm`/`.frx` Save/Export and
**Export Edited XLSM Copy** (`vba.exportNativePackage`) are implemented. The copy
compiles idle/reset authored source without execution, encodes edited native form
resources, removes invalidated VBA signatures and reopens the macro-enabled output
to check exact project bytes. File selection revalidates project/document/source/
design identity; native output commits atomically and refuses the current document
path. Copy export/download leaves the original editor modified and preserves its
original project bytes. Ordinary document Save retains those original VBA bytes;
open the edited copy to use the persisted changes. Native `.frm` Save acknowledges
only successful current-snapshot IO; browser download requests never do.
Shared revision `8800fcb07e628d4719a5c5459969503ba7e70f7c`.
See the [supported workflow and acceptance limits](../../../tiny-vba/docs/m7a5-acceptance.md).
All 10 CI steps, strict optional VBA/UI Clippy, six form UI cases, the VBA-enabled
web build and 34 native/34 Node 22 WASM package cases pass. The signature-removal
regression passes, including shared data retention and incomplete graph refusal.
Thirteen shared lockfile packages changed by revision only. M7A.4 is complete for the common profile under
ADR 0097; uncommon extensions belong to P9, `.tvform` removal to M7A.6 and final
integration to M7A.7. Fresh Office observations remain M8.5 NOT_RUN. Earlier
incremental entries below retain historical status and are superseded by this
current scope and persistence contract.

M7A.4 button click focus and input identity (2026-11-08): shared ADR 0096
implements CommandButton TakeFocusOnClick, native True defaults and False clicks
that preserve text focus/selection and deferred validation. Normal Tab/explicit
focus, authorized callbacks and Reset remain intact. Input checks canonical object
membership after ordinary writes settle and after peer index shifts; smaller
active-session slices are respected without raising limits. Package regressions
retain typed False at Stop and normal AutoTab to the button; host canvases retain
text focus. Shared revision `bcf3dc859095df7a331760d1469b8047c00cacf4`.
All 10 CI steps, strict optional VBA/UI Clippy, five form UI cases, the VBA-enabled
web build and all 33 native/33 Node 22 WASM package cases pass. The light host
canvas was inspected. Thirteen shared lockfile packages changed by revision only.
Full M7A.4 remains open; native edited saving/removal and fresh Office traces remain
M7A.5/.6 and M8.5 NOT_RUN.

M7A.4 hidden TextBox caret viewports (2026-11-08): shared ADR 0095 separates
scrolling from visible bars. None/room-hidden axes reveal canonical carets while
fixed authored bounds and translated pointer geometry remain intact. Caller-themed
accessible bars keep the existing authorized callback/default path. The host
canvas fixture now shows long text scrolled to its final row with ScrollBars=None.
No API, format, host authority or budget change. Shared revision `25a35a403e862bb8510bfe7117d95f379aeea1fc`. All 10 CI steps,
strict optional VBA/UI Clippy, five form UI cases, the VBA-enabled web build and
all 33 native/33 Node 22 WASM package cases pass. The light host canvas was
inspected. Thirteen shared lockfile packages changed by revision only. Full M7A.4, AutoSize, repeat/wheel/platform and geometry/font-change
edge cases remain open; fresh Office comparisons remain M8.5 NOT_RUN.

M7A.4 TextBox ScrollBars (2026-11-08): shared ADR 0094 implements native
ScrollBars import, ordinary Long None defaults/constants/errors and actual
bounded text viewports. Measured room, MultiLine and WordWrap determine axes;
track/drag waits for callbacks and preserves canonical text/ranges. Tagged UI
offsets reset with execution/control identity, follow caret changes and retain
scrolled pointer geometry. Native ListBox stored Both is accepted without a
public member. Package regressions check typed state at Stop and original
project preservation. Shared revision `05b8862abcf2d3204d26e1226acda372b89b6bb3`. All 10 CI steps, strict optional
VBA/UI Clippy, five form UI cases, the VBA-enabled web build and all 33 native/
33 Node 22 WASM package cases pass. The corrected light host canvas was
inspected. Thirteen shared lockfile packages changed by revision only. Full M7A.4,
AutoSize/hidden-bar/repeat/platform and container scrolling remain open.

M7A.4 SelectionMargin (2026-11-08): shared ADR 0093 implements native
TextBox/ComboBox bit-26 import, ordinary Boolean True defaults and a theme-local
left selection strip. Actual line selection/drag uses canonical UTF-16/CRLF
ranges after mouse callbacks finish, including Stop/Continue and changed policy.
Empty native edit fields render blank with accessible names. Package regressions
check typed state/ranges at Stop and original VBA project preservation. Canvas
fixtures compare the default TextBox margin with an editable ComboBox margin off.
Shared revision `295ae7adde51765e42698d392b5140a7a33c341f`. All 10 CI steps, strict optional VBA/UI Clippy, five
form UI cases, the VBA-enabled web build and all 33 native/33 Node 22 WASM
package cases pass. The light host canvas was inspected. Thirteen shared
lockfile packages changed by revision only. Full M7A.4 remains open; edited native
saving and fresh Office observations remain M7A.5 and M8.5 NOT_RUN.

M7A.4 AutoTab (2026-11-08): shared ADR 0092 implements native TextBox/ComboBox
bit-30 import, ordinary Boolean False defaults and accepted user-edit tab
traversal through existing validation/focus continuations. Actual edit cause
survives autocomplete with unchanged text; caret-only and programmatic changes
stay passive. Package coverage checks imported typed state, Stop/Continue,
authorized text edits, focus transfer and original project preservation.
Shared revision `e5a62681f074dab000894118fa71977e438e26df`. All 10 CI steps, strict optional VBA/UI Clippy, five
form UI cases, the VBA-enabled web build and all 33 native/33 Node 22 WASM
package cases pass. The light host canvas was inspected. Thirteen shared
lockfile packages changed by revision only. Full M7A.4 remains open; edited native
saving and fresh Office observations remain M7A.5 and M8.5 NOT_RUN.

M7A.4 HideSelection (2026-11-07): shared ADR 0091 implements native
TextBox/ComboBox bit-29 flags, ordinary Boolean state and True file/dynamic
defaults. Inactive selected text stays highlighted when False; both ComboBox
presentations use canonical ranges without changing edit permission or drop
state. Text/Value, selected ranges, events, original resources and all budgets
remain independent. Highlight painting uses local clipped galleys outside font
locks. Package regressions check trust refusal, typed state and nonempty ranges
at Stop/resume, and unchanged original VBA bytes. Canvas fixtures include a
runtime-created ComboBox with persistent highlighting.
Shared revision `74f89f2bfb80c461d9791175291b48d892fa6006`.
Full CI (all 10 steps), strict optional VBA/UI Clippy, five native-form UI
cases, the VBA-enabled web build and all 33 native/33 Node 22 WASM package
cases pass. The light host canvas was inspected. The 13 shared lockfile packages
changed by revision only. Remaining selection, properties,
AutoSize/scrolling, lifecycle/input, bindings and image/RefEdit/ActiveX/platform
cases stay open. Full M7A.4 remains IN PROGRESS; edited native saving is M7A.5
and fresh Office observations remain M8.5 NOT_RUN.

M7A.4 caption Alignment (2026-11-07): shared ADR 0090 implements native
CheckBox/OptionButton flag translation, ordinary Long Alignment state, Right
defaults and published constants. ToggleButton's documented disabled member is
not enabled. Real indicators and their bevels follow the selected side; caption
text stays left aligned and clipped in its own rectangle. Empty captions retain
accessible runtime names. Value/Null/groups/WordWrap/events, source/resources and
all budgets remain unchanged. Package regressions cover trust refusal, typed
properties, inert stopped inspection, one-time resume and original VBA bytes.
Canvas fixtures include a runtime-created left-caption CheckBox.
Shared revision `b8d203dd00bef56acce74deeb0d25d13811fbe1e`.
Full CI (all 10 steps), strict optional VBA/UI Clippy, five native-form UI
cases, the VBA-enabled web build and all 33 native/33 Node 22 WASM package
cases pass. The light host canvas was rendered and inspected. The 13 shared
lockfile packages changed by revision only. Remaining applicable members,
AutoSize/scrolling, lifecycle/input, binding and image/RefEdit/ActiveX/platform
cases remain open. Full M7A.4 stays IN PROGRESS; edited native saving is M7A.5
and fresh Office observations remain M8.5 NOT_RUN.

M7A.4 Boolean captions (2026-11-07): shared ADR 0088 implements native
CheckBox/OptionButton/ToggleButton WordWrap flags, True defaults and ordinary
Boolean state. Value, Null/TripleState, option groups, events and original
resources remain independent. Local caption layout retains widget input, full
accessible labels/state and authored bounds; ToggleButton captions stay centred.
Package regressions verify trust refusal, typed wrapping, stopped inspection,
one-time resume and unchanged original VBA bytes. Canvas fixtures include a
runtime-created wrapped CheckBox. ADR 0089 stabilizes the restricted WASI test
runner through value-preserving callback forwarding, with no stack/VM/capability
budget increase or added arguments, environment or preopened directories.
Shared revision `b7010bd2738bbcac87a018b94bd261c2d8533f24`.
Full CI (all 10 steps), strict optional VBA/UI Clippy, five native-form UI
cases, the VBA-enabled web build and all 33 native/33 Node 22 WASM package
cases pass. The light host canvas was rendered and inspected. All 13 shared
packages changed by revision only in the lockfile. Alignment/TextAlign and remaining
properties/defaults, AutoSize/scrolling, lifecycle/input, bindings and image/
RefEdit/ActiveX/platform cases remain open. Full M7A.4 stays IN PROGRESS; native
edited saving is M7A.5 and fresh Office observations remain M8.5 NOT_RUN.

M7A.4 caption WordWrap (2026-11-07): shared ADR 0087 decodes applicable
Label/CommandButton flags and exposes ordinary Boolean state with distinct file/
runtime defaults. Dynamic TextBox starts with WordWrap=True. Local caption
wrapping/clipping preserves fixed bounds, fonts and accelerators; unwrapped
multiline TextBox allocation retains its horizontal interaction/accessibility
bounds and complete UTF-16 text/cursor rows. Package cases cover typed wrapping
through trusted Stop/Continue, original resources, document edits/undo and denied
execution. Canvas fixtures exercise the same canonical properties.
Shared revision `2be3f2dcbf96b3798adb348649a131fe73436061`.
Full CI (all 10 steps), strict optional VBA/UI Clippy, five native-form UI
cases, VBA web and 33 native/33 actual Node 22-WASM package cases pass. The
updated document/form render was inspected. Boolean-control wrapping, AutoSize,
scrolling and the remaining property/lifecycle/input/binding/image/RefEdit/ActiveX
and platform cases stay open. Full M7A.4 remains IN PROGRESS; native edited saving
is M7A.5 and fresh Office observations remain M8.5 NOT_RUN. VM/host protocols,
global font/style policy and execution/memory/permission budgets are unchanged.

M7A.4 TextAlign (2026-11-07): shared ADR 0086 implements selected native
Label/TextBox/ListBox/ComboBox alignment with ordinary Long state and published
fmTextAlign constants. Native ParagraphAlign and public centre/right values are
translated explicitly. Font assignment, text/selection and events stay independent;
local glyph/caret positions agree, including UTF-16 pointer selection. ComboBox
popup rows keep their own alignment. Package regressions exercise typed alignment
through trusted Stop/Continue, retain document edits/undo and original VBA resources,
and keep denied execution inert. Canvas fixtures include aligned native controls.
Shared revision `77465858e83c5e59790296b5d2fb1ed3c3b39374`.
Full CI (all 10 steps), strict optional VBA/UI Clippy, five native-form UI
cases, VBA web, and 33 native/33 actual Node 22-WASM package cases pass. The
updated document/form render was inspected. Boolean/RefEdit and remaining property/default,
wrap/AutoSize, input/lifecycle/binding and ActiveX/platform cases stay open.
Full M7A.4 remains IN PROGRESS; native edited saving remains M7A.5 and fresh
Office observations remain M8.5 NOT_RUN. No VM/protocol/global style/font or
execution/memory/permission budget changes are included.

M7A.4 appearance audit (2026-11-07): shared ADRs 0083–0085 remove the
non-Office Page Font/colour exposure; Page children and client painting use the
owning MultiPage. Native records retain applicable BorderColor, BorderStyle,
SpecialEffect and BackStyle. Ordinary typed properties preserve documented
border/effect coupling, transparency and errors. Published Font aliases share the
canonical Font with exact Currency/Integer/String/Boolean values. Local painting
uses those fields and clips fixed-size controls to authored bounds. Package
cases cover Page ownership, typed appearance, trust refusal, Stop/Continue and
unchanged original VBA resources; canvas fixtures use aliases and appearance.
The follow-up list renderer removes its unconditional outside outline; canonical
BorderStyle/SpecialEffect own ListBox edges, while ComboBox popups keep their
separate frame. Canvas fixtures include runtime-created flat lists.
Shared revision `0bc169c5dd38f57634a77515a45a58370b1220b3`.
Full CI (10 steps), strict optional VBA/UI Clippy, five form UI cases,
VBA web and all 33 native/33 actual Node-WASM package cases pass. The host
render was inspected. Remaining member/default/selection, lifecycle/focus, binding,
keyboard/mouse, image/RefEdit and ActiveX/platform cases stay open. Full M7A.4
remains IN PROGRESS; native edited saving is M7A.5. Fresh Office observations
remain M8.5 NOT_RUN. Budgets and the existing macro bridge are unchanged.

M7A.4 Font-assignment update (2026-11-07): shared ADR 0082 adds ordinary typed
Font setters on native UserForms and Font-capable controls. Explicit assignment
shares an object reference; old aliases survive replacement and unload, and new
controls still copy independent container defaults. Root setters retain lazy
load/reload and unrelated objects fail nominal validation before mutation.
Package cases cover assignment across Stop/Continue, trust and original-resource
preservation; canvas input tests now use an assigned Font. Shared revision
`38dacbd178dd8f25c98b4da8338f575e85df8171` passes full CI (10 steps), strict
optional VBA/UI Clippy, five form UI cases, VBA web and 31 native/31 actual
Node-WASM package cases. The host render was inspected. Nothing/external/other-assignment, ambient/platform, appearance and the
full remaining M7A.4 checklist stay open; fresh Office comparisons are NOT_RUN.

M7A.4 Font update (2026-11-07): shared ADR 0081 retains native StdFont/TextProps
and VBFrame Font blocks with exact fractional sizes and original resources.
Ordinary Font objects expose typed attributes, independent instance state and
container defaults for runtime-added controls and Pages. Local rendering uses
canonical values and existing host fonts; optional face selection belongs to the
caller. Extreme display sizes are bounded with a visible notice while retaining
the canonical value. Host package cases cover trust refusal, typed aliases,
Stop/Continue and unchanged original VBA resources; canvas tests exercise authored
typography with existing activation/input ordering.
Pinned shared product `f899c9925de3e2d4ef838f8ac78287ec9b2f4c2b` passes full CI
(all 10 steps), strict optional VBA/UI Clippy, all five form UI cases, VBA web
and all 30 native/30 actual Node-WASM package cases. The host document render
was inspected. Font replacement, ambient/platform realization, borders/effects,
transparency and the remaining full M7A.4 checklist stay open. Fresh Office
comparisons remain M8.5 NOT_RUN.

M7A.4 colour update (2026-11-07): shared ADR 0080 retains native BackColor/ForeColor
bits and file defaults, exposes ordinary typed properties and paints authored RGB
values through the shared form renderer. Optional platform/client palette lookup
belongs to the caller's UI theme; missing client entries keep their canonical value
and show an unavailable-palette message. A runtime-created MultiPage Page-parent
lookup regression is fixed. Native package tests cover trust refusal, Stop/Continue,
one-shot document effects and unchanged original VBA resources. Host canvas tests
exercise authored form/control colours with existing activation/input ordering.
Pinned shared product `d3c99baab86fad9e33333fd4204066757baec63d` passes full CI
(10 steps), strict optional VBA/UI Clippy, all five form UI cases, VBA web and
29 native/29 actual Node-WASM package cases. The host document view was inspected.
Fonts, borders/effects, transparency, remaining ambient/platform/default/selection
and collection/lifecycle requirements stay open. Full M7A.4 remains incomplete;
fresh Office comparisons remain M8.5 NOT_RUN.

M7A.4 input-order update (2026-11-07): shared ADR 0079 orders earlier native
key/text/paste/IME input before document Deactivate. Existing keyboard receipts,
widget defaults and Change callbacks finish first; spanning composition stays
with the form through Stop. Reset/context/trust invalidation discards unapplied
input. Modal redirection cancels the later document gesture while preserving the
original keyboard receipt. No host API, authorization or budget is changed.
Canvas regressions now also pause in the preceding form Change callback.
Pinned shared product `231818a4179a6b096bcc66b1d8922ad0282ac7ce` passes full CI
(10 steps), strict optional VBA/UI Clippy, all five form UI cases, VBA web and
28 native/28 actual Node-WASM package cases. The canvas case covers stopped key
replacement/Change, Continue/Reset and one document replay; the host view was
inspected. Full M7A.4 remains incomplete; fresh Office comparisons stay M8.5 NOT_RUN.

M7A.4 document-focus update (2026-11-07): ADR 0078 adds a bounded raw-input
barrier before document canvas input. Native modeless Deactivate runs through
existing authorization; Stop retains the original gesture, Continue restores and
repaints the document before replay, and Reset/source/document/trust/geometry
changes discard stale input. Reentrant activation also cancels a partially
replayed pointer capture. RefEdit retains its explicit worksheet-selection path.
Pinned shared product `ce6482640332cb67ccfe3372886a4039c832437e` passes full CI (10 steps), strict
optional VBA/UI Clippy, all five form UI cases, VBA web and 28 native/28 actual
Node-WASM package cases. Broader M7A.4 remains incomplete; fresh
Office observations stay M8.5 NOT_RUN.

M7A.4 activation update (2026-11-06): shared ADR 0077 adds canonical native
Show/Hide/modeless activation, reentrant lifecycle guards and cross-window key
continuations. Window focus commands execute VBA through existing authorization;
external application focus and passive hover do not emit lifecycle callbacks.
The new package regression checks a stopped Deactivate handler before a one-shot
host edit and preserves original VBA resources. Automatic document-canvas raw
input buffering/replay remains unimplemented; L01/L02 and full M7A.4 stay open.
Pinned product and final validation are recorded in ROADMAP.md.

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

M7A.4 keyboard-focus update (2026-11-06): shared ADR 0071 adds root keyboard
focus, canonical Tab/SetFocus continuation and Space cancellation before button
or checkbox defaults. Native function keys are isolated from editor shortcuts.
Full CI (10 steps), strict optional VBA/UI Clippy, four form UI cases, web and
all 22 native/22 Node-WASM package cases pass on shared revision
`38578095ba1fc1ae3695f92a6c2e89cdb29d42f6`. Full M7A.4
remains incomplete.

M7A.4 accelerator update (2026-11-06): shared ADR 0072 at `7d43f082de4b7827da1e1d93c257ebeb50eca4ae`
adds persisted native accelerators, canonical Accelerator properties and selected
Label/button/boolean/Tab/Page Alt actions. Focus validation, Stop/Continue and
existing authorization apply. Caption underlines use the host theme and fonts.
Duplicate routing explicitly refuses; non-ASCII/platform ordering, mouse,
lifecycle and other checklist rows remain open. M7A.4 remains incomplete; fresh
Office observations remain M8.5 NOT_RUN. Full CI (10 steps), strict optional VBA/UI Clippy, four form UI cases, web and
all 23 native/23 Node-WASM package cases pass. The accelerator case preserves
authorization, one-shot debugger continuation, workbook effects and original
VBA resources.

M7A.4 mouse update (2026-11-06): shared ADR 0073 adds typed native
MouseDown/MouseMove/MouseUp/DblClick callbacks, indexed arguments, mutable cancel
and selected simple-control capture/defaults. Stop/Continue, outside release and
Reset preserve input identity and source. Event signatures omit unused class-field
metadata while preserving typed arguments, executable method fields and all limits.
Text/list/numeric/indexed-tab pointer defaults, lifecycle and other checklist rows
remain open; M7A.4 is incomplete and Office observations remain M8.5 NOT_RUN.
Pinned shared revision `c9f829824c56397c101ac9228c89db289b788458` passes full
CI (10 steps), strict optional VBA/UI Clippy, four form UI tests, web and all
24 native/24 Node-WASM package cases. Indexed MultiPage mouse input, stopped
callbacks, authority, workbook effects and exact original resources are verified.

Mouse command authorization is explicit under shared ADR 0073: form commands
require host preparation unless they are known inert design operations. The final
pin passes the same full CI/optional UI/web and 24 native/24 Node-WASM gates.

M7A.4 TextBox pointer update (2026-11-06): shared ADR 0074 adds selected
caret, captured drag, double-click cancellation/word selection and Shift extension
after ordinary mouse callbacks. Native AutoWordSelect is decoded with its inverse
file bit; new dynamic text controls retain their separate default. Canonical
UTF-16/CRLF, focus redirects, host authorization, Stop/Continue and Reset/source
identity remain enforced. Modal bounds stay stable while input waits. Editable
ComboBox/RefEdit, list/numeric/tab pointer defaults, further platform/selection and
lifecycle cases remain open; M7A.4 is incomplete and fresh Office comparisons
remain M8.5 NOT_RUN. Validation receipts follow below.

Accepted shared pin `3e8676cc754f59e06df64d2677c1f64af0216432`: full CI (10 steps), strict optional VBA/UI Clippy,
25 native and 25 actual Node-WASM package cases, four form UI cases and web checks
pass. The new selection case preserves authority, stopped continuation and exact
original workbook/VBA resources. M7A.4 remains incomplete.

M7A.4 composite-pointer update (2026-11-06): shared ADR 0075 adds selected
ComboBox/RefEdit text/drop/picker, list/popup row, numeric step/repeat/drag and
indexed tab/page-header defaults after ordinary callbacks. Retained modal display
metadata preserves window presentation while callbacks run, without authorizing
input. Tabs retain identity through reorder and discard removed targets. Existing
host consent, snapshots, ordinary VBA state and resource preservation remain in
force; ComboBox AutoWordSelect preserves its inverse native bit. M7A.4 and the
full runtime checklist remain incomplete. Fresh Office observations are M8.5
NOT_RUN; native edited persistence/removal remain separate.
Pinned shared product `f027c5d14f0a57c6a0ebfa0f1cc2fb393cb9c2b7` passes full CI (10 steps),
strict optional VBA/UI Clippy, four form UI cases, the VBA web build and all
26 native/26 actual Node-WASM package cases. The new modal pointer regression
preserves stopped authority, one-shot workbook effects and original resources.

M7A.4 nested-input update (2026-11-06): shared ADR 0076 services child mouse
callbacks while a parent mouse/key callback is parked. Raw input scopes share the
existing 64-event queue; held releases retain their original capture, modal areas
follow Show order, and open ComboBox popups stay above their owning form. Reset,
parent unload and owner-local keyboard barriers suppress stale/replayed defaults.
Canonical VBA state and existing host authority remain unchanged. Selected L02
cases are covered; M7A.4 and the full lifecycle/control checklist remain partial.
Fresh Office observations remain M8.5 NOT_RUN.
Pinned shared product `c99996351a80a234b23f2e81b54e210dae0bd530` passes full CI (10 steps),
strict optional VBA/UI Clippy, four form UI cases, VBA web and all 27 native/27
actual Node-WASM package cases. The nested callback regression preserves modal
ownership, separate child/parent receipts, one-shot workbook effects and original
resources.
