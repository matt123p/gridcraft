# VBA documentation

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
