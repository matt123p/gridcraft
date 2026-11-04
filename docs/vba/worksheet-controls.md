# Native worksheet Form Controls

Updated 2026-11-04. M7A.3 implements Button, CheckBox, OptionButton, DropDown,
ListBox, ScrollBar, Spinner, Label and GroupBox in the existing workbook Shape
store. Standard egui widgets supply pointer/keyboard focus. Assigned macros need
the optional VBA build and explicit workbook consent; opening a file stays inert.

Shapes.AddFormControl uses Excel's enum and point units. Shape.ControlFormat
exposes the selected native value, linked-cell/list, range, selection, list-edit
and enabled/print properties. The generated [supported index](supported.md) and
[source catalogue](reference/vba-object-model-contracts.json) give exact signatures.
The engine/control channel exposes control.add, control.set and control.activate;
the shell's vba.control command performs trusted assignment dispatch. Property
writes and import never execute assignments.

Controls use live local cell/range/name references, ordinary engine undo and
recalculation, and native OOXML/VML save/reopen. Current source parts are separate
from live state; conflicts remain inert, untouched bytes are retained and known
edits preserve unrelated XML. Application.Caller identifies the invoking control.
A paused/waiting/running macro cannot receive a second queued click.

This is a bounded Form Control profile. Worksheet ActiveX, label text binding,
editable/simple combo styles, arbitrary resources, broader visual styles, external
or multicolumn bindings and unlisted members remain unsupported. Numeric imports
need explicit native Max/Page defaults. Protected-sheet control edits/input refuse.
UserForm resource encoding and .tvform removal are separate shared milestones.

See the [full property/format/input/persistence contract and evidence](../../../tiny-vba/docs/m7a-worksheet-controls.md)
and [ADR 0041](../../../tiny-vba/docs/decisions/0041-worksheet-form-controls.md).
All fixtures are independently authored Rust builders; fresh Office comparisons
remain NOT_RUN under M8.5. No blanket Excel parity claim follows from this profile.
