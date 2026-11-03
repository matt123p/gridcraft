//! OM15 projects existing UI services through owned, cancelable host continuations.
use super::features::{Feature, absent, reads};
use super::objects::{arguments, numeric_index, required, string};
use super::{Bridge, HostCall, HostResult, InvokeKind, Object, Session, Value, failure, values};
use gridcraft_engine::model::sheet::SheetIdentity;
use serde_json::json;

/// Values documented inline on the pinned Excel.Range.CellControl page.
pub const CELL_CONTROL_CONSTANTS: &[(&str, i32)] = &[("xlTypeNone", 0), ("xlTypeUnknown", 1), ("xlTypeCheckbox", 2)];

/// Only existing native command dialogs; Excel's argument presets are not emulated.
pub const NATIVE_DIALOGS: &[(i32, &str)] =
    &[(47, "columnWidth"), (127, "rowHeight"), (256, "zoom"), (55, "insertCells"), (54, "deleteCells"), (61, "defineName"), (596, "insertLink")];
#[derive(Clone, Copy, Default, Debug)]
pub struct InteractionServices {
    pub dialogs: bool,
    pub open_picker: bool,
    pub save_picker: bool,
    pub status_visibility: bool,
}
#[derive(Clone, Debug)]
pub enum InteractionKind {
    Dialog { name: String },
    OpenFilename,
    FindFile,
    SaveFilename { suggested: String },
    StatusVisibility { value: Option<bool> },
}
#[derive(Clone, Debug)]
pub struct InteractionRequest {
    pub token: u64,
    pub doc: u64,
    pub window: u64,
    pub sheet: SheetIdentity,
    pub range_epoch: u64,
    pub selection: gridcraft_engine::Selection,
    pub kind: InteractionKind,
}
pub enum InteractionResponse {
    Path(Option<String>),
    Accepted(bool),
    Visible(bool),
}
impl InteractionKind {
    pub(crate) fn response(&self, response: InteractionResponse) -> HostResult<Value> {
        match (self, response) {
            (Self::OpenFilename | Self::FindFile | Self::SaveFilename { .. }, InteractionResponse::Path(path)) => match path {
                None => Ok(Value::Boolean(false)),
                Some(p) if p.len() <= 4096 && !p.is_empty() => values::vba_text(&p),
                _ => Err(failure(5, "picker path empty or exceeds 4096 bytes")),
            },
            (Self::Dialog { .. }, InteractionResponse::Accepted(v)) => Ok(Value::Boolean(v)),
            (Self::StatusVisibility { value: None }, InteractionResponse::Visible(v)) => Ok(Value::Boolean(v)),
            (Self::StatusVisibility { value: Some(v) }, InteractionResponse::Visible(actual)) if *v == actual => Ok(Value::Empty),
            _ => Err(failure(5, "interaction response does not match request")),
        }
    }
}
impl Bridge {
    pub(crate) fn interaction_api(&mut self, s: &mut Session, o: &Object, c: &HostCall) -> Option<HostResult<Value>> {
        let m = c.member.to_ascii_lowercase();
        if *o == Object::Application && m == "dialogs" && reads(c) {
            let f = Feature::new("Dialogs", None);
            return Some(if c.arguments.is_empty() { self.feature(f) } else { self.dialog_collection(&f, c) });
        }
        if *o == Object::Application && m == "inputbox" {
            return Some((|| {
                let a = arguments(c, &["Prompt", "Title", "Default", "Left", "Top", "HelpFile", "HelpContextID", "Type"])?;
                let _ = required(&a, 0)?;
                Err(failure(1004, "GridCraft has no general typed/range InputBox service; no dialog was opened"))
            })());
        }
        if *o == Object::Application && matches!(m.as_str(), "getopenfilename" | "getsaveasfilename" | "displaystatusbar" | "findfile") {
            return Some(Err(failure(1004, "this interaction requires a connected native host service; no UI was opened")));
        }
        if let Object::Range(si, r) = o {
            if m == "autocomplete" && reads(c) {
                return Some((|| {
                    let a = arguments(c, &["String"])?;
                    let text = string(required(&a, 0)?)?;
                    if text.len() > 32768 {
                        return Err(failure(7, "autocomplete input budget"));
                    }
                    let rect = r.rectangle()?;
                    if rect.count() != 1 {
                        return Err(failure(1004, "AutoComplete requires one cell in the native profile"));
                    }
                    self.scan_work = self.scan_work.saturating_add(10_000);
                    if self.scan_work > 1_000_000 {
                        return Err(failure(7, "macro autocomplete scan budget"));
                    }
                    let sh = s.doc().map_err(|e| failure(91, e.to_string()))?.wb.sheet(*si).ok_or_else(|| failure(91, "worksheet deleted"))?;
                    values::vba_text(&gridcraft_engine::column_completion(sh, rect.start, &text).unwrap_or_default())
                })());
            }
            if m == "cellcontrol" && reads(c) {
                return Some((|| {
                    arguments(c, &[])?;
                    let mut f = Feature::new("CellControl", Some(*si));
                    f.range = Some(r.clone());
                    self.feature(f)
                })());
            }
        }
        let Object::Feature(f) = o else {
            return None;
        };
        match f.kind {
            "Dialogs" => Some(self.dialog_collection(f, c)),
            "Dialog" => Some((|| {
                arguments(c, &[])?;
                if !reads(c) {
                    return Err(failure(438, "read-only Dialog"));
                }
                match m.as_str() {
                    "application" | "parent" => self.intern(Object::Application),
                    "show" => Err(failure(1004, "dialog service unavailable")),
                    _ => Err(failure(438, "unsupported Dialog member")),
                }
            })()),
            "CellControl" => Some(self.cell_control(s, f, c)),
            _ => None,
        }
    }
    fn dialog_collection(&mut self, _f: &Feature, c: &HostCall) -> HostResult<Value> {
        let mut all = Vec::with_capacity(NATIVE_DIALOGS.len());
        for (id, _) in NATIVE_DIALOGS {
            let mut f = Feature::new("Dialog", None);
            f.part = *id;
            all.push(f.object());
        }
        if c.kind == InvokeKind::Enumerate {
            arguments(c, &[])?;
            return self.object_array(all);
        }
        if !reads(c) {
            return Err(failure(438, "read-only Dialogs"));
        }
        match c.member.to_ascii_lowercase().as_str() {
            "count" => {
                arguments(c, &[])?;
                Ok(Value::Long(NATIVE_DIALOGS.len() as i32))
            }
            "application" | "parent" => {
                arguments(c, &[])?;
                self.intern(Object::Application)
            }
            "item" | "" | "dialogs" => {
                let a = arguments(c, &["Index"])?;
                let id = numeric_index(required(&a, 0)?)?;
                if !NATIVE_DIALOGS.iter().any(|(i, _)| *i == id) {
                    return Err(failure(1004, "this Excel dialog has no selected native projection"));
                }
                let mut f = Feature::new("Dialog", None);
                f.part = id;
                self.feature(f)
            }
            _ => Err(failure(438, "unsupported Dialogs member")),
        }
    }
    fn cell_control(&mut self, s: &mut Session, f: &Feature, c: &HostCall) -> HostResult<Value> {
        let si = f.sheet.ok_or_else(|| failure(91, "worksheet deleted"))?;
        let r = f.range.as_ref().ok_or_else(|| failure(91, "range deleted"))?;
        let rect = r.rectangle()?;
        if rect.count() > super::MAX_CELLS as u64 {
            return Err(failure(7, "cell-control range budget"));
        }
        let d = s.doc().map_err(|e| failure(91, e.to_string()))?;
        let sh = d.wb.sheet(si).ok_or_else(|| failure(91, "worksheet deleted"))?;
        arguments(c, &[])?;
        match c.member.to_ascii_lowercase().as_str() {
            "type" if reads(c) => {
                let mut found = None;
                for at in rect.iter() {
                    let t = if d.wb.styles.get(sh.style_id(at)).num_fmt.as_str() == "checkbox" { 2 } else { 0 };
                    if found.is_some_and(|f| f != t) {
                        return Ok(Value::Empty);
                    }
                    found = Some(t);
                }
                Ok(Value::Long(found.unwrap_or(0)))
            }
            "setcheckbox" if reads(c) => {
                if sh.is_protected() {
                    return Err(failure(1004, "sheet protection forbids cell-control formatting"));
                }
                let changed: Vec<_> = rect
                    .iter()
                    .filter(|at| sh.cell(*at).is_none_or(|v| v.value.is_empty() && v.formula.is_none()))
                    .map(gridcraft_engine::core::RangeRef::cell)
                    .collect();
                let events = if changed.is_empty() {
                    super::PreparedChanges::default()
                } else {
                    self.prepare_change(s, si, super::ranges::Range { areas: changed, unit: super::ranges::Unit::Cells })?
                };
                s.execute("insert.checkbox", json!({"sheet":si,"range":rect.a1()})).map_err(|e| failure(1004, e.to_string()))?;
                self.queue_changes(events);
                Ok(Value::Empty)
            }
            _ => Err(failure(438, "unsupported cell-control member")),
        }
    }
    pub(crate) fn defer_interaction(&self, s: &Session, o: &Object, c: &HostCall) -> HostResult<Option<InteractionRequest>> {
        let m = c.member.to_ascii_lowercase();
        let services = self.operations.interaction_services;
        let kind = if *o == Object::Application && m == "findfile" && reads(c) {
            arguments(c, &[])?;
            if !services.open_picker || cfg!(target_arch = "wasm32") {
                return Err(failure(1004, "FindFile requires a native filename picker and file-open service"));
            }
            InteractionKind::FindFile
        } else if *o == Object::Application && matches!(m.as_str(), "getopenfilename" | "getsaveasfilename") && reads(c) {
            let save = m == "getsaveasfilename";
            let a = arguments(
                c,
                if save {
                    &["InitialFilename", "FileFilter", "FilterIndex", "Title", "ButtonText"]
                } else {
                    &["FileFilter", "FilterIndex", "Title", "ButtonText", "MultiSelect"]
                },
            )?;
            if save {
                if a[1..].iter().any(|v| !absent(v)) {
                    return Err(failure(1004, "native save picker has no configurable filter/title/button service"));
                }
            } else {
                if a[..4].iter().any(|v| !absent(v)) || !absent(&a[4]) && super::address::boolean(&a[4])? {
                    return Err(failure(1004, "native open picker supports only default filter/title and single selection"));
                }
            }
            if save {
                if !services.save_picker {
                    return Err(failure(1004, "save filename picker unavailable"));
                }
                let suggested = if absent(&a[0]) { s.doc().map_err(|e| failure(91, e.to_string()))?.title.clone() } else { string(&a[0])? };
                if suggested.len() > 4096 {
                    return Err(failure(7, "suggested filename budget"));
                }
                InteractionKind::SaveFilename { suggested }
            } else {
                if !services.open_picker {
                    return Err(failure(1004, "open filename picker unavailable"));
                }
                InteractionKind::OpenFilename
            }
        } else if *o == Object::Application && m == "displaystatusbar" {
            let value = if c.kind == InvokeKind::Let {
                let a = arguments(c, &["Value"])?;
                Some(super::address::boolean(required(&a, 0)?)?)
            } else if reads(c) {
                arguments(c, &[])?;
                None
            } else {
                return Err(failure(438, "DisplayStatusBar requires Get/Let"));
            };
            if !services.status_visibility {
                return Err(failure(1004, "status-bar visibility service unavailable"));
            }
            InteractionKind::StatusVisibility { value }
        } else if matches!(o,Object::Feature(f) if f.kind=="Dialog") && m == "show" && reads(c) {
            let names = (1..=30).map(|i| format!("Arg{i}")).collect::<Vec<_>>();
            let a = arguments(c, &names.iter().map(String::as_str).collect::<Vec<_>>())?;
            if a.iter().any(|v| !absent(v)) {
                return Err(failure(1004, "Excel dialog argument presets have no native equivalent"));
            }
            let Object::Feature(f) = o else {
                return Err(failure(91, "dialog missing"));
            };
            let (_, name) = NATIVE_DIALOGS.iter().find(|(id, _)| *id == f.part).ok_or_else(|| failure(1004, "dialog unavailable"))?;
            if !services.dialogs {
                return Err(failure(1004, "native dialog service unavailable"));
            }
            InteractionKind::Dialog { name: (*name).into() }
        } else {
            return Ok(None);
        };
        let d = s.doc().map_err(|e| failure(91, e.to_string()))?;
        if d.selection.ranges.len() > 64 {
            return Err(failure(7, "interaction selection budget"));
        }
        if d.selection.ranges.is_empty()
            || !d.selection.active.is_valid()
            || !d.selection.anchor.is_valid()
            || d.selection.ranges.iter().any(|r| !r.start.is_valid() || !r.end.is_valid() || r.start.row > r.end.row || r.start.col > r.end.col)
        {
            return Err(failure(1004, "interaction selection is outside the worksheet"));
        }
        let sheet = d.wb.active().ok_or_else(|| failure(91, "worksheet missing"))?.identity.clone();
        Ok(Some(InteractionRequest {
            token: 0,
            doc: d.uid,
            window: s.active_window().ok_or_else(|| failure(91, "window missing"))?,
            sheet,
            range_epoch: d.range_epoch,
            selection: d.selection.clone(),
            kind,
        }))
    }
}
pub(crate) fn descriptors() -> Vec<super::MemberDescriptor> {
    let mut rows=vec![
        ("Application","FindFile","Function FindFile() As Boolean"),
        ("Application","Dialogs","Property Get Dialogs() As Dialogs"),
        ("Application","GetOpenFilename","Function GetOpenFilename(Optional FileFilter As Variant, Optional FilterIndex As Variant, Optional Title As Variant, Optional ButtonText As Variant, Optional MultiSelect As Variant) As Variant"),
        ("Application","GetSaveAsFilename","Function GetSaveAsFilename(Optional InitialFilename As Variant, Optional FileFilter As Variant, Optional FilterIndex As Variant, Optional Title As Variant, Optional ButtonText As Variant) As Variant"),
        ("Application","DisplayStatusBar","Property DisplayStatusBar As Boolean"),
        ("Dialogs","Count","Property Get Count() As Long"),("Dialogs","Item","Property Get Item(Index As Variant) As Dialog"),
        ("Dialogs","Parent","Property Get Parent() As Application"),("Dialogs","Application","Property Get Application() As Application"),
        ("Dialog","Parent","Property Get Parent() As Application"),("Dialog","Application","Property Get Application() As Application"),
        ("Range","AutoComplete","Function AutoComplete(String As String) As String"),("Range","CellControl","Property Get CellControl() As CellControl"),
        ("CellControl","Type","Property Get Type() As Variant"),("CellControl","SetCheckbox","Sub SetCheckbox()"),
    ].into_iter().map(|(owner,name,signature)|super::MemberDescriptor{owner,name,signature,availability:super::MemberAvailability::Implemented,implementation_symbol:"crates/vba/src/interaction_api.rs::Bridge::interaction_api"}).collect::<Vec<_>>();
    rows.push(super::MemberDescriptor{owner:"Dialog",name:"Show",signature:"Function Show(Optional Arg1 As Variant, Optional Arg2 As Variant, Optional Arg3 As Variant, Optional Arg4 As Variant, Optional Arg5 As Variant, Optional Arg6 As Variant, Optional Arg7 As Variant, Optional Arg8 As Variant, Optional Arg9 As Variant, Optional Arg10 As Variant, Optional Arg11 As Variant, Optional Arg12 As Variant, Optional Arg13 As Variant, Optional Arg14 As Variant, Optional Arg15 As Variant, Optional Arg16 As Variant, Optional Arg17 As Variant, Optional Arg18 As Variant, Optional Arg19 As Variant, Optional Arg20 As Variant, Optional Arg21 As Variant, Optional Arg22 As Variant, Optional Arg23 As Variant, Optional Arg24 As Variant, Optional Arg25 As Variant, Optional Arg26 As Variant, Optional Arg27 As Variant, Optional Arg28 As Variant, Optional Arg29 As Variant, Optional Arg30 As Variant) As Boolean",availability:super::MemberAvailability::Implemented,implementation_symbol:"crates/vba/src/interaction_api.rs::Bridge::defer_interaction"});
    rows.push(super::MemberDescriptor{owner:"Application",name:"InputBox",signature:"Function InputBox(Prompt As String, Optional Title As Variant, Optional Default As Variant, Optional Left As Variant, Optional Top As Variant, Optional HelpFile As Variant, Optional HelpContextID As Variant, Optional Type As Variant) As Variant",availability:super::MemberAvailability::ServiceRequired,implementation_symbol:"crates/vba/src/interaction_api.rs::Bridge::interaction_api"});
    rows
}
