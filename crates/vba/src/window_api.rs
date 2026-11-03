//! Projection of native window identities and selections. Sheet view settings stay sheet-owned.
use super::features::{Feature, reads};
use super::objects::{arguments, numeric_index, required, string};
use super::{Bridge, HostCall, HostResult, InvokeKind, Object, Session, Value, failure, ranges, values};
use gridcraft_engine::core::RangeRef;
use serde_json::json;
fn window(id: u64) -> Feature {
    let mut f = Feature::new("Window", None);
    f.key = id;
    f
}
impl Bridge {
    pub(crate) fn window_api(&mut self, s: &mut Session, o: &Object, c: &HostCall) -> Option<HostResult<Value>> {
        let m = c.member.to_ascii_lowercase();
        if matches!(o, Object::Application | Object::Workbook) {
            if m == "windows" && reads(c) {
                let mut f = Feature::new("Windows", None);
                f.part = i32::from(*o == Object::Application);
                return Some(if c.arguments.is_empty() { self.feature(f) } else { self.window_dispatch(s, &f, c) });
            }
            if *o == Object::Workbook && m == "newwindow" && reads(c) {
                return Some((|| {
                    arguments(c, &[])?;
                    if s.windows().len() >= super::MAX_OBJECTS {
                        return Err(failure(7, "window budget"));
                    }
                    self.objects.reserve(1)?;
                    s.execute("view.newWindow", json!({})).map_err(|e| failure(1004, e.to_string()))?;
                    self.feature(window(s.active_window().ok_or_else(|| failure(91, "window closed"))?))
                })());
            }
            if *o == Object::Application && m == "activewindow" && reads(c) {
                return Some((|| {
                    arguments(c, &[])?;
                    match s.active_window() {
                        Some(id) => self.feature(window(id)),
                        None => Ok(Value::Nothing),
                    }
                })());
            }
        }
        let Object::Feature(f) = o else {
            return None;
        };
        if !matches!(f.kind, "Window" | "Windows" | "WindowSheets") {
            return None;
        }
        Some(self.window_dispatch(s, f, c))
    }
    fn window_dispatch(&mut self, s: &mut Session, f: &Feature, c: &HostCall) -> HostResult<Value> {
        let m = c.member.to_ascii_lowercase();
        let doc = s.doc().map_err(|e| failure(91, e.to_string()))?.uid;
        if f.kind == "Windows" {
            let all = s.windows().iter().filter(|w| f.part == 1 || w.doc_uid == doc).collect::<Vec<_>>();
            if all.len() > super::MAX_OBJECTS {
                return Err(failure(7, "window collection budget"));
            }
            if c.kind == InvokeKind::Enumerate {
                arguments(c, &[])?;
                let objects = all.iter().map(|w| Object::Document(w.doc_uid, Box::new(window(w.id).object()))).collect();
                return self.object_array(objects);
            }
            if !reads(c) {
                return Err(failure(438, "read-only window collection"));
            }
            return match m.as_str() {
                "count" => {
                    arguments(c, &[])?;
                    Ok(Value::Long(all.len() as i32))
                }
                "application" => {
                    arguments(c, &[])?;
                    self.intern(Object::Application)
                }
                "parent" => {
                    arguments(c, &[])?;
                    self.intern(if f.part == 1 { Object::Application } else { Object::Workbook })
                }
                "item" | "" | "windows" => {
                    let a = arguments(c, &["Index"])?;
                    let v = required(&a, 0)?;
                    let w = if matches!(v, Value::String(_)) {
                        let title = string(v)?;
                        all.iter().find(|w| s.window_title(w.id).is_some_and(|t| t.eq_ignore_ascii_case(&title))).copied()
                    } else {
                        usize::try_from(numeric_index(v)?.checked_sub(1).ok_or_else(|| failure(9, "window index"))?)
                            .ok()
                            .and_then(|i| all.get(i).copied())
                    }
                    .ok_or_else(|| failure(9, "window not found"))?;
                    self.intern(Object::Document(w.doc_uid, Box::new(window(w.id).object())))
                }
                _ => Err(failure(438, "window collection member needs an unavailable window service")),
            };
        }
        let w = s.window(f.key).filter(|w| w.doc_uid == doc).ok_or_else(|| failure(91, "window closed"))?;
        if f.kind == "WindowSheets" {
            let d = s.doc().map_err(|e| failure(91, e.to_string()))?;
            let selected = &d.selected_sheets;
            if selected.len() > super::MAX_OBJECTS {
                return Err(failure(7, "selected sheets budget"));
            }
            if c.kind == InvokeKind::Enumerate {
                arguments(c, &[])?;
                return self.object_array(selected.iter().map(|si| Object::Sheet(*si)).collect());
            }
            if !reads(c) {
                return Err(failure(438, "selected sheets are read-only"));
            }
            return match m.as_str() {
                "count" => {
                    arguments(c, &[])?;
                    Ok(Value::Long(selected.len() as i32))
                }
                "parent" => {
                    arguments(c, &[])?;
                    self.intern(Object::Workbook)
                }
                "application" => {
                    arguments(c, &[])?;
                    self.intern(Object::Application)
                }
                "item" | "" | "selectedsheets" => {
                    let a = arguments(c, &["Index"])?;
                    let v = required(&a, 0)?;
                    let si = if matches!(v, Value::String(_)) {
                        let name = string(v)?;
                        selected.iter().copied().find(|si| d.wb.sheet(*si).is_some_and(|sh| sh.name.eq_ignore_ascii_case(&name)))
                    } else {
                        usize::try_from(numeric_index(v)?.checked_sub(1).ok_or_else(|| failure(9, "selected sheet index"))?)
                            .ok()
                            .and_then(|i| selected.get(i).copied())
                    }
                    .ok_or_else(|| failure(9, "selected sheet not found"))?;
                    self.intern(Object::Sheet(si))
                }
                _ => Err(failure(438, "unsupported selected sheets member")),
            };
        }
        if m == "newwindow" && reads(c) {
            arguments(c, &[])?;
            if s.windows().len() >= super::MAX_OBJECTS {
                return Err(failure(7, "window budget"));
            }
            self.objects.reserve(1)?;
            s.execute("view.newWindow", json!({})).map_err(|e| failure(1004, e.to_string()))?;
            return self.feature(window(s.active_window().ok_or_else(|| failure(91, "window closed"))?));
        }
        if matches!(m.as_str(), "activatenext" | "activateprevious") && c.kind == InvokeKind::Call {
            arguments(c, &[])?;
            if s.windows().len() > super::MAX_OBJECTS {
                return Err(failure(7, "window budget"));
            }
            let visible = s.windows().iter().filter(|w| !w.hidden).collect::<Vec<_>>();
            let at = visible.iter().position(|w| w.id == f.key).ok_or_else(|| failure(1004, "window is hidden"))?;
            let next = if m == "activatenext" { (at + 1) % visible.len() } else { (at + visible.len() - 1) % visible.len() };
            let w = visible.get(next).ok_or_else(|| failure(91, "window closed"))?;
            let target = s.document(w.doc_uid).ok_or_else(|| failure(91, "workbook closed"))?;
            if !target.trust.allows_code() && !(target.uid == self.doc && self.consent) {
                return Err(failure(70, "target workbook requires consent"));
            }
            s.execute("window.activate", json!({"window":w.id})).map_err(|e| failure(1004, e.to_string()))?;
            return Ok(Value::Empty);
        }
        if m == "visible" && c.kind == InvokeKind::Let {
            let a = arguments(c, &["Value"])?;
            let on = super::address::boolean(required(&a, 0)?)?;
            s.execute(if on { "view.unhideWindow" } else { "view.hideWindow" }, json!({"window":f.key})).map_err(|e| failure(1004, e.to_string()))?;
            return Ok(Value::Empty);
        }
        if c.kind == InvokeKind::Call && matches!(m.as_str(), "activate" | "close") {
            if m == "close" {
                // Excel ignores save options when this workbook has another window.
                arguments(c, &["SaveChanges", "Filename", "RouteWorkbook"])?;
            } else {
                arguments(c, &[])?;
            }
            s.execute(if m == "activate" { "window.activate" } else { "window.close" }, json!({"window":f.key}))
                .map_err(|e| failure(1004, e.to_string()))?;
            return Ok(Value::Empty);
        }
        if m == "selectedsheets" && reads(c) {
            let mut child = Feature::new("WindowSheets", None);
            child.key = f.key;
            return if c.arguments.is_empty() { self.feature(child) } else { self.window_dispatch(s, &child, c) };
        }
        if !reads(c) {
            return Err(failure(438, "unsupported window property assignment"));
        }
        arguments(c, &[])?;
        match m.as_str() {
            "application" => self.intern(Object::Application),
            "parent" => self.intern(Object::Workbook),
            "caption" => values::vba_text(&s.window_title(f.key).ok_or_else(|| failure(91, "window closed"))?),
            "index" => Ok(Value::Long(
                s.windows().iter().filter(|x| x.doc_uid == doc).position(|x| x.id == f.key).ok_or_else(|| failure(91, "window closed"))? as i32 + 1,
            )),
            "visible" => Ok(Value::Boolean(!w.hidden)),
            "type" => Ok(Value::Long(1)),
            "windownumber" => Ok(Value::Long(
                s.windows().iter().filter(|x| x.doc_uid == doc).position(|x| x.id == f.key).ok_or_else(|| failure(91, "window closed"))? as i32 + 1,
            )),
            "activesheet" => self.intern(Object::Sheet(s.doc().map_err(|e| failure(91, e.to_string()))?.wb.active_sheet())),
            "activecell" | "selection" | "rangeselection" => {
                let d = s.doc().map_err(|e| failure(91, e.to_string()))?;
                let r = if m == "activecell" {
                    RangeRef::cell(d.selection.active).into()
                } else {
                    ranges::Range { areas: d.selection.ranges.clone(), unit: ranges::Unit::Cells }
                };
                self.intern(Object::Range(d.wb.active_sheet(), r))
            }
            _ => Err(failure(438, "window member has no independent native state or window service")),
        }
    }
}
pub(crate) fn descriptors() -> Vec<super::MemberDescriptor> {
    use super::MemberAvailability::{Implemented as I, OutOfScope as O, ServiceRequired as S};
    [
        ("Application", "Windows", "Property Get Windows() As Windows", I),
        ("Application", "ActiveWindow", "Property Get ActiveWindow() As Window", I),
        ("Window", "NewWindow", "Function NewWindow() As Window", I),
        ("Workbook", "Windows", "Property Get Windows() As Windows", I),
        ("Workbook", "NewWindow", "Function NewWindow() As Window", I),
        ("Windows", "Count", "Property Get Count() As Long", I),
        ("Windows", "Item", "Property Get Item(Index As Variant) As Window", I),
        ("Windows", "Application", "Property Get Application() As Application", I),
        ("Windows", "Parent", "Property Get Parent() As Object", I),
        ("Window", "Application", "Property Get Application() As Application", I),
        ("Window", "Parent", "Property Get Parent() As Workbook", I),
        ("Window", "Index", "Property Get Index() As Long", I),
        ("Window", "Caption", "Property Get Caption() As String", I),
        ("Window", "Visible", "Property Visible As Boolean", I),
        ("Window", "Activate", "Sub Activate()", I),
        ("Window", "ActivateNext", "Sub ActivateNext()", I),
        ("Window", "ActivatePrevious", "Sub ActivatePrevious()", I),
        ("Window", "Close", "Function Close(Optional SaveChanges As Variant, Optional Filename As Variant, Optional RouteWorkbook As Variant) As Boolean", I),
        ("Window", "SelectedSheets", "Property Get SelectedSheets() As Sheets", I),
        ("Window", "Type", "Property Get Type() As XlWindowType", I),
        ("Window", "WindowNumber", "Property Get WindowNumber() As Long", I),
        ("Window", "ActiveSheet", "Property Get ActiveSheet() As Object", I),
        ("Window", "ActiveCell", "Property Get ActiveCell() As Range", I),
        ("Window", "Selection", "Property Get Selection() As Object", I),
        ("Window", "RangeSelection", "Property Get RangeSelection() As Range", I),
        ("Window", "Zoom", "Property Zoom As Variant", O),
        ("Window", "FreezePanes", "Property FreezePanes As Boolean", O),
        ("Window", "Split", "Property Split As Boolean", O),
        ("Window", "Panes", "Property Get Panes() As Panes", O),
        ("Window", "ScrollRow", "Property ScrollRow As Long", O),
        ("Window", "ScrollColumn", "Property ScrollColumn As Long", O),
        ("Window", "DisplayGridlines", "Property DisplayGridlines As Boolean", O),
        ("Window", "DisplayHeadings", "Property DisplayHeadings As Boolean", O),
        ("Window", "View", "Property View As XlWindowView", O),
        ("Window", "VisibleRange", "Property Get VisibleRange() As Range", S),
        ("Window", "Width", "Property Width As Double", S),
        ("Window", "Height", "Property Height As Double", S),
        ("Window", "Left", "Property Left As Double", S),
        ("Window", "Top", "Property Top As Double", S),
        ("Window", "WindowState", "Property WindowState As XlWindowState", S),
        ("Windows", "Arrange", "Sub Arrange(Optional ArrangeStyle As Variant, Optional ActiveWorkbook As Variant, Optional SyncHorizontal As Variant, Optional SyncVertical As Variant)", S),
    ].into_iter().map(|(owner,name,signature,availability)| super::MemberDescriptor {owner,name,signature,availability,implementation_symbol:"crates/vba/src/window_api.rs::Bridge::window_api"}).collect()
}
