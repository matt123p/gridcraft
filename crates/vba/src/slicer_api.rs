//! Native table slicers: one cache view per native widget, never a fabricated shared cache.
use super::features::{Feature, reads};
use super::objects::{arguments, numeric_index, required, string};
use super::{Bridge, HostCall, HostResult, InvokeKind, Object, Session, Value, failure, values};
use gridcraft_engine::model::Slicer;
use serde_json::{Value as Json, json};
fn node(kind: &'static str, si: usize, sl: &Slicer) -> Feature {
    let mut f = Feature::new(kind, Some(si));
    f.key = u64::from(sl.id);
    f.identity = Some(sl.identity.clone());
    f
}
fn nodes(s: &Session) -> HostResult<Vec<Feature>> {
    let mut out = Vec::new();
    for (si, sh) in s.doc().map_err(|e| failure(91, e.to_string()))?.wb.sheets.iter().enumerate() {
        for sl in &sh.slicers {
            if out.len() >= super::MAX_OBJECTS {
                return Err(failure(7, "slicer collection budget"));
            }
            out.push(node("SlicerCache", si, sl));
        }
    }
    Ok(out)
}
fn lookup(s: &Session, f: &Feature) -> HostResult<(usize, Slicer)> {
    let si = f.sheet.ok_or_else(|| failure(91, "slicer sheet deleted"))?;
    let sh = s.doc().map_err(|e| failure(91, e.to_string()))?.wb.sheet(si).ok_or_else(|| failure(91, "sheet"))?;
    let sl = sh
        .slicers
        .iter()
        .find(|sl| u64::from(sl.id) == f.key && f.identity.as_ref().is_some_and(|id| id.same(&sl.identity)))
        .ok_or_else(|| failure(91, "slicer deleted or replaced"))?;
    if sh.slicers.iter().filter(|x| x.id == sl.id).count() != 1 || sl.caption.len() > 65536 || sl.table.len() > 1024 || sl.column.len() > 1024 {
        return Err(failure(7, "slicer metadata or identity budget"));
    }
    if sl.selected.as_ref().is_some_and(|v| v.len() > super::MAX_OBJECTS || v.iter().any(|n| n.len() > 1024)) {
        return Err(failure(7, "slicer selection budget"));
    }
    Ok((si, sl.clone()))
}
fn items(s: &Session, si: usize, sl: &Slicer) -> HostResult<Vec<(String, bool)>> {
    let wb = &s.doc().map_err(|e| failure(91, e.to_string()))?.wb;
    let sh = wb.sheet(si).ok_or_else(|| failure(91, "sheet"))?;
    let tb = sh.tables.iter().find(|t| t.name.eq_ignore_ascii_case(&sl.table)).ok_or_else(|| failure(91, "slicer source table missing"))?;
    if tb.range.count() > super::MAX_CELLS as u64 {
        return Err(failure(7, "slicer source budget"));
    }
    if tb.column_index(&sl.column).is_none() {
        return Err(failure(91, "slicer source column missing"));
    }
    let out = gridcraft_engine::cmd::slicer::slicer_items(wb, si, sl);
    if out.len() > super::MAX_OBJECTS || out.iter().any(|(t, _)| t.len() > 1024) {
        return Err(failure(7, "slicer item budget"));
    }
    Ok(out)
}
impl Bridge {
    pub(crate) fn slicer_api(&mut self, s: &mut Session, o: &Object, call: &HostCall) -> Option<HostResult<Value>> {
        if *o == Object::Workbook && call.member.eq_ignore_ascii_case("SlicerCaches") && reads(call) {
            let f = Feature::new("SlicerCaches", None);
            return Some(if call.arguments.is_empty() { self.feature(f) } else { self.slicer_dispatch(s, &f, call) });
        }
        let Object::Feature(f) = o else {
            return None;
        };
        if !matches!(f.kind, "SlicerCaches" | "SlicerCache" | "Slicers" | "Slicer" | "SlicerItems" | "SlicerItem") {
            return None;
        }
        Some(self.slicer_dispatch(s, f, call))
    }
    fn slicer_dispatch(&mut self, s: &mut Session, f: &Feature, call: &HostCall) -> HostResult<Value> {
        let m = call.member.to_ascii_lowercase();
        if f.kind == "SlicerCaches" {
            let all = nodes(s)?;
            if call.kind == InvokeKind::Enumerate {
                arguments(call, &[])?;
                return self.object_array(all.into_iter().map(Feature::object).collect());
            }
            if !reads(call) {
                return Err(failure(438, "slicer cache collection is read-only"));
            }
            return match m.as_str() {
                "count" => {
                    arguments(call, &[])?;
                    Ok(Value::Long(all.len() as i32))
                }
                "parent" => {
                    arguments(call, &[])?;
                    self.intern(Object::Workbook)
                }
                "application" => {
                    arguments(call, &[])?;
                    self.intern(Object::Application)
                }
                "item" | "" | "slicercaches" => {
                    let a = arguments(call, &["Index"])?;
                    let v = required(&a, 0)?;
                    let n = if matches!(v, Value::String(_)) {
                        let name = string(v)?;
                        all.iter().find(|f| name.eq_ignore_ascii_case(&format!("SlicerCache{}", f.key)))
                    } else {
                        usize::try_from(numeric_index(v)?.checked_sub(1).ok_or_else(|| failure(9, "cache index"))?).ok().and_then(|i| all.get(i))
                    }
                    .ok_or_else(|| failure(9, "slicer cache not found"))?;
                    self.feature(n.clone())
                }
                _ => Err(failure(1004, "unattached/shared/OLAP/timeline caches have no native model")),
            };
        }
        let (si, sl) = lookup(s, f)?;
        if m == "application" && reads(call) {
            arguments(call, &[])?;
            return self.intern(Object::Application);
        }
        if f.kind == "Slicers" {
            if call.kind == InvokeKind::Enumerate {
                arguments(call, &[])?;
                return self.object_array(vec![node("Slicer", si, &sl).object()]);
            }
            if !reads(call) {
                return Err(failure(438, "slicer collection read-only"));
            }
            return match m.as_str() {
                "count" => {
                    arguments(call, &[])?;
                    Ok(Value::Long(1))
                }
                "parent" => {
                    arguments(call, &[])?;
                    self.feature(node("SlicerCache", si, &sl))
                }
                "item" | "" => {
                    let a = arguments(call, &["Index"])?;
                    let v = required(&a, 0)?;
                    if if matches!(v, Value::String(_)) {
                        !string(v)?.eq_ignore_ascii_case(&format!("Slicer{}", sl.id))
                    } else {
                        numeric_index(v)? != 1
                    } {
                        return Err(failure(9, "slicer not found"));
                    }
                    self.feature(node("Slicer", si, &sl))
                }
                _ => Err(failure(1004, "native widgets do not share independently-created caches")),
            };
        }
        if matches!(f.kind, "SlicerItems" | "SlicerItem") {
            return self.slicer_items(s, f, &sl, call);
        }
        if f.kind == "SlicerCache" {
            if reads(call) && matches!(m.as_str(), "slicers" | "sliceritems") {
                let n = node(if m == "slicers" { "Slicers" } else { "SlicerItems" }, si, &sl);
                if m == "sliceritems" {
                    items(s, si, &sl)?;
                }
                if call.arguments.is_empty() {
                    return self.feature(n);
                }
                let mut indexed = call.clone();
                indexed.member = "Item".into();
                return self.slicer_dispatch(s, &n, &indexed);
            }
            if !reads(call) {
                return Err(failure(1004, "native slicer cache binding is read-only"));
            }
            arguments(call, &[])?;
            return match m.as_str() {
                "parent" => self.intern(Object::Workbook),
                "index" => Ok(Value::Long(
                    nodes(s)?.iter().position(|n| n == &node("SlicerCache", si, &sl)).ok_or_else(|| failure(91, "slicer cache"))? as i32 + 1,
                )),
                "sourcetype" | "slicercachetype" => Ok(Value::Long(1)),
                "name" => values::vba_text(&format!("SlicerCache{}", sl.id)),
                "sourcename" => values::vba_text(&sl.column),
                "olap" => Ok(Value::Boolean(false)),
                "filtercleared" => Ok(Value::Boolean(sl.selected.is_none())),
                "slicers" => self.feature(node("Slicers", si, &sl)),
                "sliceritems" => {
                    items(s, si, &sl)?;
                    self.feature(node("SlicerItems", si, &sl))
                }
                "listobject" => {
                    let tb = s
                        .doc()
                        .map_err(|e| failure(91, e.to_string()))?
                        .wb
                        .sheet(si)
                        .and_then(|sh| sh.tables.iter().find(|tb| tb.name.eq_ignore_ascii_case(&sl.table)))
                        .ok_or_else(|| failure(91, "source table"))?;
                    self.feature(Self::table_node(si, tb))
                }
                "clearmanualfilter" | "clearallfilters" => self.slicer_command(s, f, "slicer.selectAll", json!({})).map(|_| Value::Empty),
                "delete" => self.slicer_command(s, f, "slicer.delete", json!({})).map(|_| Value::Empty),
                _ => Err(failure(1004, "no native shared cache, hierarchy or timeline state")),
            };
        }
        if reads(call) {
            arguments(call, &[])?;
            return match m.as_str() {
                "name" => values::vba_text(&format!("Slicer{}", sl.id)),
                "caption" => values::vba_text(&sl.caption),
                "numberofcolumns" => Ok(Value::Long(sl.columns as i32)),
                "slicercache" => self.feature(node("SlicerCache", si, &sl)),
                "shape" => {
                    let mut n = node("Shape", si, &sl);
                    n.text = Some("slicer".into());
                    self.feature(n)
                }
                "delete" => self.slicer_command(s, f, "slicer.delete", json!({})).map(|_| Value::Empty),
                "width" => Ok(Value::Double(f64::from(sl.anchor.width))),
                "height" => Ok(Value::Double(f64::from(sl.anchor.height))),
                "left" | "top" => {
                    let sh = s.doc().map_err(|e| failure(91, e.to_string()))?.wb.sheet(si).ok_or_else(|| failure(91, "sheet"))?;
                    Ok(Value::Double(if m == "left" {
                        sh.col_left(sl.anchor.cell.col) + f64::from(sl.anchor.dx)
                    } else {
                        sh.row_top(sl.anchor.cell.row) + f64::from(sl.anchor.dy)
                    }))
                }
                _ => Err(failure(1004, "slicer member has no native state")),
            };
        }
        if call.kind != InvokeKind::Let {
            return Err(failure(438, "slicer property requires Let"));
        }
        let a = arguments(call, &["Value"])?;
        let v = required(&a, 0)?;
        let mut p = json!({"kind":"slicer","id":sl.id});
        match m.as_str() {
            "caption" => {
                let text = string(v)?;
                if text.len() > 65536 {
                    return Err(failure(7, "caption budget"));
                }
                p["caption"] = json!(text);
            }
            "numberofcolumns" => {
                let n = numeric_index(v)?;
                if !(1..=10).contains(&n) {
                    return Err(failure(1004, "native columns are 1..10"));
                }
                p["columns"] = json!(n);
            }
            "width" | "height" => {
                let n = super::format_api::number(v)?;
                if !(if m == "width" { 1.0..=10000.0 } else { 0.0..=10000.0 }).contains(&n) {
                    return Err(failure(1004, "native slicer size bounds"));
                }
                p[m] = json!(n);
            }
            "left" | "top" => {
                let n = super::format_api::number(v)?;
                if !(0.0..=1e9).contains(&n) {
                    return Err(failure(1004, "slicer position bounds"));
                }
                let sh = s.doc().map_err(|e| failure(91, e.to_string()))?.wb.sheet(si).ok_or_else(|| failure(91, "sheet"))?;
                let mut at = sl.anchor.cell;
                let offset = if m == "left" {
                    at.col = sh.col_at(n);
                    sh.col_left(at.col)
                } else {
                    at.row = sh.row_at(n);
                    sh.row_top(at.row)
                };
                p["at"] = json!(at.a1());
                p[if m == "left" { "dx" } else { "dy" }] = json!(n - offset);
            }
            _ => return Err(failure(1004, "unsupported native slicer option")),
        };
        self.slicer_command(s, f, "object.move", p)?;
        Ok(Value::Empty)
    }
    fn slicer_items(&mut self, s: &mut Session, f: &Feature, sl: &Slicer, call: &HostCall) -> HostResult<Value> {
        let si = f.sheet.ok_or_else(|| failure(91, "sheet"))?;
        let list = items(s, si, sl)?;
        let m = call.member.to_ascii_lowercase();
        if f.kind == "SlicerItems" {
            if call.kind == InvokeKind::Enumerate {
                arguments(call, &[])?;
                return self.object_array(
                    list.iter()
                        .map(|(label, _)| {
                            let mut n = node("SlicerItem", si, sl);
                            n.label = Some(label.clone());
                            n.object()
                        })
                        .collect(),
                );
            }
            if !reads(call) {
                return Err(failure(438, "slicer items read-only"));
            }
            return match m.as_str() {
                "count" => {
                    arguments(call, &[])?;
                    Ok(Value::Long(list.len() as i32))
                }
                "parent" => {
                    arguments(call, &[])?;
                    self.feature(node("SlicerCache", si, sl))
                }
                "item" | "" => {
                    let a = arguments(call, &["Index"])?;
                    let v = required(&a, 0)?;
                    let item = if matches!(v, Value::String(_)) {
                        let text = string(v)?;
                        list.iter().find(|(label, _)| label.eq_ignore_ascii_case(&text))
                    } else {
                        usize::try_from(numeric_index(v)?.checked_sub(1).ok_or_else(|| failure(9, "item index"))?).ok().and_then(|i| list.get(i))
                    }
                    .ok_or_else(|| failure(9, "slicer item not found"))?;
                    let mut n = node("SlicerItem", si, sl);
                    n.label = Some(item.0.clone());
                    self.feature(n)
                }
                _ => Err(failure(438, "unsupported SlicerItems member")),
            };
        }
        let (label, selected) = list
            .iter()
            .find(|(label, _)| f.label.as_ref().is_some_and(|n| n.eq_ignore_ascii_case(label)))
            .ok_or_else(|| failure(91, "slicer item removed"))?;
        if reads(call) {
            arguments(call, &[])?;
            return match m.as_str() {
                "name" | "caption" | "value" | "sourcename" => values::vba_text(label),
                "selected" => Ok(Value::Boolean(*selected)),
                "parent" => self.feature(node("SlicerCache", si, sl)),
                _ => Err(failure(1004, "no native slicer item state")),
            };
        }
        if call.kind != InvokeKind::Let || m != "selected" {
            return Err(failure(438, "only Selected is mutable"));
        }
        let a = arguments(call, &["Value"])?;
        let on = super::address::boolean(required(&a, 0)?)?;
        if on != *selected {
            self.slicer_command(s, f, "slicer.toggle", json!({"value":label,"multi":true}))?;
        }
        Ok(Value::Empty)
    }
    fn slicer_command(&mut self, s: &mut Session, f: &Feature, command: &str, mut p: Json) -> HostResult<Json> {
        let (si, sl) = lookup(s, f)?;
        if command == "slicer.toggle" {
            items(s, si, &sl)?;
        }
        if s.doc().map_err(|e| failure(91, e.to_string()))?.wb.sheet(si).is_none_or(|sh| sh.is_protected()) {
            return Err(failure(1004, "slicer/table is protected"));
        }
        p["sheet"] = json!(si);
        p["slicer"] = json!(sl.id);
        s.execute(command, p).map_err(|e| failure(1004, e.to_string()))
    }
}
pub(crate) fn descriptors() -> Vec<super::MemberDescriptor> {
    [
        ("SlicerCache", "Index", "Property Get Index() As Long", super::MemberAvailability::Implemented),
        ("SlicerCache", "SourceType", "Property Get SourceType() As XlPivotTableSourceType", super::MemberAvailability::Implemented),
        ("SlicerCache", "SlicerCacheType", "Property Get SlicerCacheType() As XlSlicerCacheType", super::MemberAvailability::Implemented),
        ("Workbook", "SlicerCaches", "Property Get SlicerCaches() As SlicerCaches", super::MemberAvailability::Implemented),
        ("SlicerCaches", "Parent", "Property Get Parent() As Workbook", super::MemberAvailability::Implemented),
        ("SlicerCaches", "Application", "Property Get Application() As Application", super::MemberAvailability::Implemented),
        ("SlicerCaches", "Count", "Property Get Count() As Long", super::MemberAvailability::Implemented),
        ("SlicerCaches", "Item", "Function Item(Index As Variant) As SlicerCache", super::MemberAvailability::Implemented),
        ("SlicerCaches", "Add", "Function Add() As SlicerCache", super::MemberAvailability::OutOfScope),
        ("SlicerCache", "Parent", "Property Get Parent() As Workbook", super::MemberAvailability::Implemented),
        ("SlicerCache", "Application", "Property Get Application() As Application", super::MemberAvailability::Implemented),
        ("SlicerCache", "Name", "Property Get Name() As String", super::MemberAvailability::Implemented),
        ("SlicerCache", "SourceName", "Property Get SourceName() As String", super::MemberAvailability::Implemented),
        ("SlicerCache", "OLAP", "Property Get OLAP() As Boolean", super::MemberAvailability::Implemented),
        ("SlicerCache", "FilterCleared", "Property Get FilterCleared() As Boolean", super::MemberAvailability::Implemented),
        ("SlicerCache", "Slicers", "Property Get Slicers() As Slicers", super::MemberAvailability::Implemented),
        ("SlicerCache", "SlicerItems", "Property Get SlicerItems() As SlicerItems", super::MemberAvailability::Implemented),
        ("SlicerCache", "ListObject", "Property Get ListObject() As ListObject", super::MemberAvailability::Implemented),
        ("SlicerCache", "ClearManualFilter", "Sub ClearManualFilter()", super::MemberAvailability::Implemented),
        ("SlicerCache", "ClearAllFilters", "Sub ClearAllFilters()", super::MemberAvailability::Implemented),
        ("SlicerCache", "Delete", "Sub Delete()", super::MemberAvailability::Implemented),
        ("Slicers", "Parent", "Property Get Parent() As SlicerCache", super::MemberAvailability::Implemented),
        ("Slicers", "Application", "Property Get Application() As Application", super::MemberAvailability::Implemented),
        ("Slicers", "Count", "Property Get Count() As Long", super::MemberAvailability::Implemented),
        ("Slicers", "Item", "Function Item(Index As Variant) As Slicer", super::MemberAvailability::Implemented),
        ("Slicers", "Add", "Function Add() As Slicer", super::MemberAvailability::OutOfScope),
        ("Slicer", "Application", "Property Get Application() As Application", super::MemberAvailability::Implemented),
        ("Slicer", "Name", "Property Get Name() As String", super::MemberAvailability::Implemented),
        ("Slicer", "Caption", "Property Caption As String", super::MemberAvailability::Implemented),
        ("Slicer", "NumberOfColumns", "Property NumberOfColumns As Long", super::MemberAvailability::Implemented),
        ("Slicer", "Width", "Property Width As Single", super::MemberAvailability::Implemented),
        ("Slicer", "Height", "Property Height As Single", super::MemberAvailability::Implemented),
        ("Slicer", "Top", "Property Top As Single", super::MemberAvailability::Implemented),
        ("Slicer", "Left", "Property Left As Single", super::MemberAvailability::Implemented),
        ("Slicer", "Shape", "Property Get Shape() As Shape", super::MemberAvailability::Implemented),
        ("Slicer", "SlicerCache", "Property Get SlicerCache() As SlicerCache", super::MemberAvailability::Implemented),
        ("Slicer", "Delete", "Sub Delete()", super::MemberAvailability::Implemented),
        ("SlicerItems", "Parent", "Property Get Parent() As SlicerCache", super::MemberAvailability::Implemented),
        ("SlicerItems", "Application", "Property Get Application() As Application", super::MemberAvailability::Implemented),
        ("SlicerItems", "Count", "Property Get Count() As Long", super::MemberAvailability::Implemented),
        ("SlicerItems", "Item", "Function Item(Index As Variant) As SlicerItem", super::MemberAvailability::Implemented),
        ("SlicerItem", "Parent", "Property Get Parent() As SlicerCache", super::MemberAvailability::Implemented),
        ("SlicerItem", "Application", "Property Get Application() As Application", super::MemberAvailability::Implemented),
        ("SlicerItem", "Name", "Property Get Name() As String", super::MemberAvailability::Implemented),
        ("SlicerItem", "Caption", "Property Get Caption() As String", super::MemberAvailability::Implemented),
        ("SlicerItem", "SourceName", "Property Get SourceName() As String", super::MemberAvailability::Implemented),
        ("SlicerItem", "Value", "Property Get Value() As String", super::MemberAvailability::Implemented),
        ("SlicerItem", "Selected", "Property Selected As Boolean", super::MemberAvailability::Implemented),
    ]
    .into_iter()
    .map(|(owner, name, signature, availability)| super::MemberDescriptor {
        owner,
        name,
        signature,
        availability,
        implementation_symbol: "crates/vba/src/slicer_api.rs::Bridge::slicer_api",
    })
    .collect()
}
