//! Native manual markers plus derived pagination boundaries. Handles expire on document edits.
use super::features::{Feature, reads};
use super::objects::{arguments, numeric_index, required};
use super::{Bridge, HostCall, HostResult, InvokeKind, Object, Session, Value, failure, ranges};
use gridcraft_engine::cmd::print::{PrintOptions, layout_sheet};
use gridcraft_engine::core::{CellRef, RangeRef};
use serde_json::json;
use std::collections::BTreeMap;
fn horizontal(f: &Feature) -> bool {
    matches!(f.kind, "HPageBreaks" | "HPageBreak")
}
fn breaks(s: &Session, si: usize, h: bool) -> HostResult<Vec<(u32, bool)>> {
    let d = s.doc().map_err(|e| failure(91, e.to_string()))?;
    let sh = d.wb.sheet(si).ok_or_else(|| failure(91, "worksheet deleted"))?;
    super::page_setup_api::validate_settings(&sh.print)?;
    let manual = if h { &sh.print.row_breaks } else { &sh.print.col_breaks };
    if manual.iter().any(|n| *n == 0 || *n >= if h { gridcraft_engine::core::MAX_ROWS } else { gridcraft_engine::core::MAX_COLS }) {
        return Err(failure(1004, "invalid native page break bounds"));
    }
    if manual.len() > 1026 {
        return Err(failure(7, "manual break budget"));
    }
    let mut all = BTreeMap::new();
    if let Some(lay) = layout_sheet(&d.wb, si, &PrintOptions::default(), None) {
        for (start, _) in (if h { lay.row_bands } else { lay.col_bands }).into_iter().skip(1) {
            all.insert(start, false);
        }
    }
    for n in manual {
        all.insert(*n, true);
    }
    if all.len() > super::MAX_OBJECTS {
        return Err(failure(7, "page break budget"));
    }
    Ok(all.into_iter().collect())
}
fn node(s: &Session, si: usize, h: bool, n: u32) -> HostResult<Feature> {
    let mut f = Feature::new(if h { "HPageBreak" } else { "VPageBreak" }, Some(si));
    f.key = u64::from(n);
    f.text = Some(s.doc().map_err(|e| failure(91, e.to_string()))?.revision.to_string());
    Ok(f)
}
fn update(s: &mut Session, si: usize, h: bool, n: u32, add: bool) -> HostResult<Value> {
    let sh = s.doc().map_err(|e| failure(91, e.to_string()))?.wb.sheet(si).ok_or_else(|| failure(91, "worksheet deleted"))?;
    let mut list = if h { sh.print.row_breaks.clone() } else { sh.print.col_breaks.clone() };
    if n == 0 || n >= if h { gridcraft_engine::core::MAX_ROWS } else { gridcraft_engine::core::MAX_COLS } {
        return Err(failure(1004, "break must be after the first row/column and inside the sheet"));
    }
    if list.iter().any(|n| *n == 0 || *n >= if h { gridcraft_engine::core::MAX_ROWS } else { gridcraft_engine::core::MAX_COLS }) {
        return Err(failure(1004, "invalid native page break bounds"));
    }
    list.retain(|i| *i != n);
    if add {
        list.push(n);
    }
    list.sort_unstable();
    list.dedup();
    if list.len() > 1026 {
        return Err(failure(7, "manual break budget"));
    }
    let mut p = json!({"sheet":si});
    p[if h { "rowBreaks" } else { "colBreaks" }] = json!(list.iter().map(|n| n + 1).collect::<Vec<_>>());
    s.execute("pageLayout.pageSetup", p).map_err(|e| failure(1004, e.to_string()))?;
    Ok(Value::Empty)
}
impl Bridge {
    pub(crate) fn page_break_api(&mut self, s: &mut Session, o: &Object, c: &HostCall) -> Option<HostResult<Value>> {
        let m = c.member.to_ascii_lowercase();
        if let Object::Sheet(si) = o {
            if matches!(m.as_str(), "hpagebreaks" | "vpagebreaks") && reads(c) {
                let f = Feature::new(if m == "hpagebreaks" { "HPageBreaks" } else { "VPageBreaks" }, Some(*si));
                return Some(if c.arguments.is_empty() { self.feature(f) } else { self.page_break_dispatch(s, &f, c) });
            }
            if m == "resetallpagebreaks" && c.kind == InvokeKind::Call {
                return Some((|| {
                    arguments(c, &[])?;
                    s.execute("pageLayout.pageSetup", json!({"sheet":si,"rowBreaks":[],"colBreaks":[]})).map_err(|e| failure(1004, e.to_string()))?;
                    Ok(Value::Empty)
                })());
            }
        }
        if let Object::Range(si, r) = o
            && m == "pagebreak"
        {
            return Some(self.range_page_break(s, *si, r, c));
        }
        let Object::Feature(f) = o else {
            return None;
        };
        if !matches!(f.kind, "HPageBreaks" | "VPageBreaks" | "HPageBreak" | "VPageBreak") {
            return None;
        }
        Some(self.page_break_dispatch(s, f, c))
    }
    fn page_break_dispatch(&mut self, s: &mut Session, f: &Feature, c: &HostCall) -> HostResult<Value> {
        let si = f.sheet.ok_or_else(|| failure(91, "worksheet deleted"))?;
        let h = horizontal(f);
        let m = c.member.to_ascii_lowercase();
        let all = breaks(s, si, h)?;
        if matches!(f.kind, "HPageBreaks" | "VPageBreaks") {
            if c.kind == InvokeKind::Enumerate {
                arguments(c, &[])?;
                let nodes = all.iter().map(|(n, _)| node(s, si, h, *n).map(Feature::object)).collect::<HostResult<Vec<_>>>()?;
                return self.object_array(nodes);
            }
            if m == "add" && reads(c) {
                let a = arguments(c, &["Before"])?;
                let (owner, cell) = self.break_before(s, required(&a, 0)?)?;
                if owner != si {
                    return Err(failure(1004, "break receiver and Before must have the same worksheet"));
                }
                self.objects.reserve(1)?;
                let n = if h { cell.row } else { cell.col };
                update(s, si, h, n, true)?;
                return self.feature(node(s, si, h, n)?);
            }
            if !reads(c) {
                return Err(failure(438, "read-only page break collection"));
            }
            return match m.as_str() {
                "count" => {
                    arguments(c, &[])?;
                    Ok(Value::Long(all.len() as i32))
                }
                "parent" => {
                    arguments(c, &[])?;
                    self.intern(Object::Sheet(si))
                }
                "application" => {
                    arguments(c, &[])?;
                    self.intern(Object::Application)
                }
                "item" | "" | "hpagebreaks" | "vpagebreaks" => {
                    let a = arguments(c, &["Index"])?;
                    let i = numeric_index(required(&a, 0)?)?;
                    let n = usize::try_from(i.checked_sub(1).ok_or_else(|| failure(9, "page break index"))?)
                        .ok()
                        .and_then(|i| all.get(i))
                        .map(|x| x.0)
                        .ok_or_else(|| failure(9, "page break index"))?;
                    self.feature(node(s, si, h, n)?)
                }
                _ => Err(failure(438, "unsupported page break collection member")),
            };
        }
        super::page_setup_api::fresh(s, f)?;
        let n = u32::try_from(f.key).map_err(|_| failure(91, "page break deleted"))?;
        let manual = all.iter().find(|(p, _)| *p == n).map(|(_, manual)| *manual).ok_or_else(|| failure(91, "page break deleted"))?;
        if m == "delete" && c.kind == InvokeKind::Call {
            arguments(c, &[])?;
            if !manual {
                return Err(failure(1004, "automatic breaks cannot be deleted"));
            }
            return update(s, si, h, n, false);
        }
        if m == "type" && c.kind == InvokeKind::Let {
            let a = arguments(c, &["Value"])?;
            let v = numeric_index(required(&a, 0)?)?;
            if v != -4135 && !(v == -4142 && manual) {
                return Err(failure(1004, "only manual breaks can be inserted/removed"));
            }
            return update(s, si, h, n, v == -4135);
        }
        if m == "location" && c.kind == InvokeKind::Set {
            let a = arguments(c, &["Value"])?;
            if !h || !manual {
                return Err(failure(1004, "only horizontal manual break Location assignment is supported"));
            }
            let (owner, cell) = self.break_before(s, required(&a, 0)?)?;
            if owner != si || cell.row == 0 {
                return Err(failure(1004, "break location must be on the receiver sheet after row 1"));
            }
            let mut rows =
                s.doc().map_err(|e| failure(91, e.to_string()))?.wb.sheet(si).ok_or_else(|| failure(91, "sheet"))?.print.row_breaks.clone();
            rows.retain(|r| *r != n);
            rows.push(cell.row);
            rows.sort_unstable();
            rows.dedup();
            s.execute("pageLayout.pageSetup", json!({"sheet":si,"rowBreaks":rows.iter().map(|n|n+1).collect::<Vec<_>>()}))
                .map_err(|e| failure(1004, e.to_string()))?;
            return Ok(Value::Empty);
        }
        if !reads(c) {
            return Err(failure(438, "unsupported page break assignment"));
        }
        arguments(c, &[])?;
        match m.as_str() {
            "parent" => self.intern(Object::Sheet(si)),
            "application" => self.intern(Object::Application),
            "location" => self.intern(Object::Range(si, RangeRef::cell(if h { CellRef::new(n, 0) } else { CellRef::new(0, n) }).into())),
            "type" => Ok(Value::Long(if manual { -4135 } else { -4105 })),
            "extent" if manual => Ok(Value::Long(1)),
            _ => Err(failure(438, "page break extent/dragging needs an independent native model or viewport service")),
        }
    }
    fn break_before(&self, s: &Session, v: &Value) -> HostResult<(usize, CellRef)> {
        let Value::Object(reference) = v else {
            return Err(failure(13, "Before requires a Range"));
        };
        let Object::Range(si, r) = self.object(s, *reference)? else {
            return Err(failure(13, "Before requires a Range"));
        };
        Ok((si, r.rectangle()?.start))
    }
    fn range_page_break(&mut self, s: &mut Session, si: usize, r: &ranges::Range, c: &HostCall) -> HostResult<Value> {
        let rect = r.rectangle()?;
        let h = rect.is_full_rows();
        let v = rect.is_full_cols();
        if c.kind == InvokeKind::Let && h && v {
            let a = arguments(c, &["Value"])?;
            if numeric_index(required(&a, 0)?)? != -4142 {
                return Err(failure(1004, "Cells.PageBreak only supports clearing all manual markers"));
            }
            s.execute("pageLayout.pageSetup", json!({"sheet":si,"rowBreaks":[],"colBreaks":[]})).map_err(|e| failure(1004, e.to_string()))?;
            return Ok(Value::Empty);
        }
        if h == v || h && rect.start.row != rect.end.row || v && rect.start.col != rect.end.col {
            return Err(failure(1004, "PageBreak requires one whole row or column"));
        }
        let n = if h { rect.start.row } else { rect.start.col };
        if reads(c) {
            arguments(c, &[])?;
            let found = breaks(s, si, h)?.into_iter().find(|(p, _)| *p == n);
            return Ok(Value::Long(match found {
                None => -4142,
                Some((_, true)) => -4135,
                _ => -4105,
            }));
        }
        if c.kind != InvokeKind::Let {
            return Err(failure(438, "PageBreak requires Let"));
        }
        let a = arguments(c, &["Value"])?;
        let value = numeric_index(required(&a, 0)?)?;
        if !matches!(value, -4135 | -4142) {
            return Err(failure(1004, "automatic breaks are derived, not assignable"));
        }
        if value == -4142 && n == 0 {
            return Ok(Value::Empty);
        }
        update(s, si, h, n, value == -4135)
    }
}

pub(crate) fn descriptors() -> Vec<super::MemberDescriptor> {
    [
        ("Worksheet", "HPageBreaks", "Property Get HPageBreaks() As HPageBreaks"),
        ("Worksheet", "VPageBreaks", "Property Get VPageBreaks() As VPageBreaks"),
        ("Worksheet", "ResetAllPageBreaks", "Sub ResetAllPageBreaks()"),
        ("Range", "PageBreak", "Property PageBreak As Long"),
        ("HPageBreaks", "Count", "Property Get Count() As Long"),
        ("HPageBreaks", "Item", "Property Get Item(Index As Variant) As HPageBreak"),
        ("HPageBreaks", "Add", "Function Add(Before As Range) As HPageBreak"),
        ("HPageBreaks", "Parent", "Property Get Parent() As Worksheet"),
        ("HPageBreaks", "Application", "Property Get Application() As Application"),
        ("HPageBreak", "Type", "Property Type As XlPageBreak"),
        ("HPageBreak", "Location", "Property Location As Range"),
        ("HPageBreak", "Delete", "Sub Delete()"),
        ("HPageBreak", "Extent", "Property Get Extent() As XlPageBreakExtent"),
        ("HPageBreak", "Parent", "Property Get Parent() As Worksheet"),
        ("HPageBreak", "Application", "Property Get Application() As Application"),
        ("VPageBreaks", "Count", "Property Get Count() As Long"),
        ("VPageBreaks", "Item", "Property Get Item(Index As Variant) As VPageBreak"),
        ("VPageBreaks", "Add", "Function Add(Before As Range) As VPageBreak"),
        ("VPageBreaks", "Parent", "Property Get Parent() As Worksheet"),
        ("VPageBreaks", "Application", "Property Get Application() As Application"),
        ("VPageBreak", "Type", "Property Type As XlPageBreak"),
        ("VPageBreak", "Location", "Property Location As Range"),
        ("VPageBreak", "Delete", "Sub Delete()"),
        ("VPageBreak", "Extent", "Property Get Extent() As XlPageBreakExtent"),
        ("VPageBreak", "Parent", "Property Get Parent() As Worksheet"),
        ("VPageBreak", "Application", "Property Get Application() As Application"),
    ]
    .into_iter()
    .map(|(owner, name, signature)| super::MemberDescriptor {
        owner,
        name,
        signature,
        availability: super::MemberAvailability::Implemented,
        implementation_symbol: "crates/vba/src/page_break_api.rs::Bridge::page_break_api",
    })
    .collect()
}
