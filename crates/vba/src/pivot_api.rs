//! OM13 local PivotTable graph. Cache views mirror the native one-cache-per-report export.
use super::features::{Feature, absent, reads};
use super::objects::{arguments, numeric_index, required, string};
use super::{Bridge, HostCall, HostResult, InvokeKind, Object, Session, Value, failure, values};
use gridcraft_engine::{
    core::{CellRef, RangeRef},
    model::PivotTable,
    pivot as pv,
};
use serde_json::{Value as Json, json};
pub(crate) fn node(kind: &'static str, sheet: usize, pt: &PivotTable) -> Feature {
    let mut f = Feature::new(kind, Some(sheet));
    f.key = u64::from(pt.id);
    f.identity = Some(pt.identity.clone());
    f
}
pub(crate) fn lookup(s: &Session, f: &Feature) -> HostResult<(usize, PivotTable)> {
    let sheet = f.sheet.ok_or_else(|| failure(91, "pivot owner deleted"))?;
    let pt = s
        .doc()
        .map_err(|e| failure(91, e.to_string()))?
        .wb
        .sheet(sheet)
        .and_then(|sh| sh.pivots.iter().find(|pt| u64::from(pt.id) == f.key && f.identity.as_ref().is_some_and(|id| id.same(&pt.identity))))
        .ok_or_else(|| failure(91, "PivotTable deleted or replaced"))?;
    if pt.name.len() > 1024
        || pt.source.len() > 65536
        || pt.rows.len() + pt.columns.len() > 32
        || pt.values.len() > 256
        || pt.filters.len() > 256
        || pt.calculated_fields.len() > 256
    {
        return Err(failure(7, "pivot metadata budget"));
    }
    if pt.values.iter().any(|v| v.name.len() > 1024 || v.source_col.len() > 1024 || v.number_format.as_ref().is_some_and(|n| n.len() > 255))
        || pt.calculated_fields.iter().any(|f| f.name.len() > 1024 || f.formula.len() > 65536)
        || pt
            .rows
            .iter()
            .chain(&pt.columns)
            .any(|f| f.source_col.len() > 1024 || f.collapsed_items.len() > super::MAX_OBJECTS || f.collapsed_items.iter().any(|n| n.len() > 1024))
        || pt.filters.iter().any(|f| {
            f.source_col.len() > 1024 || f.selected.as_ref().is_some_and(|v| v.len() > super::MAX_OBJECTS || v.iter().any(|n| n.len() > 1024))
        })
    {
        return Err(failure(7, "pivot field metadata budget"));
    }
    Ok((sheet, pt.clone()))
}
pub(crate) fn source(s: &Session, sheet: usize, pt: &PivotTable) -> HostResult<pv::Source> {
    let wb = &s.doc().map_err(|e| failure(91, e.to_string()))?.wb;
    let (_, r) = pv::resolve_source(wb, &pt.source, sheet).map_err(|e| failure(1004, e))?;
    if r.count() > super::MAX_CELLS as u64 || r.width() > 256 {
        return Err(failure(7, "pivot source budget"));
    }
    let src = pv::read_source(wb, &pt.source, sheet).map_err(|e| failure(1004, e))?;
    if src.headers.iter().any(|t| t.len() > 1024) {
        return Err(failure(7, "pivot field name budget"));
    }
    Ok(src)
}
fn pivots(s: &Session, sheet: Option<usize>, kind: &'static str) -> HostResult<Vec<Feature>> {
    let d = s.doc().map_err(|e| failure(91, e.to_string()))?;
    let mut out = Vec::new();
    for (si, sh) in d.wb.sheets.iter().enumerate().filter(|(si, _)| sheet.is_none_or(|i| i == *si)) {
        for pt in &sh.pivots {
            if out.len() >= super::MAX_OBJECTS {
                return Err(failure(7, "pivot collection budget"));
            }
            out.push(node(kind, si, pt));
        }
    }
    Ok(out)
}
impl Bridge {
    pub(crate) fn pivot_api(&mut self, s: &mut Session, o: &Object, call: &HostCall) -> Option<HostResult<Value>> {
        let m = call.member.to_ascii_lowercase();
        if (matches!(o, Object::Workbook) && m == "pivotcaches" || matches!(o, Object::Sheet(_)) && m == "pivottables") && reads(call) {
            let f = Feature::new(
                if m == "pivotcaches" { "PivotCaches" } else { "PivotTables" },
                if let Object::Sheet(si) = o { Some(*si) } else { None },
            );
            return Some(if call.arguments.is_empty() { self.feature(f) } else { self.pivot_collection(s, &f, call) });
        }
        if let Object::Range(si, range) = o
            && matches!(m.as_str(), "pivottable" | "pivotcell" | "showdetail")
        {
            if m == "showdetail"
                && !range.rectangle().ok().is_some_and(|rect| {
                    s.doc()
                        .ok()
                        .and_then(|doc| doc.wb.sheet(*si))
                        .is_some_and(|sh| sh.pivots.iter().any(|pt| pt.last_range.is_some_and(|r| r.contains_range(&rect))))
                })
            {
                return None;
            }
            return Some((|| {
                let rect = range.rectangle()?;
                let pt = s
                    .doc()
                    .map_err(|e| failure(91, e.to_string()))?
                    .wb
                    .sheet(*si)
                    .and_then(|sh| sh.pivots.iter().find(|pt| pt.last_range.is_some_and(|r| r.contains_range(&rect))))
                    .ok_or_else(|| failure(1004, "range is not in a PivotTable"))?;
                if m == "showdetail" {
                    if call.kind != InvokeKind::Let {
                        return Err(failure(1004, "use True to request native drilldown"));
                    }
                    let a = arguments(call, &["Value"])?;
                    if !super::address::boolean(required(&a, 0)?)? || !rect.is_single() {
                        return Err(failure(1004, "drilldown requires one native value cell and True"));
                    }
                    let f = node("PivotTable", *si, pt);
                    return self.pivot_command(s, &f, "pivot.showDetails", json!({"cell":rect.start.a1()})).map(|_| Value::Empty);
                }
                if !reads(call) {
                    return Err(failure(438, "pivot range view is read-only"));
                }
                arguments(call, &[])?;
                let mut f = node(if m == "pivotcell" { "PivotCell" } else { "PivotTable" }, *si, pt);
                if m == "pivotcell" {
                    if !rect.is_single() {
                        return Err(failure(1004, "PivotCell requires one cell"));
                    }
                    f.range = Some(rect.into());
                }
                self.feature(f)
            })());
        }
        let Object::Feature(f) = o else {
            return None;
        };
        if f.kind == "Chart" && m == "pivotlayout" && reads(call) {
            return Some(self.pivot_chart_layout(s, f, call));
        }
        if !matches!(
            f.kind,
            "PivotCaches"
                | "PivotCache"
                | "PivotTables"
                | "PivotTable"
                | "PivotFields"
                | "PivotField"
                | "PivotItems"
                | "PivotItem"
                | "CalculatedFields"
                | "PivotCell"
                | "PivotLayout"
        ) {
            return None;
        }
        Some(self.pivot_dispatch(s, f, call))
    }
    fn pivot_dispatch(&mut self, s: &mut Session, f: &Feature, call: &HostCall) -> HostResult<Value> {
        if matches!(f.kind, "PivotCaches" | "PivotTables") {
            return self.pivot_collection(s, f, call);
        }
        let (sheet, pt) = lookup(s, f)?;
        let m = call.member.to_ascii_lowercase();
        if matches!(f.kind, "PivotFields" | "PivotField" | "PivotItems" | "PivotItem" | "CalculatedFields") {
            return self.pivot_fields(s, f, &pt, call);
        }
        if m == "application" && reads(call) {
            arguments(call, &[])?;
            return self.intern(Object::Application);
        }
        if f.kind == "PivotLayout" {
            return self.pivot_layout(s, f, &pt, call);
        }
        if f.kind == "PivotCell" {
            arguments(call, &[])?;
            if !reads(call) {
                return Err(failure(438, "PivotCell is read-only"));
            }
            return match m.as_str() {
                "pivottable" => self.feature(node("PivotTable", sheet, &pt)),
                "parent" | "range" => self.intern(Object::Range(sheet, f.range.clone().ok_or_else(|| failure(91, "pivot cell deleted"))?)),
                "pivotcelltype" => {
                    let rect = f.range.as_ref().ok_or_else(|| failure(91, "pivot cell"))?.rectangle()?;
                    let area = pt.last_range.ok_or_else(|| failure(91, "report missing"))?;
                    if !area.contains(rect.start) {
                        return Err(failure(91, "cell left report"));
                    }
                    source(s, sheet, &pt)?;
                    let out = pv::current_output(&s.doc().map_err(|e| failure(91, e.to_string()))?.wb, &pt, sheet).map_err(|e| failure(1004, e))?;
                    let role = out
                        .grid
                        .get((rect.start.row - area.start.row) as usize)
                        .and_then(|row| row.get((rect.start.col - area.start.col) as usize))
                        .ok_or_else(|| failure(91, "pivot cell"))?
                        .kind;
                    use pv::CellKind as K;
                    Ok(Value::Long(match role {
                        K::Data => 0,
                        K::RowLabel | K::ColLabel => 1,
                        K::Subtotal => 2,
                        K::GrandTotal => 3,
                        K::Header => 5,
                        K::Filter => 6,
                        K::Empty | K::Placeholder => 9,
                    }))
                }
                _ => Err(failure(1004, "no native cell field/axis object metadata")),
            };
        }
        if f.kind == "PivotCache" {
            return self.pivot_cache(s, f, &pt, call);
        }
        if reads(call) {
            if matches!(m.as_str(), "pivotfields" | "rowfields" | "columnfields" | "pagefields" | "datafields" | "hiddenfields" | "calculatedfields")
            {
                let mut n = node(if m == "calculatedfields" { "CalculatedFields" } else { "PivotFields" }, sheet, &pt);
                n.part = match m.as_str() {
                    "rowfields" => 1,
                    "columnfields" => 2,
                    "pagefields" => 3,
                    "datafields" => 4,
                    "hiddenfields" => 5,
                    "calculatedfields" => 6,
                    _ => 0,
                };
                n.ancestors.push(pt.source_identity.clone());
                return if call.arguments.is_empty() { self.feature(n) } else { self.pivot_fields(s, &n, &pt, call) };
            }
            if m == "addfields" {
                return self.pivot_add_fields(s, f, call);
            }
            if m == "subtotallocation" {
                let a = arguments(call, &["Location"])?;
                let loc = match numeric_index(required(&a, 0)?)? {
                    1 => "top",
                    2 => "bottom",
                    _ => return Err(failure(1004, "subtotal location")),
                };
                return self.pivot_command(s, f, "pivot.layout", json!({"subtotals":loc})).map(|_| Value::Empty);
            }
            if m == "adddatafield" {
                return self.pivot_add_data(s, f, &pt, call);
            }
            if m == "rowaxislayout" {
                let a = arguments(call, &["RowLayout"])?;
                let layout = match numeric_index(required(&a, 0)?)? {
                    0 => "compact",
                    1 => "tabular",
                    2 => "outline",
                    _ => return Err(failure(1004, "unsupported row layout")),
                };
                return self.pivot_command(s, f, "pivot.layout", json!({"layout":layout})).map(|_| Value::Empty);
            }
            arguments(call, &[])?;
            return match m.as_str() {
                "cleartable" => self
                    .pivot_command(s, f, "pivot.addField", json!({"areas":{"rows":[],"columns":[],"filters":[],"values":[]}}))
                    .map(|_| Value::Empty),
                "clearallfilters" => self.pivot_command(s, f, "pivot.filter", json!({"all":true})).map(|_| Value::Empty),
                "location" => {
                    let name = &s.doc().map_err(|e| failure(91, e.to_string()))?.wb.sheet(sheet).ok_or_else(|| failure(91, "sheet"))?.name;
                    values::vba_text(&pv::range_text(name, RangeRef::cell(pt.anchor)))
                }
                "pivotchart" => {
                    let d = s.doc().map_err(|e| failure(91, e.to_string()))?;
                    let sh = d.wb.sheet(sheet).ok_or_else(|| failure(91, "sheet"))?;
                    let (si, c) =
                        d.wb.sheets
                            .iter()
                            .enumerate()
                            .flat_map(|(si, sh)| sh.charts.iter().map(move |c| (si, c)))
                            .find(|(_, c)| c.pivot.as_deref().is_some_and(|link| pv::links_to(link, &sh.name, &pt.name)))
                            .ok_or_else(|| failure(1004, "PivotTable has no native PivotChart"))?;
                    let mut n = Feature::new("Chart", Some(si));
                    n.key = u64::from(c.id);
                    n.identity = Some(c.identity.clone());
                    self.feature(n)
                }
                "parent" => self.intern(Object::Sheet(sheet)),
                "name" => values::vba_text(&pt.name),
                "sourcedata" => values::vba_text(&pt.source),
                "datapivotfield" => self.pivot_values_node(f, &pt),
                "pivotcache" => self.feature(node("PivotCache", sheet, &pt)),
                "cacheindex" => Ok(Value::Long(
                    pivots(s, None, "PivotCache")?.iter().position(|n| n == &node("PivotCache", sheet, &pt)).ok_or_else(|| failure(91, "cache"))?
                        as i32
                        + 1,
                )),
                "refreshtable" => self.pivot_command(s, f, "pivot.refresh", json!({})).map(|_| Value::Boolean(true)),
                "update" => self.pivot_command(s, f, "pivot.refresh", json!({})).map(|_| Value::Empty),
                "rowgrand" => Ok(Value::Boolean(pt.grand_totals_rows)),
                "columngrand" => Ok(Value::Boolean(pt.grand_totals_cols)),
                "displayfieldcaptions" => Ok(Value::Boolean(pt.show_headers)),
                "enableshowdetails" | "enabledrilldown" => Ok(Value::Boolean(pt.enable_show_details)),
                "tablestyle2" => values::vba_text(&pt.style),
                "tablerange1" | "tablerange2" | "databodyrange" => self.pivot_range(s, sheet, &pt, &m),
                _ => Err(failure(1004, "PivotTable member has no native profile")),
            };
        }
        if call.kind != InvokeKind::Let {
            return Err(failure(438, "unsupported pivot invocation"));
        }
        let a = arguments(call, &["Value"])?;
        let v = required(&a, 0)?;
        let (command, p) = match m.as_str() {
            "rowgrand" => ("pivot.layout", json!({"grandTotalsRows":super::address::boolean(v)?})),
            "columngrand" => ("pivot.layout", json!({"grandTotalsCols":super::address::boolean(v)?})),
            "displayfieldcaptions" => ("pivot.layout", json!({"showHeaders":super::address::boolean(v)?})),
            "tablestyle2" => ("pivot.layout", json!({"style":string(v)?})),
            "enableshowdetails" | "enabledrilldown" => ("pivot.options", json!({"enableShowDetails":super::address::boolean(v)?})),
            "sourcedata" => {
                let text = self.pivot_source_arg(s, v, sheet)?;
                ("pivot.changeSource", json!({"source":text}))
            }
            _ => return Err(failure(1004, "no native mutable PivotTable setting")),
        };
        self.pivot_command(s, f, command, p)?;
        Ok(Value::Empty)
    }
    fn pivot_collection(&mut self, s: &mut Session, f: &Feature, call: &HostCall) -> HostResult<Value> {
        let kind = if f.kind == "PivotCaches" { "PivotCache" } else { "PivotTable" };
        let nodes = pivots(s, f.sheet, kind)?;
        let m = call.member.to_ascii_lowercase();
        if call.kind == InvokeKind::Enumerate {
            arguments(call, &[])?;
            return self.object_array(nodes.into_iter().map(Feature::object).collect());
        }
        if !reads(call) {
            return Err(failure(438, "pivot collection is read-only"));
        }
        match m.as_str() {
            "count" => {
                arguments(call, &[])?;
                Ok(Value::Long(nodes.len() as i32))
            }
            "parent" => {
                arguments(call, &[])?;
                if let Some(si) = f.sheet { self.intern(Object::Sheet(si)) } else { self.intern(Object::Workbook) }
            }
            "application" => {
                arguments(call, &[])?;
                self.intern(Object::Application)
            }
            "item" | "" | "pivotcaches" | "pivottables" => {
                let a = arguments(call, &["Index"])?;
                let v = required(&a, 0)?;
                let n = if matches!(v, Value::String(_)) && kind == "PivotTable" {
                    let name = string(v)?;
                    nodes.iter().find(|n| lookup(s, n).is_ok_and(|(_, pt)| pt.name.eq_ignore_ascii_case(&name)))
                } else {
                    usize::try_from(numeric_index(v)?.checked_sub(1).ok_or_else(|| failure(9, "pivot index"))?).ok().and_then(|i| nodes.get(i))
                }
                .ok_or_else(|| failure(9, "pivot not found"))?;
                self.feature(n.clone())
            }
            "add" if kind == "PivotTable" => self.pivot_add(s, f, call),
            _ => Err(failure(1004, "native caches are report-owned; unattached/shared/provider caches are unavailable")),
        }
    }
    pub(crate) fn pivot_command(&mut self, s: &mut Session, f: &Feature, command: &str, mut p: Json) -> HostResult<Json> {
        let (si, pt) = lookup(s, f)?;
        source(s, si, &pt)?;
        let d = s.doc().map_err(|e| failure(91, e.to_string()))?;
        let sh = d.wb.sheet(si).ok_or_else(|| failure(91, "sheet"))?;
        if sh.is_protected() {
            return Err(failure(1004, "PivotTable report is protected"));
        }
        if sh.pivots.iter().filter(|other| other.id == pt.id).count() != 1 {
            return Err(failure(1004, "ambiguous pivot ID"));
        }
        // A refresh can rebind PivotCharts on other sheets. Preflight those drawing locks too.
        if d.wb.sheets.iter().any(|sheet| {
            sheet.is_protected() && sheet.charts.iter().any(|c| c.pivot.as_deref().is_some_and(|link| pv::links_to(link, &sh.name, &pt.name)))
        }) {
            return Err(failure(1004, "linked PivotChart is protected"));
        }
        p["sheet"] = json!(si);
        p["pivot"] = json!(pt.id);
        s.execute(command, p).map_err(|e| failure(1004, e.to_string()))
    }
    fn pivot_cache(&mut self, s: &mut Session, f: &Feature, pt: &PivotTable, call: &HostCall) -> HostResult<Value> {
        let sheet = f.sheet.ok_or_else(|| failure(91, "sheet"))?;
        let m = call.member.to_ascii_lowercase();
        if !reads(call) {
            if call.kind == InvokeKind::Let && m == "sourcedata" {
                let a = arguments(call, &["Value"])?;
                let text = self.pivot_source_arg(s, required(&a, 0)?, sheet)?;
                self.pivot_command(s, f, "pivot.changeSource", json!({"source":text}))?;
                return Ok(Value::Empty);
            }
            return Err(failure(1004, "no mutable provider/cache service state"));
        }
        arguments(call, &[])?;
        match m.as_str() {
            "parent" => self.intern(Object::Workbook),
            "sourcedata" => values::vba_text(&pt.source),
            "sourcetype" => Ok(Value::Long(1)),
            "olap" => Ok(Value::Boolean(false)),
            "recordcount" => Ok(Value::Long(source(s, sheet, pt)?.rows.len() as i32)),
            "index" => Ok(Value::Long(pivots(s, None, "PivotCache")?.iter().position(|n| n == f).ok_or_else(|| failure(91, "cache"))? as i32 + 1)),
            "refresh" => self.pivot_command(s, f, "pivot.refresh", json!({})).map(|_| Value::Empty),
            _ => Err(failure(1004, "no native shared cache/provider/async service")),
        }
    }
    pub(crate) fn pivot_source_arg(&self, s: &Session, v: &Value, sheet: usize) -> HostResult<String> {
        let text = if matches!(v, Value::Object(_)) { self.range_binding(s, v)? } else { string(v)? };
        if text.is_empty() || text.len() > 65536 {
            return Err(failure(7, "pivot source text budget"));
        }
        let wb = &s.doc().map_err(|e| failure(91, e.to_string()))?.wb;
        let (si, r) = pv::resolve_source(wb, &text, sheet).map_err(|e| failure(1004, e))?;
        if r.count() > super::MAX_CELLS as u64 || r.width() > 256 {
            return Err(failure(7, "pivot source budget"));
        }
        // Freeze unqualified ranges to their explicit owner; native named sources remain named.
        if !text.contains('!') && RangeRef::parse(text.trim_start_matches('=')).is_some() {
            let name = &wb.sheet(si).ok_or_else(|| failure(91, "source sheet"))?.name;
            Ok(pv::range_text(name, r))
        } else {
            Ok(text)
        }
    }
    fn pivot_add(&mut self, s: &mut Session, f: &Feature, call: &HostCall) -> HostResult<Value> {
        let a = arguments(call, &["SourceType", "SourceData", "TableDestination", "TableName", "ReadData", "DefaultVersion"])?;
        if numeric_index(required(&a, 0)?)? != 1 || a.iter().skip(4).any(|v| !absent(v)) {
            return Err(failure(1004, "only native local database sources and default options are supported"));
        }
        let sheet = f.sheet.ok_or_else(|| failure(1004, "destination worksheet required"))?;
        let src = self.pivot_source_arg(s, required(&a, 1)?, sheet)?;
        let destination = match required(&a, 2)? {
            Value::Object(r) => {
                let Object::Range(si, r) = self.object(s, *r)? else {
                    return Err(failure(13, "destination Range required"));
                };
                let r = r.rectangle()?;
                if si != sheet || !r.is_single() {
                    return Err(failure(1004, "destination must be one cell on the receiver worksheet"));
                }
                pv::range_text(&s.doc().map_err(|e| failure(91, e.to_string()))?.wb.sheet(si).ok_or_else(|| failure(91, "sheet"))?.name, r)
            }
            v => {
                let t = string(v)?;
                let (si, r) = pv::resolve_source(&s.doc().map_err(|e| failure(91, e.to_string()))?.wb, &t, sheet).map_err(|e| failure(1004, e))?;
                if si != sheet || !r.is_single() {
                    return Err(failure(1004, "destination must be one cell on receiver"));
                }
                pv::range_text(&s.doc().map_err(|e| failure(91, e.to_string()))?.wb.sheet(si).ok_or_else(|| failure(91, "sheet"))?.name, r)
            }
        };
        if s.doc()
            .map_err(|e| failure(91, e.to_string()))?
            .wb
            .sheet(sheet)
            .is_none_or(|sh| sh.is_protected() || sh.pivots.len() >= super::MAX_OBJECTS)
        {
            return Err(failure(1004, "destination protected or pivot budget exceeded"));
        }
        let mut p = json!({"source":src,"destination":destination,"activate":false});
        if !absent(&a[3]) {
            let name = string(&a[3])?;
            if name.is_empty() || name.len() > 255 {
                return Err(failure(1004, "pivot name bounds"));
            }
            p["name"] = json!(name);
        }
        self.objects.reserve(1)?;
        let result = s.execute("insert.pivotTable", p).map_err(|e| failure(1004, e.to_string()))?;
        let id = result["id"].as_u64().ok_or_else(|| failure(1004, "missing pivot identity"))?;
        let pt = s
            .doc()
            .map_err(|e| failure(91, e.to_string()))?
            .wb
            .sheet(sheet)
            .and_then(|sh| sh.pivots.iter().find(|pt| u64::from(pt.id) == id))
            .ok_or_else(|| failure(91, "created pivot missing"))?;
        self.feature(node("PivotTable", sheet, pt))
    }
    fn pivot_range(&mut self, s: &Session, sheet: usize, pt: &PivotTable, member: &str) -> HostResult<Value> {
        let Some(mut area) = pt.last_range else {
            return Ok(Value::Nothing);
        };
        if member != "tablerange2" {
            source(s, sheet, pt)?;
            let out = pv::current_output(&s.doc().map_err(|e| failure(91, e.to_string()))?.wb, pt, sheet).map_err(|e| failure(1004, e))?;
            if member == "tablerange1" {
                area.start.row = area.start.row.saturating_add(out.body_offset);
            } else {
                let mut bounds: Option<RangeRef> = None;
                for (ri, row) in out.grid.iter().enumerate() {
                    for (ci, c) in row.iter().enumerate() {
                        if matches!(c.kind, pv::CellKind::Data | pv::CellKind::Subtotal | pv::CellKind::GrandTotal)
                            && matches!(c.value, gridcraft_engine::core::Value::Number(_) | gridcraft_engine::core::Value::Error(_))
                        {
                            let at = CellRef::new(area.start.row + ri as u32, area.start.col + ci as u32);
                            bounds = Some(bounds.map_or(RangeRef::cell(at), |b| b.union(&RangeRef::cell(at))));
                        }
                    }
                }
                let Some(b) = bounds else {
                    return Ok(Value::Nothing);
                };
                area = b;
            }
        }
        self.intern(Object::Range(sheet, area.into()))
    }
    fn pivot_chart_layout(&mut self, s: &mut Session, f: &Feature, call: &HostCall) -> HostResult<Value> {
        arguments(call, &[])?;
        let sheet = f.sheet.ok_or_else(|| failure(91, "chart sheet"))?;
        let d = s.doc().map_err(|e| failure(91, e.to_string()))?;
        let chart =
            d.wb.sheet(sheet)
                .and_then(|sh| sh.charts.iter().find(|c| u64::from(c.id) == f.key && f.identity.as_ref().is_some_and(|id| id.same(&c.identity))))
                .ok_or_else(|| failure(91, "chart"))?;
        let (sh, name) = chart.pivot.as_deref().and_then(pv::parse_chart_link).ok_or_else(|| failure(1004, "chart has no native PivotLayout"))?;
        let si = d.wb.sheet_index(&sh).ok_or_else(|| failure(91, "pivot sheet"))?;
        let pt = d.wb.sheet(si).and_then(|sh| sh.pivots.iter().find(|pt| pt.name.eq_ignore_ascii_case(&name))).ok_or_else(|| failure(91, "pivot"))?;
        let mut n = node("PivotLayout", si, pt);
        n.span = Some((sheet as u32, chart.id));
        n.ancestors.push(chart.identity.clone());
        self.feature(n)
    }
    fn pivot_layout(&mut self, s: &Session, f: &Feature, pt: &PivotTable, call: &HostCall) -> HostResult<Value> {
        arguments(call, &[])?;
        if !reads(call) {
            return Err(failure(438, "PivotLayout is read-only"));
        }
        let (si, id) = f.span.ok_or_else(|| failure(91, "chart"))?;
        let c = s
            .doc()
            .map_err(|e| failure(91, e.to_string()))?
            .wb
            .sheet(si as usize)
            .and_then(|sh| sh.charts.iter().find(|c| c.id == id && f.ancestors.first().is_some_and(|token| token.same(&c.identity))))
            .ok_or_else(|| failure(91, "chart deleted"))?;
        match call.member.to_ascii_lowercase().as_str() {
            "pivottable" => self.feature(node("PivotTable", f.sheet.ok_or_else(|| failure(91, "pivot"))?, pt)),
            "parent" => {
                let mut n = Feature::new("Chart", Some(si as usize));
                n.key = u64::from(id);
                n.identity = Some(c.identity.clone());
                self.feature(n)
            }
            _ => Err(failure(1004, "no native PivotLayout settings")),
        }
    }
}
pub(crate) fn descriptors() -> Vec<super::MemberDescriptor> {
    [
        ("Worksheet","PivotTables","Function PivotTables(Optional Index As Variant) As Object",super::MemberAvailability::Implemented),
        ("Workbook","PivotCaches","Function PivotCaches() As PivotCaches",super::MemberAvailability::Implemented),
        ("PivotTables","Parent","Property Get Parent() As Worksheet",super::MemberAvailability::Implemented),
        ("PivotTables","Application","Property Get Application() As Application",super::MemberAvailability::Implemented),
        ("PivotTables","Count","Property Get Count() As Long",super::MemberAvailability::Implemented),
        ("PivotTables","Item","Function Item(Index As Variant) As PivotTable",super::MemberAvailability::Implemented),
        ("PivotCaches","Parent","Property Get Parent() As Workbook",super::MemberAvailability::Implemented),
        ("PivotCaches","Application","Property Get Application() As Application",super::MemberAvailability::Implemented),
        ("PivotCaches","Count","Property Get Count() As Long",super::MemberAvailability::Implemented),
        ("PivotCaches","Item","Function Item(Index As Variant) As PivotCache",super::MemberAvailability::Implemented),
        ("PivotTables","Add","Function Add(SourceType As XlPivotTableSourceType, SourceData As Variant, TableDestination As Variant, Optional TableName As Variant, Optional ReadData As Variant, Optional DefaultVersion As Variant) As PivotTable",super::MemberAvailability::Implemented),
        ("PivotCaches","Create","Function Create(SourceType As XlPivotTableSourceType, Optional SourceData As Variant, Optional Version As Variant) As PivotCache",super::MemberAvailability::OutOfScope),
        ("PivotCache","Parent","Property Get Parent() As Workbook",super::MemberAvailability::Implemented),
        ("PivotCache","Application","Property Get Application() As Application",super::MemberAvailability::Implemented),
        ("PivotCache","SourceData","Property SourceData As Variant",super::MemberAvailability::Implemented),
        ("PivotCache","SourceType","Property Get SourceType() As XlPivotTableSourceType",super::MemberAvailability::Implemented),
        ("PivotCache","OLAP","Property Get OLAP() As Boolean",super::MemberAvailability::Implemented),
        ("PivotCache","RecordCount","Property Get RecordCount() As Long",super::MemberAvailability::Implemented),
        ("PivotCache","Index","Property Get Index() As Long",super::MemberAvailability::Implemented),
        ("PivotCache","Refresh","Sub Refresh()",super::MemberAvailability::Implemented),
        ("PivotCache","CreatePivotTable","Function CreatePivotTable() As Object",super::MemberAvailability::OutOfScope),
        ("PivotCache","CreatePivotChart","Function CreatePivotChart() As Object",super::MemberAvailability::OutOfScope),
        ("PivotCache","MakeConnection","Function MakeConnection() As Object",super::MemberAvailability::OutOfScope),
        ("PivotTable","Parent","Property Get Parent() As Worksheet",super::MemberAvailability::Implemented),
        ("PivotTable","Application","Property Get Application() As Application",super::MemberAvailability::Implemented),
        ("PivotTable","Name","Property Get Name() As String",super::MemberAvailability::Implemented),
        ("PivotTable","SourceData","Property SourceData As Variant",super::MemberAvailability::Implemented),
        ("PivotTable","CacheIndex","Property Get CacheIndex() As Long",super::MemberAvailability::Implemented),
        ("PivotTable","PivotCache","Property Get PivotCache() As PivotCache",super::MemberAvailability::Implemented),
        ("PivotTable","RowGrand","Property RowGrand As Boolean",super::MemberAvailability::Implemented),
        ("PivotTable","ColumnGrand","Property ColumnGrand As Boolean",super::MemberAvailability::Implemented),
        ("PivotTable","DisplayFieldCaptions","Property DisplayFieldCaptions As Boolean",super::MemberAvailability::Implemented),
        ("PivotTable","EnableDrilldown","Property EnableDrilldown As Boolean",super::MemberAvailability::Implemented),
        ("PivotTable","TableStyle2","Property TableStyle2 As Variant",super::MemberAvailability::Implemented),
        ("PivotTable","TableRange1","Property Get TableRange1() As Range",super::MemberAvailability::Implemented),
        ("PivotTable","TableRange2","Property Get TableRange2() As Range",super::MemberAvailability::Implemented),
        ("PivotTable","DataBodyRange","Property Get DataBodyRange() As Range",super::MemberAvailability::Implemented),
        ("PivotTable","DataPivotField","Property Get DataPivotField() As PivotField",super::MemberAvailability::Implemented),
        ("PivotTable","Location","Property Get Location() As String",super::MemberAvailability::Implemented),
        ("PivotTable","PivotChart","Property Get PivotChart() As Chart",super::MemberAvailability::Implemented),
        ("PivotTable","PivotFields","Function PivotFields(Optional Index As Variant) As Object",super::MemberAvailability::Implemented),
        ("PivotTable","RowFields","Function RowFields(Optional Index As Variant) As Object",super::MemberAvailability::Implemented),
        ("PivotTable","ColumnFields","Function ColumnFields(Optional Index As Variant) As Object",super::MemberAvailability::Implemented),
        ("PivotTable","PageFields","Function PageFields(Optional Index As Variant) As Object",super::MemberAvailability::Implemented),
        ("PivotTable","DataFields","Function DataFields(Optional Index As Variant) As Object",super::MemberAvailability::Implemented),
        ("PivotTable","HiddenFields","Function HiddenFields(Optional Index As Variant) As Object",super::MemberAvailability::Implemented),
        ("PivotTable","CalculatedFields","Function CalculatedFields(Optional Index As Variant) As Object",super::MemberAvailability::Implemented),
        ("PivotTable","AddDataField","Function AddDataField(Field As Object, Optional Caption As Variant, Optional Function As Variant) As PivotField",super::MemberAvailability::Implemented),
        ("PivotTable","AddFields","Sub AddFields(Optional RowFields As Variant, Optional ColumnFields As Variant, Optional PageFields As Variant, Optional AddToTable As Variant)",super::MemberAvailability::Implemented),
        ("PivotTable","RowAxisLayout","Sub RowAxisLayout(RowLayout As XlLayoutRowType)",super::MemberAvailability::Implemented),
        ("PivotTable","SubtotalLocation","Sub SubtotalLocation(Location As XlSubtotalLocationType)",super::MemberAvailability::Implemented),
        ("PivotTable","RefreshTable","Function RefreshTable() As Boolean",super::MemberAvailability::Implemented),
        ("PivotTable","Update","Sub Update()",super::MemberAvailability::Implemented),
        ("PivotTable","ClearTable","Sub ClearTable()",super::MemberAvailability::Implemented),
        ("PivotTable","ClearAllFilters","Sub ClearAllFilters()",super::MemberAvailability::Implemented),
        ("Range","PivotTable","Property Get PivotTable() As PivotTable",super::MemberAvailability::Implemented),
        ("Range","PivotCell","Property Get PivotCell() As PivotCell",super::MemberAvailability::Implemented),
        ("PivotCell","Parent","Property Get Parent() As Range",super::MemberAvailability::Implemented),
        ("PivotCell","Application","Property Get Application() As Application",super::MemberAvailability::Implemented),
        ("PivotCell","PivotTable","Property Get PivotTable() As PivotTable",super::MemberAvailability::Implemented),
        ("PivotCell","Range","Property Get Range() As Range",super::MemberAvailability::Implemented),
        ("PivotCell","PivotCellType","Property Get PivotCellType() As XlPivotCellType",super::MemberAvailability::Implemented),
        ("Chart","PivotLayout","Property Get PivotLayout() As PivotLayout",super::MemberAvailability::Implemented),
        ("PivotLayout","Parent","Property Get Parent() As Chart",super::MemberAvailability::Implemented),
        ("PivotLayout","Application","Property Get Application() As Application",super::MemberAvailability::Implemented),
        ("PivotLayout","PivotTable","Property Get PivotTable() As PivotTable",super::MemberAvailability::Implemented),
    ].into_iter().map(|(owner,name,signature,availability)|super::MemberDescriptor{owner,name,signature,availability,implementation_symbol:"crates/vba/src/pivot_api.rs::Bridge::pivot_api"}).collect()
}
