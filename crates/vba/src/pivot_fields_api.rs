//! Source-qualified fields and item views use native pivot computation and filtering.
use super::features::{Feature, absent, reads};
use super::objects::{arguments, numeric_index, required, string};
use super::pivot_api::{node, source};
use super::{Bridge, HostCall, HostResult, InvokeKind, Object, Session, Value, failure, values};
use gridcraft_engine::{
    model::{PivotFunc, PivotShowAs, PivotTable},
    pivot as pv,
};
use serde_json::json;
fn field(pt: &PivotTable, sheet: usize, name: &str, part: i32) -> Feature {
    let mut f = node("PivotField", sheet, pt);
    f.text = Some(name.into());
    f.part = part;
    f.ancestors.push(pt.source_identity.clone());
    if part == 2
        && let Some(c) = pt.calculated_fields.iter().find(|c| c.name.eq_ignore_ascii_case(name))
    {
        f.ancestors.push(c.identity.clone());
    }
    f
}
fn data_field(pt: &PivotTable, sheet: usize, index: usize) -> HostResult<Feature> {
    let v = pt.values.get(index).ok_or_else(|| failure(9, "data field index"))?;
    let mut f = field(pt, sheet, &v.source_col, 1);
    f.ancestors.push(v.identity.clone());
    Ok(f)
}
fn area<'a>(pt: &'a PivotTable, name: &str, part: i32) -> Option<(&'a str, usize)> {
    if part != 1 {
        if let Some(i) = pt.rows.iter().position(|x| x.source_col.eq_ignore_ascii_case(name)) {
            return Some(("rows", i));
        }
        if let Some(i) = pt.columns.iter().position(|x| x.source_col.eq_ignore_ascii_case(name)) {
            return Some(("columns", i));
        }
        if let Some(i) = pt.filters.iter().position(|x| x.source_col.eq_ignore_ascii_case(name)) {
            return Some(("filters", i));
        }
    }
    pt.values.iter().position(|x| x.source_col.eq_ignore_ascii_case(name)).map(|i| ("values", i))
}
fn report_filters(pt: &PivotTable) -> Vec<usize> {
    pt.filters
        .iter()
        .enumerate()
        .filter(|(_, f)| !pt.rows.iter().chain(&pt.columns).any(|a| a.source_col.eq_ignore_ascii_case(&f.source_col)))
        .map(|(i, _)| i)
        .collect()
}
fn display_position(pt: &PivotTable, area: &str, index: usize) -> HostResult<usize> {
    Ok(if area == "rows" && pt.values_on_rows && pt.values.len() >= 2 {
        index + usize::from(index >= pt.values_position.unwrap_or(pt.rows.len()))
    } else if area == "filters" {
        report_filters(pt).iter().position(|i| *i == index).ok_or_else(|| failure(91, "report filter removed"))?
    } else {
        index
    })
}
fn field_nodes(pt: &PivotTable, src: &pv::Source, sheet: usize, part: i32) -> HostResult<Vec<Feature>> {
    let out = match part {
        1 => pt.rows.iter().map(|f| field(pt, sheet, &f.source_col, 0)).collect(),
        2 => pt.columns.iter().map(|f| field(pt, sheet, &f.source_col, 0)).collect(),
        3 => pt
            .filters
            .iter()
            .filter(|f| !pt.rows.iter().chain(&pt.columns).any(|x| x.source_col.eq_ignore_ascii_case(&f.source_col)))
            .map(|f| field(pt, sheet, &f.source_col, 0))
            .collect(),
        4 => (0..pt.values.len()).map(|i| data_field(pt, sheet, i)).collect::<HostResult<Vec<_>>>()?,
        6 => pt.calculated_fields.iter().map(|f| field(pt, sheet, &f.name, 2)).collect(),
        _ => {
            let mut list = src
                .headers
                .iter()
                .map(|n| field(pt, sheet, n, 0))
                .chain(pt.calculated_fields.iter().map(|c| field(pt, sheet, &c.name, 2)))
                .collect::<Vec<_>>();
            if part == 5 {
                list.retain(|f| f.text.as_deref().is_some_and(|n| area(pt, n, 0).is_none()));
            } else {
                for i in 0..pt.values.len() {
                    list.push(data_field(pt, sheet, i)?);
                }
            }
            list
        }
    };
    let mut out = out;
    if pt.values.len() >= 2 && matches!(part, 0..=2) {
        let mut f = field(pt, sheet, "Σ Values", 3);
        f.part = 3;
        if part == 0 {
            out.push(f);
        } else if part == 1 && pt.values_on_rows {
            let at = pt.values_position.unwrap_or(out.len()).min(out.len());
            out.insert(at, f);
        } else if part == 2 && !pt.values_on_rows {
            out.push(f);
        }
    }
    Ok(out)
}
fn field_name(pt: &PivotTable, f: &Feature) -> HostResult<String> {
    if f.part == 1 {
        let id = f.ancestors.get(1).ok_or_else(|| failure(91, "data field lifetime"))?;
        Ok(pt.values.iter().find(|v| id.same(&v.identity)).ok_or_else(|| failure(91, "data field deleted"))?.name.clone())
    } else {
        Ok(f.text.clone().ok_or_else(|| failure(91, "field name"))?)
    }
}
fn labels(s: &Session, sheet: usize, pt: &PivotTable, name: &str) -> HostResult<Vec<String>> {
    let src = source(s, sheet, pt)?;
    let col = src.col(name).ok_or_else(|| failure(1004, "calculated fields have no native item records"))?;
    let group = pt.rows.iter().chain(&pt.columns).find(|x| x.source_col.eq_ignore_ascii_case(name)).map(|f| f.date_group).unwrap_or_default();
    let out = pv::distinct_labels(&s.doc().map_err(|e| failure(91, e.to_string()))?.wb, &src, col, group, super::MAX_OBJECTS + 1);
    if out.len() > super::MAX_OBJECTS || out.iter().any(|label| label.len() > 1024) {
        return Err(failure(7, "pivot item collection budget"));
    }
    Ok(out)
}
impl Bridge {
    pub(crate) fn pivot_fields(&mut self, s: &mut Session, f: &Feature, pt: &PivotTable, call: &HostCall) -> HostResult<Value> {
        let sheet = f.sheet.ok_or_else(|| failure(91, "field sheet"))?;
        if !f.ancestors.first().is_some_and(|id| id.same(&pt.source_identity)) {
            return Err(failure(91, "pivot source replaced"));
        }
        let src = source(s, sheet, pt)?;
        let m = call.member.to_ascii_lowercase();
        if f.kind == "PivotFields" || f.kind == "CalculatedFields" {
            return self.pivot_field_collection(s, f, pt, &src, call);
        }
        let name = f.text.as_deref().ok_or_else(|| failure(91, "field"))?;
        if f.part == 3 {
            return self.pivot_values_axis(s, f, pt, call);
        }
        if f.part == 2 {
            if !pt.calculated_fields.iter().any(|c| c.name.eq_ignore_ascii_case(name) && f.ancestors.get(1).is_some_and(|id| id.same(&c.identity))) {
                return Err(failure(91, "calculated field deleted"));
            }
        } else if src.col(name).is_none() && !pt.calculated_fields.iter().any(|c| c.name.eq_ignore_ascii_case(name)) {
            return Err(failure(91, "source field missing"));
        }
        let caption = field_name(pt, f)?;
        if m == "application" && reads(call) {
            arguments(call, &[])?;
            return self.intern(Object::Application);
        }
        if matches!(f.kind, "PivotItems" | "PivotItem") {
            return self.pivot_items(s, f, pt, call);
        }
        let ai = if f.part == 1 {
            pt.values.iter().position(|v| f.ancestors.get(1).is_some_and(|id| id.same(&v.identity))).map(|i| ("values", i))
        } else {
            area(pt, name, f.part)
        };
        if m == "parent" && reads(call) {
            arguments(call, &[])?;
            return self.feature(node("PivotTable", sheet, pt));
        }
        if matches!(m.as_str(), "pivotitems" | "visibleitems" | "hiddenitems") && reads(call) {
            let mut n = f.clone();
            n.kind = "PivotItems";
            n.span = if m == "pivotitems" { None } else { Some((if m == "visibleitems" { 1 } else { 2 }, 0)) };
            if call.arguments.is_empty() {
                return self.feature(n);
            }
            let mut indexed = call.clone();
            indexed.member = "Item".into();
            return self.pivot_items(s, &n, pt, &indexed);
        }
        if m == "autosort" && reads(call) {
            let a = arguments(call, &["Order", "Field", "PivotLine", "CustomSubtotal"])?;
            if a.iter().skip(2).any(|v| !absent(v)) || (!absent(&a[1]) && !string(&a[1])?.eq_ignore_ascii_case(name)) {
                return Err(failure(1004, "native sort supports source item labels only"));
            }
            let order = match numeric_index(required(&a, 0)?)? {
                1 => "asc",
                2 => "desc",
                -4135 => "none",
                _ => return Err(failure(1004, "sort order")),
            };
            self.pivot_command(s, f, "pivot.sort", json!({"field":name,"order":order}))?;
            return Ok(Value::Empty);
        }
        if matches!(m.as_str(), "clearallfilters" | "clearmanualfilter") && reads(call) {
            arguments(call, &[])?;
            self.pivot_command(s, f, "pivot.filter", json!({"field":name,"selected":null}))?;
            return Ok(Value::Empty);
        }
        if m == "delete" && reads(call) {
            arguments(call, &[])?;
            if f.part != 2 {
                return Err(failure(1004, "source fields cannot be deleted; change Orientation"));
            }
            self.pivot_command(s, f, "pivot.calculatedField", json!({"name":name,"delete":true}))?;
            return Ok(Value::Empty);
        }
        if reads(call) {
            arguments(call, &[])?;
            return match m.as_str() {
                "name" | "caption" | "value" | "" => values::vba_text(&caption),
                "sourcename" | "sourcecaption" => values::vba_text(name),
                "iscalculated" => Ok(Value::Boolean(pt.calculated_fields.iter().any(|c| c.name.eq_ignore_ascii_case(name)))),
                "orientation" => Ok(Value::Long(match ai.map(|a| a.0) {
                    Some("rows") => 1,
                    Some("columns") => 2,
                    Some("filters") => 3,
                    Some("values") => 4,
                    _ => 0,
                })),
                "position" => {
                    let (area, index) = ai.ok_or_else(|| failure(1004, "hidden field has no area position"))?;
                    Ok(Value::Long(display_position(pt, area, index)? as i32 + 1))
                }
                "function" | "calculation" | "numberformat" => {
                    let i = ai.filter(|a| a.0 == "values").map(|a| a.1).ok_or_else(|| failure(1004, "use a data field for value settings"))?;
                    let v = pt.values.get(i).ok_or_else(|| failure(91, "data field"))?;
                    match m.as_str() {
                        "function" => Ok(Value::Long(func_number(v.func))),
                        "calculation" => Ok(Value::Long(show_number(v.show_as))),
                        _ => values::vba_text(
                            v.number_format
                                .as_deref()
                                .or_else(|| src.col(name).and_then(|i| src.formats.get(i).map(String::as_str)))
                                .unwrap_or("General"),
                        ),
                    }
                }
                "formula" => values::vba_text(
                    &pt.calculated_fields
                        .iter()
                        .find(|c| c.name.eq_ignore_ascii_case(name))
                        .ok_or_else(|| failure(1004, "not a calculated field"))?
                        .formula,
                ),
                "showdetail" => Ok(Value::Boolean(
                    pt.rows
                        .iter()
                        .chain(&pt.columns)
                        .find(|x| x.source_col.eq_ignore_ascii_case(name))
                        .ok_or_else(|| failure(1004, "field not on an axis"))?
                        .collapsed_items
                        .is_empty(),
                )),
                "currentpage" => {
                    if ai.map(|a| a.0) != Some("filters") {
                        return Err(failure(1004, "field is not a report filter"));
                    }
                    let selected = pt.filters.iter().find(|x| x.source_col.eq_ignore_ascii_case(name)).and_then(|x| x.selected.as_ref());
                    match selected {
                        None => values::vba_text("(All)"),
                        Some(items) if items.len() == 1 => {
                            let mut n = f.clone();
                            n.kind = "PivotItem";
                            n.label = Some(
                                labels(s, sheet, pt, name)?
                                    .into_iter()
                                    .find(|n| items.first().is_some_and(|item| n.eq_ignore_ascii_case(item)))
                                    .ok_or_else(|| failure(91, "selected item removed"))?,
                            );
                            self.feature(n)
                        }
                        _ => Err(failure(1004, "report filter contains multiple selections")),
                    }
                }
                _ => Err(failure(1004, "field member has no native state")),
            };
        }
        if call.kind != InvokeKind::Let {
            return Err(failure(438, "field property requires Let"));
        }
        let a = arguments(call, &["Value"])?;
        let v = required(&a, 0)?;
        let (command, p) = match m.as_str() {
            "orientation" => {
                let to = match numeric_index(v)? {
                    0 => None,
                    1 => Some("rows"),
                    2 => Some("columns"),
                    3 => Some("filters"),
                    4 => Some("values"),
                    _ => return Err(failure(1004, "orientation")),
                };
                match (ai, to) {
                    (Some((from, _)), Some(to)) if from == to => return Ok(Value::Empty),
                    (Some((from, _)), Some(to)) => ("pivot.moveField", json!({"field":if f.part==1{&caption}else{name},"from":from,"to":to})),
                    (None, Some(to)) => ("pivot.addField", json!({"field":name,"area":to})),
                    (Some((from, _)), None) => ("pivot.removeField", json!({"field":if f.part==1{&caption}else{name},"area":from})),
                    (None, None) => return Ok(Value::Empty),
                }
            }
            "position" => {
                let (from, index) = ai.ok_or_else(|| failure(1004, "field not in layout"))?;
                let pos = numeric_index(v)?;
                let pseudo = from == "rows" && pt.values_on_rows && pt.values.len() >= 2;
                let filters = report_filters(pt);
                let len = match from {
                    "rows" => pt.rows.len() + usize::from(pseudo),
                    "columns" => pt.columns.len(),
                    "filters" => filters.len(),
                    _ => pt.values.len(),
                };
                if pos < 1 || pos as usize > len {
                    return Err(failure(1004, "position bounds"));
                }
                let target = pos as usize - 1;
                if display_position(pt, from, index)? == target {
                    return Ok(Value::Empty);
                }
                let mut p = json!({"field":if from=="values"{&caption}else{name},"from":from,"to":from,"position":target});
                if pseudo {
                    let old = pt.values_position.unwrap_or(pt.rows.len());
                    let current = display_position(pt, from, index)?;
                    let mut new = old - usize::from(current < old);
                    if target <= new {
                        new += 1;
                    }
                    p["position"] = json!(target - usize::from(target > new));
                    p["valuesPosition"] = json!(new);
                }
                if from == "filters" {
                    p["position"] = json!(filters.get(target).copied().ok_or_else(|| failure(9, "filter position"))?);
                }
                ("pivot.moveField", p)
            }
            "showdetail" => ("pivot.collapse", json!({"field":name,"collapse":!super::address::boolean(v)?})),
            "currentpage" => {
                if ai.map(|a| a.0) != Some("filters") {
                    return Err(failure(1004, "not a report filter"));
                }
                let label = string(v)?;
                let all = labels(s, sheet, pt, name)?;
                if label == "(All)" {
                    ("pivot.filter", json!({"field":name,"selected":null}))
                } else {
                    let label = all.iter().find(|n| n.eq_ignore_ascii_case(&label)).ok_or_else(|| failure(9, "item not found"))?;
                    ("pivot.filter", json!({"field":name,"selected":[label]}))
                }
            }
            "formula" => {
                if f.part != 2 {
                    return Err(failure(1004, "not a calculated field"));
                }
                let text = string(v)?;
                if text.len() > 65536 {
                    return Err(failure(7, "calculated formula budget"));
                }
                ("pivot.calculatedField", json!({"name":name,"formula":text,"values":false}))
            }
            "name" | "caption" | "function" | "calculation" | "numberformat" => {
                if matches!(m.as_str(), "name" | "caption") && f.part != 1 {
                    return Err(failure(1004, "source/calculated field names are read-only"));
                }
                let (_, i) = ai.filter(|a| a.0 == "values").ok_or_else(|| failure(1004, "only data-field captions and settings are mutable"))?;
                let mut p = json!({"index":i});
                match m.as_str() {
                    "function" => p["func"] = json!(func_name(numeric_index(v)?)?),
                    "calculation" => p["showAs"] = json!(show_name(numeric_index(v)?)?),
                    "numberformat" => p["numberFormat"] = json!(string(v)?),
                    _ => {
                        let caption = string(v)?;
                        if caption.len() > 255 {
                            return Err(failure(1004, "caption bounds"));
                        }
                        p["name"] = json!(caption);
                    }
                }
                ("pivot.valueSettings", p)
            }
            _ => return Err(failure(1004, "no native mutable field option")),
        };
        self.pivot_command(s, f, command, p)?;
        Ok(Value::Empty)
    }
    fn pivot_field_collection(&mut self, s: &mut Session, f: &Feature, pt: &PivotTable, src: &pv::Source, call: &HostCall) -> HostResult<Value> {
        let sheet = f.sheet.ok_or_else(|| failure(91, "sheet"))?;
        let nodes = field_nodes(pt, src, sheet, f.part)?;
        let m = call.member.to_ascii_lowercase();
        if call.kind == InvokeKind::Enumerate {
            arguments(call, &[])?;
            return self.object_array(nodes.into_iter().map(Feature::object).collect());
        }
        if !reads(call) {
            return Err(failure(438, "field collection is read-only"));
        }
        match m.as_str() {
            "parent" => {
                arguments(call, &[])?;
                self.feature(node("PivotTable", sheet, pt))
            }
            "application" => {
                arguments(call, &[])?;
                self.intern(Object::Application)
            }
            "count" => {
                arguments(call, &[])?;
                Ok(Value::Long(nodes.len() as i32))
            }
            "item" | "" | "pivotfields" | "rowfields" | "columnfields" | "pagefields" | "datafields" | "hiddenfields" | "calculatedfields" => {
                let a = arguments(call, &["Index"])?;
                let v = required(&a, 0)?;
                let n = if matches!(v, Value::String(_)) {
                    let name = string(v)?;
                    nodes.iter().find(|f| field_name(pt, f).is_ok_and(|n| n.eq_ignore_ascii_case(&name)))
                } else {
                    usize::try_from(numeric_index(v)?.checked_sub(1).ok_or_else(|| failure(9, "field index"))?).ok().and_then(|i| nodes.get(i))
                }
                .ok_or_else(|| failure(9, "field not found"))?;
                self.feature(n.clone())
            }
            "add" if f.kind == "CalculatedFields" => {
                let a = arguments(call, &["Name", "Formula", "UseStandardFormula"])?;
                if !absent(&a[2]) && !super::address::boolean(&a[2])? {
                    return Err(failure(1004, "localized calculated formulas unavailable"));
                }
                let name = string(required(&a, 0)?)?.trim().to_owned();
                let formula = string(required(&a, 1)?)?;
                if name.len() > 255 || formula.len() > 65536 {
                    return Err(failure(7, "calculated field budget"));
                }
                if pt.calculated_fields.iter().any(|c| c.name.eq_ignore_ascii_case(&name)) {
                    return Err(failure(1004, "calculated field already exists"));
                }
                self.objects.reserve(1)?;
                self.pivot_command(s, f, "pivot.calculatedField", json!({"name":name,"formula":formula,"values":false}))?;
                let (_, pt) = super::pivot_api::lookup(s, f)?;
                self.feature(field(&pt, sheet, &name, 2))
            }
            _ => Err(failure(438, "unsupported field collection member")),
        }
    }
    pub(crate) fn pivot_values_node(&mut self, f: &Feature, pt: &PivotTable) -> HostResult<Value> {
        if pt.values.len() < 2 {
            return Err(failure(1004, "native values axis requires multiple value fields"));
        }
        self.feature(field(pt, f.sheet.ok_or_else(|| failure(91, "sheet"))?, "Σ Values", 3))
    }
    fn pivot_values_axis(&mut self, s: &mut Session, f: &Feature, pt: &PivotTable, call: &HostCall) -> HostResult<Value> {
        if pt.values.len() < 2 {
            return Err(failure(91, "values axis removed"));
        }
        let m = call.member.to_ascii_lowercase();
        let from = if pt.values_on_rows { "rows" } else { "columns" };
        let pos = if pt.values_on_rows { pt.values_position.unwrap_or(pt.rows.len()) } else { pt.columns.len() };
        if reads(call) {
            arguments(call, &[])?;
            return match m.as_str() {
                "name" | "caption" | "value" | "" => values::vba_text("Σ Values"),
                "parent" => self.feature(node("PivotTable", f.sheet.ok_or_else(|| failure(91, "sheet"))?, pt)),
                "application" => self.intern(Object::Application),
                "orientation" => Ok(Value::Long(if pt.values_on_rows { 1 } else { 2 })),
                "position" => Ok(Value::Long(pos as i32 + 1)),
                _ => Err(failure(1004, "no native values-axis setting")),
            };
        }
        if call.kind != InvokeKind::Let {
            return Err(failure(438, "values axis property requires Let"));
        }
        let a = arguments(call, &["Value"])?;
        let n = numeric_index(required(&a, 0)?)?;
        let p = match m.as_str() {
            "orientation" => {
                let to = match n {
                    1 => "rows",
                    2 => "columns",
                    _ => return Err(failure(1004, "values axis belongs to rows or columns")),
                };
                json!({"field":"Σ Values","from":from,"to":to})
            }
            "position" => {
                let len = if pt.values_on_rows { pt.rows.len() + 1 } else { pt.columns.len() + 1 };
                if n < 1 || n as usize > len || !pt.values_on_rows && n as usize != len {
                    return Err(failure(1004, "native values-axis position bounds"));
                }
                json!({"field":"Σ Values","from":from,"to":from,"position":n-1})
            }
            _ => return Err(failure(1004, "values axis field is read-only")),
        };
        self.pivot_command(s, f, "pivot.moveField", p)?;
        Ok(Value::Empty)
    }
    pub(crate) fn pivot_add_fields(&mut self, s: &mut Session, f: &Feature, call: &HostCall) -> HostResult<Value> {
        let a = arguments(call, &["RowFields", "ColumnFields", "PageFields", "AddToTable"])?;
        let mut areas = json!({});
        for (i, key) in [(0, "rows"), (1, "columns"), (2, "filters")] {
            if absent(&a[i]) {
                continue;
            }
            let names = if let Value::Array(array) = &a[i] {
                if array.elements().len() > 32 {
                    return Err(failure(7, "field list budget"));
                }
                array.elements().iter().map(string).collect::<HostResult<Vec<_>>>()?
            } else {
                vec![string(&a[i])?]
            };
            if names.iter().any(|n| n.is_empty() || n.len() > 1024) {
                return Err(failure(1004, "field name bounds"));
            }
            areas[key] = json!(names);
        }
        if areas.as_object().is_none_or(|a| a.is_empty()) {
            return Err(failure(450, "at least one field area is required"));
        }
        self.pivot_command(s, f, "pivot.addField", json!({"areas":areas,"append":!absent(&a[3])&&super::address::boolean(&a[3])?}))?;
        Ok(Value::Empty)
    }
    pub(crate) fn pivot_add_data(&mut self, s: &mut Session, f: &Feature, pt: &PivotTable, call: &HostCall) -> HostResult<Value> {
        let a = arguments(call, &["Field", "Caption", "Function"])?;
        let Value::Object(r) = required(&a, 0)? else {
            return Err(failure(13, "PivotField required"));
        };
        let Object::Feature(field) = self.object(s, *r)? else {
            return Err(failure(13, "PivotField required"));
        };
        if field.kind != "PivotField"
            || !field.identity.as_ref().is_some_and(|id| id.same(&pt.identity))
            || field.sheet != f.sheet
            || !field.ancestors.first().is_some_and(|id| id.same(&pt.source_identity))
        {
            return Err(failure(1004, "field must belong to this PivotTable source"));
        }
        let mut probe = call.clone();
        probe.kind = InvokeKind::Get;
        probe.member = "Name".into();
        probe.arguments.clear();
        self.pivot_fields(s, &field, pt, &probe)?;
        let mut p = json!({"field":field.text,"area":"values"});
        if !absent(&a[1]) {
            let name = string(&a[1])?;
            if name.len() > 255 {
                return Err(failure(1004, "caption bounds"));
            }
            p["name"] = json!(name);
        }
        if !absent(&a[2]) {
            p["func"] = json!(func_name(numeric_index(&a[2])?)?);
        }
        self.objects.reserve(1)?;
        self.pivot_command(s, f, "pivot.addField", p)?;
        let (si, pt) = super::pivot_api::lookup(s, f)?;
        self.feature(data_field(&pt, si, pt.values.len().checked_sub(1).ok_or_else(|| failure(1004, "no data field added"))?)?)
    }
    fn pivot_items(&mut self, s: &mut Session, f: &Feature, pt: &PivotTable, call: &HostCall) -> HostResult<Value> {
        let si = f.sheet.ok_or_else(|| failure(91, "sheet"))?;
        let name = f.text.as_deref().ok_or_else(|| failure(91, "field"))?;
        let mut list = labels(s, si, pt, name)?;
        if f.kind == "PivotItems"
            && let Some((mode, _)) = f.span
        {
            let selected = pt.filters.iter().find(|f| f.source_col.eq_ignore_ascii_case(name)).and_then(|f| f.selected.as_ref());
            list.retain(|label| selected.is_none_or(|sel| sel.iter().any(|v| v.eq_ignore_ascii_case(label))) == (mode == 1));
        }
        let m = call.member.to_ascii_lowercase();
        if f.kind == "PivotItems" {
            if call.kind == InvokeKind::Enumerate {
                arguments(call, &[])?;
                return self.object_array(
                    list.iter()
                        .map(|label| {
                            let mut n = f.clone();
                            n.kind = "PivotItem";
                            n.span = None;
                            n.label = Some(label.clone());
                            n.object()
                        })
                        .collect(),
                );
            }
            if !reads(call) {
                return Err(failure(438, "items collection read-only"));
            }
            match m.as_str() {
                "count" => {
                    arguments(call, &[])?;
                    return Ok(Value::Long(list.len() as i32));
                }
                "parent" => {
                    arguments(call, &[])?;
                    let mut n = f.clone();
                    n.kind = "PivotField";
                    n.span = None;
                    return self.feature(n);
                }
                "application" => {
                    arguments(call, &[])?;
                    return self.intern(Object::Application);
                }
                "item" | "" | "pivotitems" => {
                    let a = arguments(call, &["Index"])?;
                    let v = required(&a, 0)?;
                    let label = if matches!(v, Value::String(_)) {
                        let label = string(v)?;
                        list.iter().find(|n| n.eq_ignore_ascii_case(&label))
                    } else {
                        usize::try_from(numeric_index(v)?.checked_sub(1).ok_or_else(|| failure(9, "item index"))?).ok().and_then(|i| list.get(i))
                    }
                    .ok_or_else(|| failure(9, "item not found"))?;
                    let mut n = f.clone();
                    n.kind = "PivotItem";
                    n.span = None;
                    n.label = Some(label.clone());
                    return self.feature(n);
                }
                _ => return Err(failure(438, "unsupported PivotItems member")),
            }
        }
        let label =
            f.label.as_deref().filter(|label| list.iter().any(|l| l.eq_ignore_ascii_case(label))).ok_or_else(|| failure(91, "pivot item removed"))?;
        let selected = pt.filters.iter().find(|f| f.source_col.eq_ignore_ascii_case(name)).and_then(|f| f.selected.as_ref());
        let visible = selected.is_none_or(|sel| sel.iter().any(|v| v.eq_ignore_ascii_case(label)));
        if reads(call) {
            arguments(call, &[])?;
            return match m.as_str() {
                "parent" => {
                    let mut n = f.clone();
                    n.kind = "PivotField";
                    n.label = None;
                    self.feature(n)
                }
                "name" | "caption" | "value" | "sourcename" | "" => values::vba_text(label),
                "visible" => Ok(Value::Boolean(visible)),
                "recordcount" => {
                    let src = source(s, si, pt)?;
                    let col = src.col(name).ok_or_else(|| failure(91, "source field"))?;
                    let group =
                        pt.rows.iter().chain(&pt.columns).find(|f| f.source_col.eq_ignore_ascii_case(name)).map(|f| f.date_group).unwrap_or_default();
                    let fmt = src.formats.get(col).map(String::as_str).unwrap_or("General");
                    let wb = &s.doc().map_err(|e| failure(91, e.to_string()))?.wb;
                    Ok(Value::Long(
                        src.rows
                            .iter()
                            .filter(|row| row.get(col).is_some_and(|v| pv::item_label(wb, v, group, fmt).eq_ignore_ascii_case(label)))
                            .count() as i32,
                    ))
                }
                "position" => Ok(Value::Long(list.iter().position(|l| l.eq_ignore_ascii_case(label)).ok_or_else(|| failure(91, "item"))? as i32 + 1)),
                "showdetail" => Ok(Value::Boolean(
                    !pt.rows
                        .iter()
                        .chain(&pt.columns)
                        .find(|f| f.source_col.eq_ignore_ascii_case(name))
                        .ok_or_else(|| failure(1004, "item not on axis"))?
                        .collapsed_items
                        .iter()
                        .any(|n| n.eq_ignore_ascii_case(label)),
                )),
                _ => Err(failure(1004, "no independent native pivot item state")),
            };
        }
        if call.kind != InvokeKind::Let {
            return Err(failure(438, "item property requires Let"));
        }
        let a = arguments(call, &["Value"])?;
        let on = super::address::boolean(required(&a, 0)?)?;
        match m.as_str() {
            "visible" => {
                if visible == on {
                    return Ok(Value::Empty);
                }
                let mut next = selected.cloned().unwrap_or_else(|| list.clone());
                if on {
                    next.push(label.into());
                } else {
                    next.retain(|l| !l.eq_ignore_ascii_case(label));
                }
                if next.is_empty() {
                    return Err(failure(1004, "at least one item must remain visible"));
                }
                self.pivot_command(s, f, "pivot.filter", json!({"field":name,"selected":next}))?;
            }
            "showdetail" => {
                self.pivot_command(s, f, "pivot.collapse", json!({"field":name,"items":[label],"collapse":!on}))?;
            }
            _ => return Err(failure(1004, "no mutable native item field")),
        }
        Ok(Value::Empty)
    }
}
pub(crate) fn func_number(f: PivotFunc) -> i32 {
    match f {
        PivotFunc::Sum => -4157,
        PivotFunc::Count => -4112,
        PivotFunc::Average => -4106,
        PivotFunc::Max => -4136,
        PivotFunc::Min => -4139,
        PivotFunc::Product => -4149,
        PivotFunc::CountNumbers => -4113,
        PivotFunc::StdDev => -4155,
        PivotFunc::StdDevP => -4156,
        PivotFunc::Var => -4164,
        PivotFunc::VarP => -4165,
    }
}
fn func_name(n: i32) -> HostResult<&'static str> {
    match n {
        -4157 => Ok("sum"),
        -4112 => Ok("count"),
        -4106 => Ok("average"),
        -4136 => Ok("max"),
        -4139 => Ok("min"),
        -4149 => Ok("product"),
        -4113 => Ok("countNumbers"),
        -4155 => Ok("stdDev"),
        -4156 => Ok("stdDevP"),
        -4164 => Ok("var"),
        -4165 => Ok("varP"),
        _ => Err(failure(1004, "aggregation unavailable")),
    }
}
fn show_number(f: PivotShowAs) -> i32 {
    match f {
        PivotShowAs::Normal => -4143,
        PivotShowAs::PercentOfGrandTotal => 8,
        PivotShowAs::PercentOfColumnTotal => 7,
        PivotShowAs::PercentOfRowTotal => 6,
        PivotShowAs::RunningTotal => 5,
        PivotShowAs::Rank => 15,
    }
}
fn show_name(n: i32) -> HostResult<&'static str> {
    match n {
        -4143 => Ok("normal"),
        8 => Ok("percentOfGrandTotal"),
        7 => Ok("percentOfColumnTotal"),
        6 => Ok("percentOfRowTotal"),
        5 => Ok("runningTotal"),
        15 => Ok("rank"),
        _ => Err(failure(1004, "calculation unavailable")),
    }
}
pub(crate) fn descriptors() -> Vec<super::MemberDescriptor> {
    [
        ("PivotFields", "Parent", "Property Get Parent() As PivotTable", super::MemberAvailability::Implemented),
        ("PivotFields", "Application", "Property Get Application() As Application", super::MemberAvailability::Implemented),
        ("PivotFields", "Count", "Property Get Count() As Long", super::MemberAvailability::Implemented),
        ("PivotFields", "Item", "Function Item(Index As Variant) As PivotField", super::MemberAvailability::Implemented),
        ("CalculatedFields", "Parent", "Property Get Parent() As PivotTable", super::MemberAvailability::Implemented),
        ("CalculatedFields", "Application", "Property Get Application() As Application", super::MemberAvailability::Implemented),
        ("CalculatedFields", "Count", "Property Get Count() As Long", super::MemberAvailability::Implemented),
        ("CalculatedFields", "Item", "Function Item(Index As Variant) As PivotField", super::MemberAvailability::Implemented),
        ("PivotItems", "Parent", "Property Get Parent() As PivotField", super::MemberAvailability::Implemented),
        ("PivotItems", "Application", "Property Get Application() As Application", super::MemberAvailability::Implemented),
        ("PivotItems", "Count", "Property Get Count() As Long", super::MemberAvailability::Implemented),
        ("PivotItems", "Item", "Function Item(Index As Variant) As PivotItem", super::MemberAvailability::Implemented),
        (
            "CalculatedFields",
            "Add",
            "Function Add(Name As String, Formula As String, Optional UseStandardFormula As Variant) As PivotField",
            super::MemberAvailability::Implemented,
        ),
        ("PivotField", "Parent", "Property Get Parent() As PivotTable", super::MemberAvailability::Implemented),
        ("PivotField", "Application", "Property Get Application() As Application", super::MemberAvailability::Implemented),
        ("PivotField", "Name", "Property Name As String", super::MemberAvailability::Implemented),
        ("PivotField", "Caption", "Property Caption As String", super::MemberAvailability::Implemented),
        ("PivotField", "SourceName", "Property Get SourceName() As String", super::MemberAvailability::Implemented),
        ("PivotField", "SourceCaption", "Property Get SourceCaption() As String", super::MemberAvailability::Implemented),
        ("PivotField", "IsCalculated", "Property Get IsCalculated() As Boolean", super::MemberAvailability::Implemented),
        ("PivotField", "Orientation", "Property Orientation As XlPivotFieldOrientation", super::MemberAvailability::Implemented),
        ("PivotField", "Position", "Property Position As Long", super::MemberAvailability::Implemented),
        ("PivotField", "Function", "Property Function As XlConsolidationFunction", super::MemberAvailability::Implemented),
        ("PivotField", "Calculation", "Property Calculation As XlPivotFieldCalculation", super::MemberAvailability::Implemented),
        ("PivotField", "NumberFormat", "Property NumberFormat As String", super::MemberAvailability::Implemented),
        ("PivotField", "Formula", "Property Formula As String", super::MemberAvailability::Implemented),
        ("PivotField", "ShowDetail", "Property ShowDetail As Variant", super::MemberAvailability::Implemented),
        ("PivotField", "CurrentPage", "Property CurrentPage As Variant", super::MemberAvailability::Implemented),
        ("PivotField", "Value", "Property Get Value() As String", super::MemberAvailability::Implemented),
        ("PivotField", "PivotItems", "Function PivotItems(Optional Index As Variant) As Object", super::MemberAvailability::Implemented),
        ("PivotField", "VisibleItems", "Function VisibleItems(Optional Index As Variant) As Object", super::MemberAvailability::Implemented),
        ("PivotField", "HiddenItems", "Function HiddenItems(Optional Index As Variant) As Object", super::MemberAvailability::Implemented),
        (
            "PivotField",
            "AutoSort",
            "Sub AutoSort(Order As Long, Field As String, Optional PivotLine As Variant, Optional CustomSubtotal As Variant)",
            super::MemberAvailability::Implemented,
        ),
        ("PivotField", "ClearAllFilters", "Sub ClearAllFilters()", super::MemberAvailability::Implemented),
        ("PivotField", "ClearManualFilter", "Sub ClearManualFilter()", super::MemberAvailability::Implemented),
        ("PivotField", "Delete", "Sub Delete()", super::MemberAvailability::Implemented),
        ("PivotItem", "Parent", "Property Get Parent() As PivotField", super::MemberAvailability::Implemented),
        ("PivotItem", "Application", "Property Get Application() As Application", super::MemberAvailability::Implemented),
        ("PivotItem", "Name", "Property Get Name() As String", super::MemberAvailability::Implemented),
        ("PivotItem", "Caption", "Property Get Caption() As String", super::MemberAvailability::Implemented),
        ("PivotItem", "Value", "Property Get Value() As String", super::MemberAvailability::Implemented),
        ("PivotItem", "SourceName", "Property Get SourceName() As String", super::MemberAvailability::Implemented),
        ("PivotItem", "Visible", "Property Visible As Boolean", super::MemberAvailability::Implemented),
        ("PivotItem", "ShowDetail", "Property ShowDetail As Variant", super::MemberAvailability::Implemented),
        ("PivotItem", "Position", "Property Get Position() As Long", super::MemberAvailability::Implemented),
        ("PivotItem", "RecordCount", "Property Get RecordCount() As Long", super::MemberAvailability::Implemented),
    ]
    .into_iter()
    .map(|(owner, name, signature, availability)| super::MemberDescriptor {
        owner,
        name,
        signature,
        availability,
        implementation_symbol: "crates/vba/src/pivot_fields_api.rs::Bridge::pivot_fields",
    })
    .collect()
}
