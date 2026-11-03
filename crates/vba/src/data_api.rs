//! OM06: persistent VBA sort configuration and native data-command projections.
use super::features::{Feature, absent, reads};
use super::objects::{arguments, numeric_index, required, string};
use super::{Bridge, HostCall, HostResult, InvokeKind, Object, ObjectRef, Session, Value, failure, values};
use gridcraft_engine::{
    core::RangeRef,
    model::{AutoFilter, CfOperator, ErrorStyle, FilterCriterion, Validation, ValidationKind, sheet::SheetIdentity},
};
use serde_json::{Value as Json, json};
#[derive(Clone)]
pub(crate) struct SortField {
    id: u64,
    key: ObjectRef,
    order: i32,
    custom: Option<Vec<String>>,
}
#[derive(Clone)]
pub(crate) struct SortState {
    doc: u64,
    identity: SheetIdentity,
    table: Option<SheetIdentity>,
    range: Option<ObjectRef>,
    fields: Vec<SortField>,
    next: u64,
    header: i32,
    case: bool,
    orientation: i32,
}
fn validate_table_view(session: &Session, sheet: usize, key: u64, f: &Feature) -> HostResult<()> {
    if key == 0 {
        return Ok(());
    }
    let table = session
        .doc()
        .map_err(|e| failure(91, e.to_string()))?
        .wb
        .sheet(sheet)
        .and_then(|sh| sh.tables.iter().find(|t| u64::from(t.id) == key))
        .ok_or_else(|| failure(91, "table was deleted"))?;
    if !f.identity.as_ref().is_some_and(|id| id.same(&table.identity)) {
        return Err(failure(91, "table view identity is stale"));
    }
    Ok(())
}
fn order(value: &Value, default: i32) -> HostResult<i32> {
    let n = if absent(value) { default } else { numeric_index(value)? };
    if !matches!(n, 1 | 2) {
        return Err(failure(1004, "invalid sort order/orientation"));
    }
    Ok(n)
}
pub(crate) fn text(value: &Value) -> HostResult<String> {
    if matches!(value, Value::String(_)) {
        string(value)
    } else {
        match values::content(value, false)? {
            gridcraft_engine::CellContent::Value(v) => Ok(v.display()),
            _ => Err(failure(13, "scalar value required")),
        }
    }
}
fn dv_type(n: i32) -> HostResult<&'static str> {
    match n {
        0 => Ok("any"),
        1 => Ok("whole"),
        2 => Ok("decimal"),
        3 => Ok("list"),
        4 => Ok("date"),
        5 => Ok("time"),
        6 => Ok("textLength"),
        7 => Ok("custom"),
        _ => Err(failure(1004, "unsupported validation type")),
    }
}
fn op(n: i32) -> HostResult<&'static str> {
    match n {
        1 => Ok("between"),
        2 => Ok("notBetween"),
        3 => Ok("equal"),
        4 => Ok("notEqual"),
        5 => Ok("greater"),
        6 => Ok("less"),
        7 => Ok("greaterOrEqual"),
        8 => Ok("lessOrEqual"),
        _ => Err(failure(1004, "invalid comparison operator")),
    }
}
fn op_number(op: CfOperator) -> i32 {
    match op {
        CfOperator::Between => 1,
        CfOperator::NotBetween => 2,
        CfOperator::Equal => 3,
        CfOperator::NotEqual => 4,
        CfOperator::Greater => 5,
        CfOperator::Less => 6,
        CfOperator::GreaterOrEqual => 7,
        CfOperator::LessOrEqual => 8,
    }
}
impl Bridge {
    pub(crate) fn sort_state(&mut self, session: &Session, sheet: usize, table: u64) -> HostResult<usize> {
        let doc = session.doc().map_err(|e| failure(91, e.to_string()))?;
        let sh = doc.wb.sheet(sheet).ok_or_else(|| failure(91, "worksheet deleted"))?;
        let table_identity = if table == 0 {
            None
        } else {
            Some(sh.tables.iter().find(|t| u64::from(t.id) == table).ok_or_else(|| failure(91, "sort table deleted"))?.identity.clone())
        };
        if let Some(i) = self.sorts.iter().position(|s| {
            s.doc == doc.uid
                && s.identity.same(&sh.identity)
                && match (&s.table, &table_identity) {
                    (None, None) => true,
                    (Some(a), Some(b)) => a.same(b),
                    _ => false,
                }
        }) {
            return Ok(i);
        }
        if self.sorts.len() >= super::MAX_OBJECTS {
            return Err(failure(7, "sort configuration budget"));
        }
        self.sorts.push(SortState {
            doc: doc.uid,
            identity: sh.identity.clone(),
            table: table_identity,
            range: None,
            fields: vec![],
            next: 1,
            header: if table == 0 { 2 } else { 1 },
            case: false,
            orientation: 1,
        });
        Ok(self.sorts.len() - 1)
    }
    pub(crate) fn data_api(&mut self, session: &mut Session, object: &Object, call: &HostCall) -> Option<HostResult<Value>> {
        let member = call.member.to_ascii_lowercase();
        if let Object::Sheet(sheet) = object {
            if member == "sort" && reads(call) {
                return Some((|| {
                    arguments(call, &[])?;
                    self.sort_state(session, *sheet, 0)?;
                    self.feature(Feature::new("Sort", Some(*sheet)))
                })());
            }
            if matches!(member.as_str(), "autofilter" | "filtermode" | "autofiltermode" | "showalldata") {
                return Some(self.sheet_filter(session, *sheet, call, &member));
            }
        }
        if let Object::Range(sheet, range) = object {
            if member == "validation" && reads(call) {
                return Some((|| {
                    arguments(call, &[])?;
                    let mut f = Feature::new("Validation", Some(*sheet));
                    f.range = Some(range.clone());
                    self.feature(f)
                })());
            }
            if matches!(
                member.as_str(),
                "autofilter" | "sort" | "removeduplicates" | "texttocolumns" | "group" | "ungroup" | "showdetail" | "subtotal"
            ) {
                return Some(self.range_data(session, *sheet, range, call, &member));
            }
        }
        if let Object::Feature(f) = object {
            if matches!(f.kind, "Sort" | "SortFields" | "SortField") {
                return Some(self.sort_dispatch(session, f, call));
            }
            if f.kind == "Validation" {
                return Some(self.validation_api(session, f, call));
            }
            if matches!(f.kind, "AutoFilter" | "Filters" | "Filter") {
                return Some(self.filter_dispatch(session, f, call));
            }
        }
        None
    }
    fn sort_dispatch(&mut self, session: &mut Session, f: &Feature, call: &HostCall) -> HostResult<Value> {
        let sheet = f.sheet.ok_or_else(|| failure(91, "sort owner gone"))?;
        let table_key = if f.kind == "SortField" { u64::from(f.part as u32) } else { f.key };
        validate_table_view(session, sheet, table_key, f)?;
        let index = self.sort_state(session, sheet, table_key)?;
        let member = call.member.to_ascii_lowercase();
        if member == "application" && reads(call) {
            arguments(call, &[])?;
            return self.intern(Object::Application);
        }
        if member == "parent" && reads(call) {
            arguments(call, &[])?;
            if f.kind == "Sort" {
                if table_key == 0 {
                    return self.intern(Object::Sheet(sheet));
                }
                let table = session
                    .doc()
                    .map_err(|e| failure(91, e.to_string()))?
                    .wb
                    .sheet(sheet)
                    .and_then(|sh| sh.tables.iter().find(|t| u64::from(t.id) == table_key))
                    .ok_or_else(|| failure(91, "table deleted"))?;
                return self.feature(Self::table_node(sheet, table));
            }
            let mut parent = f.clone();
            parent.kind = if f.kind == "SortFields" { "Sort" } else { "SortFields" };
            parent.key = table_key;
            parent.part = 0;
            return self.feature(parent);
        }
        match f.kind {
            "Sort" => match member.as_str() {
                "sortfields" if reads(call) => {
                    let mut node = Feature::new("SortFields", Some(sheet));
                    node.identity = self.sorts[index].table.clone();
                    node.key = table_key;
                    if call.arguments.is_empty() {
                        self.feature(node)
                    } else {
                        let mut call = call.clone();
                        call.member = "Item".into();
                        self.sort_dispatch(session, &node, &call)
                    }
                }
                "setrange" => {
                    if table_key != 0 {
                        return Err(failure(1004, "table Sort owns its native range"));
                    }
                    let args = arguments(call, &["Rng"])?;
                    let reference = self.range_argument(session, sheet, required(&args, 0)?)?;
                    let Object::Range(_, range) = self.object(session, reference)? else {
                        return Err(failure(13, "Range required"));
                    };
                    values::dimensions(range.rectangle()?)?;
                    self.sorts[index].range = Some(reference);
                    Ok(Value::Empty)
                }
                "apply" => {
                    arguments(call, &[])?;
                    let state = self.sorts[index].clone();
                    self.apply_sort(session, sheet, &state)
                }
                "header" | "matchcase" | "orientation" => {
                    if reads(call) {
                        arguments(call, &[])?;
                        let state = &self.sorts[index];
                        return Ok(match member.as_str() {
                            "header" => Value::Long(state.header),
                            "orientation" => Value::Long(state.orientation),
                            _ => Value::Boolean(state.case),
                        });
                    }
                    if call.kind != InvokeKind::Let {
                        return Err(failure(13, "sort assignment requires Let"));
                    }
                    let args = arguments(call, &["Value"])?;
                    let value = required(&args, 0)?;
                    match member.as_str() {
                        "header" => {
                            let n = numeric_index(value)?;
                            if !(0..=2).contains(&n) {
                                return Err(failure(1004, "invalid Header"));
                            }
                            self.sorts[index].header = n;
                        }
                        "orientation" => self.sorts[index].orientation = order(value, 1)?,
                        _ => self.sorts[index].case = super::address::boolean(value)?,
                    };
                    Ok(Value::Empty)
                }
                "sortmethod" => Err(failure(1004, "GridCraft's Unicode comparison does not implement Pinyin or Stroke collation")),
                _ => Err(failure(438, "unsupported Sort member")),
            },
            "SortFields" => {
                if call.kind == InvokeKind::Enumerate {
                    arguments(call, &[])?;
                    return self.object_array(
                        self.sorts[index]
                            .fields
                            .iter()
                            .map(|field| {
                                let mut f = Feature::new("SortField", Some(sheet));
                                f.key = field.id;
                                f.part = table_key as i32;
                                f.identity = self.sorts[index].table.clone();
                                f.object()
                            })
                            .collect(),
                    );
                }
                match member.as_str() {
                    "count" if reads(call) => {
                        arguments(call, &[])?;
                        Ok(Value::Long(self.sorts[index].fields.len() as i32))
                    }
                    "clear" => {
                        arguments(call, &[])?;
                        self.sorts[index].fields.clear();
                        Ok(Value::Empty)
                    }
                    "item" | "" if reads(call) => {
                        let args = arguments(call, &["Index"])?;
                        let i = usize::try_from(numeric_index(required(&args, 0)?)?.checked_sub(1).ok_or_else(|| failure(9, "sort field index"))?)
                            .map_err(|_| failure(9, "sort field index"))?;
                        let field = self.sorts[index].fields.get(i).ok_or_else(|| failure(9, "sort field index"))?;
                        let mut node = Feature::new("SortField", Some(sheet));
                        node.key = field.id;
                        node.part = table_key as i32;
                        node.identity = self.sorts[index].table.clone();
                        self.feature(node)
                    }
                    "add" | "add2" => {
                        let args = arguments(
                            call,
                            if member == "add2" {
                                &["Key", "SortOn", "Order", "CustomOrder", "DataOption", "SubField"]
                            } else {
                                &["Key", "SortOn", "Order", "CustomOrder", "DataOption"]
                            },
                        )?;
                        if !absent(&args[1]) && numeric_index(&args[1])? != 0
                            || !absent(&args[4]) && numeric_index(&args[4])? != 0
                            || args.get(5).is_some_and(|v| !absent(v))
                        {
                            return Err(failure(1004, "only value sorting with normal data coercion is supported"));
                        }
                        let key = self.range_argument(session, sheet, required(&args, 0)?)?;
                        let order = order(&args[2], 1)?;
                        let custom = if absent(&args[3]) {
                            None
                        } else {
                            let text = string(&args[3])?;
                            if text.len() > 4096 {
                                return Err(failure(7, "custom order budget"));
                            }
                            Some(text.split(',').map(str::to_string).collect())
                        };
                        self.objects.reserve(1)?;
                        let state = &mut self.sorts[index];
                        if state.fields.len() >= 64 {
                            return Err(failure(7, "sort key budget"));
                        }
                        let id = state.next;
                        state.next = state.next.checked_add(1).ok_or_else(|| failure(7, "sort identity exhausted"))?;
                        state.fields.push(SortField { id, key, order, custom });
                        let mut node = Feature::new("SortField", Some(sheet));
                        node.key = id;
                        node.part = table_key as i32;
                        node.identity = self.sorts[index].table.clone();
                        self.feature(node)
                    }
                    _ => Err(failure(438, "unsupported SortFields member")),
                }
            }
            _ => {
                let i = self.sorts[index].fields.iter().position(|field| field.id == f.key).ok_or_else(|| failure(91, "sort field was removed"))?;
                if member == "delete" {
                    arguments(call, &[])?;
                    self.sorts[index].fields.remove(i);
                    return Ok(Value::Empty);
                }
                if reads(call) {
                    arguments(call, &[])?;
                    let field = &self.sorts[index].fields[i];
                    return match member.as_str() {
                        "key" => Ok(Value::Object(field.key)),
                        "order" => Ok(Value::Long(field.order)),
                        "sorton" | "dataoption" => Ok(Value::Long(0)),
                        "customorder" => values::vba_text(&field.custom.as_ref().map(|v| v.join(",")).unwrap_or_default()),
                        _ => Err(failure(438, "unsupported SortField member")),
                    };
                }
                if call.kind != InvokeKind::Let {
                    return Err(failure(13, "sort field assignment requires Let"));
                }
                let args = arguments(call, &["Value"])?;
                let value = required(&args, 0)?;
                match member.as_str() {
                    "order" => self.sorts[index].fields[i].order = order(value, 1)?,
                    "dataoption" | "sorton" => {
                        if numeric_index(value)? != 0 {
                            return Err(failure(1004, "unsupported sort mode"));
                        }
                    }
                    _ => return Err(failure(438, "unsupported SortField assignment")),
                };
                Ok(Value::Empty)
            }
        }
    }
    pub(crate) fn range_argument(&mut self, session: &Session, sheet: usize, value: &Value) -> HostResult<ObjectRef> {
        let reference = if let Value::Object(reference) = value {
            *reference
        } else {
            let (owner, range) = Self::resolve_text(session, sheet, &string(value)?, None, 0, &mut 0)?;
            let Value::Object(reference) = self.intern(Object::Range(owner, range))? else {
                return Err(failure(13, "Range required"));
            };
            reference
        };
        let Object::Range(owner, range) = self.object(session, reference)? else {
            return Err(failure(13, "Range required"));
        };
        if owner != sheet {
            return Err(failure(1004, "range belongs to another worksheet"));
        }
        range.rectangle()?;
        Ok(reference)
    }
    fn apply_sort(&mut self, session: &mut Session, sheet: usize, state: &SortState) -> HostResult<Value> {
        let range = if let Some(identity) = &state.table {
            let table = session
                .doc()
                .map_err(|e| failure(91, e.to_string()))?
                .wb
                .sheet(sheet)
                .and_then(|sh| sh.tables.iter().find(|t| t.identity.same(identity)))
                .ok_or_else(|| failure(91, "sort table deleted"))?;
            table.range.into()
        } else {
            let reference = state.range.ok_or_else(|| failure(1004, "Sort.SetRange is required"))?;
            let Object::Range(owner, range) = self.object(session, reference)? else {
                return Err(failure(91, "sort range deleted"));
            };
            if owner != sheet {
                return Err(failure(1004, "sort owner mismatch"));
            }
            range
        };
        let rect = range.rectangle()?;
        values::dimensions(rect)?;
        if state.fields.is_empty() {
            return Err(failure(1004, "Sort requires at least one key; no dialog is opened"));
        }
        let mut keys = Vec::new();
        for field in &state.fields {
            let Object::Range(owner, key) = self.object(session, field.key)? else {
                return Err(failure(91, "sort key deleted"));
            };
            if owner != sheet {
                return Err(failure(1004, "sort key owner mismatch"));
            }
            let key = key.rectangle()?;
            let mut params = if state.orientation == 1 {
                if !(rect.start.col..=rect.end.col).contains(&key.start.col) {
                    return Err(failure(1004, "sort key outside range"));
                }
                json!({"column":gridcraft_engine::core::col_to_letters(key.start.col)})
            } else {
                if !(rect.start.row..=rect.end.row).contains(&key.start.row) {
                    return Err(failure(1004, "sort key outside range"));
                }
                json!({"row":key.start.row+1})
            };
            params["order"] = json!(if field.order == 1 { "asc" } else { "desc" });
            if let Some(custom) = &field.custom {
                params["customList"] = json!(custom);
            }
            keys.push(params);
        }
        let mut params =
            json!({"sheet":sheet,"range":rect.a1(),"keys":keys,"orientation":if state.orientation==1{"rows"}else{"columns"},"matchCase":state.case});
        if state.header != 0 {
            params["header"] = json!(state.header == 1);
        }
        let changes = self.prepare_change(session, sheet, range)?;
        session.execute("data.sort", params).map_err(|e| failure(1004, e.to_string()))?;
        self.queue_changes(changes);
        Ok(Value::Empty)
    }
    fn sheet_filter(&mut self, session: &mut Session, sheet: usize, call: &HostCall, member: &str) -> HostResult<Value> {
        let sh = session.doc().map_err(|e| failure(91, e.to_string()))?.wb.sheet(sheet).ok_or_else(|| failure(91, "worksheet deleted"))?;
        if reads(call) {
            arguments(call, &[])?;
            return match member {
                "autofilter" => {
                    if sh.autofilter.is_some() {
                        self.feature(Feature::new("AutoFilter", Some(sheet)))
                    } else {
                        Ok(Value::Nothing)
                    }
                }
                "autofiltermode" => Ok(Value::Boolean(sh.autofilter.is_some())),
                "filtermode" => Ok(Value::Boolean(
                    sh.autofilter.as_ref().is_some_and(|af| !af.criteria.is_empty()) || sh.tables.iter().any(|t| !t.filter.is_empty()),
                )),
                "showalldata" => {
                    session.execute("data.clearFilter", json!({"sheet":sheet})).map_err(|e| failure(1004, e.to_string()))?;
                    Ok(Value::Empty)
                }
                _ => Err(failure(438, "unsupported filter member")),
            };
        }
        if member == "autofiltermode" && call.kind == InvokeKind::Let {
            let args = arguments(call, &["Value"])?;
            if super::address::boolean(required(&args, 0)?)? {
                return Err(failure(1004, "use Range.AutoFilter to enable filtering"));
            }
            if let Some(af) = &sh.autofilter {
                session.execute("data.filter", json!({"sheet":sheet,"range":af.range.a1()})).map_err(|e| failure(1004, e.to_string()))?;
            }
            return Ok(Value::Empty);
        }
        Err(failure(438, "unsupported filter assignment"))
    }
    fn range_data(&mut self, session: &mut Session, sheet: usize, range: &super::ranges::Range, call: &HostCall, member: &str) -> HostResult<Value> {
        let rect = range.rectangle()?;
        if !matches!(member, "group" | "ungroup" | "showdetail") {
            values::dimensions(rect)?;
        }
        if member == "autofilter" {
            return self.autofilter_range(session, sheet, rect, call);
        }
        if member == "sort" {
            let args = arguments(
                call,
                &[
                    "Key1",
                    "Order1",
                    "Key2",
                    "Type",
                    "Order2",
                    "Key3",
                    "Order3",
                    "Header",
                    "OrderCustom",
                    "MatchCase",
                    "Orientation",
                    "SortMethod",
                    "DataOption1",
                    "DataOption2",
                    "DataOption3",
                ],
            )?;
            if [3, 8, 11, 12, 13, 14].iter().any(|i| !absent(&args[*i])) {
                return Err(failure(1004, "this Range.Sort option is unavailable"));
            }
            let Value::Object(reference) = self.intern(Object::Range(sheet, range.clone()))? else {
                return Err(failure(13, "Range required"));
            };
            let mut state = SortState {
                doc: self.context,
                identity: Default::default(),
                table: None,
                range: Some(reference),
                fields: vec![],
                next: 1,
                header: if absent(&args[7]) { 2 } else { numeric_index(&args[7])? },
                case: !absent(&args[9]) && super::address::boolean(&args[9])?,
                orientation: order(&args[10], 1)?,
            };
            if !(0..=2).contains(&state.header) {
                return Err(failure(1004, "invalid Header"));
            }
            required(&args, 0)?;
            for (key, ord) in [(0, 1), (2, 4), (5, 6)] {
                if !absent(&args[key]) {
                    let reference = self.range_argument(session, sheet, &args[key])?;
                    state.fields.push(SortField { id: key as u64, key: reference, order: order(&args[ord], 1)?, custom: None });
                }
            }
            return self.apply_sort(session, sheet, &state);
        }
        if matches!(member, "group" | "ungroup" | "showdetail")
            && session.doc().map_err(|e| failure(91, e.to_string()))?.wb.sheet(sheet).is_some_and(|sh| sh.is_protected())
        {
            return Err(failure(1004, "outline changes require unprotected worksheet"));
        }
        let mut params = json!({"sheet":sheet,"range":rect.a1()});
        let command = match member {
            "removeduplicates" => {
                let args = arguments(call, &["Columns", "Header"])?;
                let cols = match required(&args, 0)? {
                    Value::Array(a) => a.elements().to_vec(),
                    value => vec![value.clone()],
                };
                if cols.is_empty() || cols.len() > rect.width() as usize {
                    return Err(failure(1004, "invalid duplicate columns"));
                }
                let mut columns = Vec::new();
                for v in cols {
                    let n = numeric_index(&v)?;
                    if n < 1 || n as u32 > rect.width() {
                        return Err(failure(1004, "duplicate column outside range"));
                    }
                    columns.push(gridcraft_engine::core::col_to_letters(rect.start.col + n as u32 - 1));
                }
                params["columns"] = json!(columns);
                if !absent(&args[1]) {
                    let h = numeric_index(&args[1])?;
                    if !(0..=2).contains(&h) {
                        return Err(failure(1004, "invalid Header"));
                    }
                    if h != 0 {
                        params["header"] = json!(h == 1);
                    }
                }
                "data.removeDuplicates"
            }
            "group" | "ungroup" => {
                let args = arguments(call, &["Start", "End", "By", "Periods"])?;
                if args.iter().any(|v| !absent(v)) {
                    return Err(failure(1004, "pivot grouping options are unavailable"));
                }
                if range.unit == super::ranges::Unit::Columns {
                    params["cols"] = json!(format!(
                        "{}:{}",
                        gridcraft_engine::core::col_to_letters(rect.start.col),
                        gridcraft_engine::core::col_to_letters(rect.end.col)
                    ));
                } else {
                    params["rows"] = json!(format!("{}:{}", rect.start.row + 1, rect.end.row + 1));
                }
                if member == "group" { "data.group" } else { "data.ungroup" }
            }
            "showdetail" => {
                if call.kind != InvokeKind::Let {
                    return Err(failure(438, "ShowDetail read is unavailable"));
                }
                let args = arguments(call, &["Value"])?;
                params[if range.unit == super::ranges::Unit::Columns { "cols" } else { "rows" }] =
                    json!(if range.unit == super::ranges::Unit::Columns {
                        format!("{}:{}", gridcraft_engine::core::col_to_letters(rect.start.col), gridcraft_engine::core::col_to_letters(rect.end.col))
                    } else {
                        format!("{}:{}", rect.start.row + 1, rect.end.row + 1)
                    });
                if super::address::boolean(required(&args, 0)?)? { "data.showDetail" } else { "data.hideDetail" }
            }
            "texttocolumns" => self.text_to_columns_params(session, sheet, rect, call, &mut params)?,
            "subtotal" => {
                let args = arguments(call, &["GroupBy", "Function", "TotalList", "Replace", "PageBreaks", "SummaryBelowData"])?;
                if !absent(&args[4]) && super::address::boolean(&args[4])? || !absent(&args[5]) && numeric_index(&args[5])? != 1 {
                    return Err(failure(1004, "page breaks or summaries above data are unavailable"));
                }
                let replace = absent(&args[3]) || super::address::boolean(&args[3])?;
                if replace
                    && session.doc().map_err(|e| failure(91, e.to_string()))?.wb.sheet(sheet).is_some_and(|sh| {
                        sh.cells.iter_range(rect).any(|(_, c)| c.formula.as_ref().is_some_and(|f| f.text.to_ascii_uppercase().contains("SUBTOTAL(")))
                    })
                {
                    return Err(failure(1004, "replacing existing subtotals is unavailable"));
                }
                let group = numeric_index(required(&args, 0)?)?;
                if group < 1 || group as u32 > rect.width() {
                    return Err(failure(1004, "GroupBy outside range"));
                }
                params["groupBy"] = json!(gridcraft_engine::core::col_to_letters(rect.start.col + group as u32 - 1));
                params["function"] = json!(match numeric_index(required(&args, 1)?)? {
                    -4157 => "sum",
                    -4112 => "count",
                    -4106 => "average",
                    -4136 => "max",
                    -4139 => "min",
                    -4149 => "product",
                    _ => return Err(failure(1004, "unsupported subtotal function")),
                });
                let cols = match required(&args, 2)? {
                    Value::Array(a) => a.elements().to_vec(),
                    v => vec![v.clone()],
                };
                let mut columns = Vec::new();
                for v in cols {
                    let i = numeric_index(&v)?;
                    if i < 1 || i as u32 > rect.width() {
                        return Err(failure(1004, "subtotal column outside range"));
                    }
                    columns.push(gridcraft_engine::core::col_to_letters(rect.start.col + i as u32 - 1));
                }
                if columns.is_empty() {
                    return Err(failure(1004, "TotalList cannot be empty"));
                }
                params["columns"] = json!(columns);
                "data.subtotal"
            }
            _ => return Err(failure(438, "unsupported Range data member")),
        };
        let changes = if matches!(member, "texttocolumns" | "subtotal") {
            match session.preview_data_mutation(command, &params).map_err(|e| failure(1004, e.to_string()))? {
                Some((owner, range)) => {
                    values::dimensions(range)?;
                    self.prepare_change(session, owner, range.into())?
                }
                None => super::PreparedChanges::default(),
            }
        } else if member == "removeduplicates" {
            self.prepare_change(session, sheet, range.clone())?
        } else {
            super::PreparedChanges::default()
        };
        session.execute(command, params).map_err(|e| failure(1004, e.to_string()))?;
        self.queue_changes(changes);
        Ok(Value::Empty)
    }
    fn text_to_columns_params(
        &mut self,
        session: &Session,
        sheet: usize,
        rect: RangeRef,
        call: &HostCall,
        params: &mut Json,
    ) -> HostResult<&'static str> {
        let args = arguments(
            call,
            &[
                "Destination",
                "DataType",
                "TextQualifier",
                "ConsecutiveDelimiter",
                "Tab",
                "Semicolon",
                "Comma",
                "Space",
                "Other",
                "OtherChar",
                "FieldInfo",
                "DecimalSeparator",
                "ThousandsSeparator",
                "TrailingMinusNumbers",
            ],
        )?;
        let data_type = if absent(&args[1]) { 1 } else { numeric_index(&args[1])? };
        if args.iter().skip(11).any(|v| !absent(v)) || !matches!(data_type, 1 | 2) {
            return Err(failure(1004, "this TextToColumns conversion option is unavailable"));
        }
        let mut breaks = Vec::new();
        if !absent(&args[10]) {
            let Value::Array(fields) = &args[10] else {
                return Err(failure(13, "FieldInfo requires an array of field pairs"));
            };
            if fields.elements().len() > 256 {
                return Err(failure(7, "FieldInfo budget"));
            }
            let pairs: Vec<Vec<&Value>> =
                if fields.bounds().len() == 2 && fields.bounds().get(1).is_some_and(|(a, b)| i64::from(*b) - i64::from(*a) + 1 == 2) {
                    let rows = fields.elements().len() / 2;
                    (0..rows).map(|i| vec![&fields.elements()[i], &fields.elements()[i + rows]]).collect()
                } else {
                    fields
                        .elements()
                        .iter()
                        .map(|pair| {
                            if let Value::Array(pair) = pair {
                                Ok(pair.elements().iter().collect())
                            } else {
                                Err(failure(13, "FieldInfo requires pairs or a two-column array"))
                            }
                        })
                        .collect::<HostResult<_>>()?
                };
            for pair in pairs {
                let [position, format] = pair.as_slice() else {
                    return Err(failure(13, "FieldInfo pair requires two elements"));
                };
                let n = numeric_index(position)?;
                if n < if data_type == 1 { 1 } else { 0 } || numeric_index(format)? != 1 {
                    return Err(failure(1004, "only native general field conversion is available"));
                }
                if data_type == 2 {
                    breaks.push(n);
                }
            }
        }
        if data_type == 2 {
            params["fixedWidths"] = json!(breaks);
        }
        if rect.width() != 1 {
            return Err(failure(1004, "TextToColumns requires one source column"));
        }
        if !absent(&args[0]) {
            let reference = self.range_argument(session, sheet, &args[0])?;
            let Object::Range(_, r) = self.object(session, reference)? else {
                return Err(failure(13, "Range required"));
            };
            params["destination"] = json!(r.first()?.start.a1());
        }
        let mut delimiters = Vec::new();
        for (i, ch) in [(4, "\t"), (5, ";"), (6, ","), (7, " ")] {
            if absent(&args[i]) {
                if i == 4 {
                    delimiters.push(ch);
                }
            } else if super::address::boolean(&args[i])? {
                delimiters.push(ch);
            }
        }
        if !absent(&args[8]) && super::address::boolean(&args[8])? {
            let ch = string(required(&args, 9)?)?;
            if ch.chars().count() != 1 {
                return Err(failure(1004, "OtherChar must be one character"));
            }
            params["other"] = json!(ch);
        }
        params["delimiters"] = json!(delimiters);
        params["treatConsecutive"] = json!(!absent(&args[3]) && super::address::boolean(&args[3])?);
        if !absent(&args[2]) {
            params["textQualifier"] = json!(match numeric_index(&args[2])? {
                1 => "\"",
                2 => "'",
                -4142 => "\u{0}",
                _ => return Err(failure(1004, "invalid TextQualifier")),
            });
        }
        Ok("data.textToColumns")
    }
    pub(crate) fn native_filter(&self, session: &Session, sheet: usize, key: u64) -> HostResult<AutoFilter> {
        let sh = session.doc().map_err(|e| failure(91, e.to_string()))?.wb.sheet(sheet).ok_or_else(|| failure(91, "worksheet deleted"))?;
        if key == 0 {
            sh.autofilter.clone().ok_or_else(|| failure(1004, "worksheet has no AutoFilter"))
        } else {
            let table = sh.tables.iter().find(|t| u64::from(t.id) == key).ok_or_else(|| failure(91, "table was deleted"))?;
            Ok(AutoFilter { range: table.range, criteria: table.filter.clone() })
        }
    }
    fn autofilter_range(&mut self, session: &mut Session, sheet: usize, rect: RangeRef, call: &HostCall) -> HostResult<Value> {
        let args = arguments(call, &["Field", "Criteria1", "Operator", "Criteria2", "VisibleDropDown"])?;
        if !absent(&args[4]) && !super::address::boolean(&args[4])? {
            return Err(failure(1004, "per-column dropdown visibility is unavailable"));
        }
        let sh = session.doc().map_err(|e| failure(91, e.to_string()))?.wb.sheet(sheet).ok_or_else(|| failure(91, "worksheet deleted"))?;
        let table = sh.tables.iter().find(|t| t.range == rect).map(|t| t.name.clone());
        let mut params = json!({"sheet":sheet,"range":rect.a1()});
        if absent(&args[0]) {
            if args.iter().skip(1).any(|v| !absent(v)) {
                return Err(failure(449, "Field is required with criteria"));
            }
            session.execute("data.filter", params).map_err(|e| failure(1004, e.to_string()))?;
            return Ok(Value::Empty);
        }
        let field = numeric_index(&args[0])?;
        if field < 1 || field as u32 > rect.width() {
            return Err(failure(1004, "filter Field outside range"));
        }
        params["column"] = json!(gridcraft_engine::core::col_to_letters(rect.start.col + field as u32 - 1));
        let operation = if absent(&args[2]) { 1 } else { numeric_index(&args[2])? };
        if absent(&args[1]) {
            if !absent(&args[3]) {
                return Err(failure(449, "Criteria1 is required"));
            }
            params["clear"] = json!(true);
        } else if let Value::Array(array) = &args[1] {
            if operation != 7 {
                return Err(failure(1004, "array criteria require xlFilterValues"));
            }
            if array.elements().len() > 65536 {
                return Err(failure(7, "filter criteria budget"));
            }
            params["values"] = json!(array.elements().iter().map(text).collect::<HostResult<Vec<_>>>()?);
        } else if matches!(operation, 3..=6) {
            let count = numeric_index(&args[1])?;
            if !(1..=1000).contains(&count) {
                return Err(failure(1004, "invalid top/bottom count"));
            }
            params["top"] = json!({"count":count,"bottom":matches!(operation,4|6),"percent":matches!(operation,5|6)});
        } else {
            if !matches!(operation, 1 | 2 | 7) {
                return Err(failure(1004, "this AutoFilter operator is unavailable"));
            }
            let criterion = text(&args[1])?;
            if operation == 7 {
                if !absent(&args[3]) {
                    return Err(failure(1004, "Criteria2 cannot accompany xlFilterValues"));
                }
                params["values"] = json!([criterion]);
            } else {
                let split = |s: String| {
                    let operator = ["<=", ">=", "<>", "<", ">", "="].into_iter().find(|op| s.starts_with(op)).unwrap_or("=");
                    let value = s.strip_prefix(operator).unwrap_or(&s).to_string();
                    (operator.to_string(), value)
                };
                let (a, b) = split(criterion);
                let mut custom = json!({"op":a,"value":b,"and":operation==1});
                if !absent(&args[3]) {
                    let (a, b) = split(text(&args[3])?);
                    custom["op2"] = json!(a);
                    custom["value2"] = json!(b);
                }
                params["custom"] = custom;
            }
        }
        if let Some(table) = table {
            params["table"] = json!(table);
        } else if sh.autofilter.as_ref().is_none_or(|af| af.range != rect) {
            session.execute("data.filter", json!({"sheet":sheet,"range":rect.a1(),"ensure":true})).map_err(|e| failure(1004, e.to_string()))?;
        }
        session.execute("data.filterBy", params).map_err(|e| failure(1004, e.to_string()))?;
        Ok(Value::Empty)
    }
    fn filter_dispatch(&mut self, session: &mut Session, f: &Feature, call: &HostCall) -> HostResult<Value> {
        let sheet = f.sheet.ok_or_else(|| failure(91, "filter owner deleted"))?;
        validate_table_view(session, sheet, f.key, f)?;
        let af = self.native_filter(session, sheet, f.key)?;
        let member = call.member.to_ascii_lowercase();
        if member == "application" && reads(call) {
            arguments(call, &[])?;
            return self.intern(Object::Application);
        }
        if member == "parent" && reads(call) {
            arguments(call, &[])?;
            if f.kind == "AutoFilter" {
                if f.key == 0 {
                    return self.intern(Object::Sheet(sheet));
                }
                let mut p = f.clone();
                p.kind = "ListObject";
                return self.feature(p);
            }
            let mut parent = f.clone();
            parent.kind = if f.kind == "Filters" { "AutoFilter" } else { "Filters" };
            parent.part = 0;
            return self.feature(parent);
        }
        if f.kind == "AutoFilter" && member == "filters" && reads(call) && !call.arguments.is_empty() {
            let mut f = f.clone();
            f.kind = "Filters";
            let mut call = call.clone();
            call.member = "Item".into();
            return self.filter_dispatch(session, &f, &call);
        }
        if f.kind == "AutoFilter" {
            arguments(call, &[])?;
            return match member.as_str() {
                "range" => self.intern(Object::Range(sheet, af.range.into())),
                "filters" => {
                    let mut n = f.clone();
                    n.kind = "Filters";
                    self.feature(n)
                }
                "filtermode" => Ok(Value::Boolean(!af.criteria.is_empty())),
                "showalldata" | "applyfilter" => {
                    let mut params = json!({"sheet":sheet});
                    if f.key != 0 {
                        let name = session
                            .doc()
                            .map_err(|e| failure(91, e.to_string()))?
                            .wb
                            .sheet(sheet)
                            .and_then(|sh| sh.tables.iter().find(|t| u64::from(t.id) == f.key))
                            .ok_or_else(|| failure(91, "table deleted"))?
                            .name
                            .clone();
                        params["table"] = json!(name);
                    }
                    session
                        .execute(if member == "showalldata" { "data.clearFilter" } else { "data.reapply" }, params)
                        .map_err(|e| failure(1004, e.to_string()))?;
                    Ok(Value::Empty)
                }
                _ => Err(failure(438, "unsupported AutoFilter member")),
            };
        }
        if f.kind == "Filters" {
            if call.kind == InvokeKind::Enumerate {
                arguments(call, &[])?;
                return self.object_array(
                    (0..af.range.width())
                        .map(|part| {
                            let mut f = f.clone();
                            f.kind = "Filter";
                            f.part = part as i32;
                            f.object()
                        })
                        .collect(),
                );
            }
            if member == "count" {
                arguments(call, &[])?;
                return Ok(Value::Long(af.range.width() as i32));
            }
            if matches!(member.as_str(), "item" | "") {
                let args = arguments(call, &["Index"])?;
                let index = numeric_index(required(&args, 0)?)?;
                if index < 1 || index as u32 > af.range.width() {
                    return Err(failure(9, "filter index"));
                }
                let mut n = f.clone();
                n.kind = "Filter";
                n.part = index - 1;
                return self.feature(n);
            }
            return Err(failure(438, "unsupported Filters member"));
        }
        arguments(call, &[])?;
        let criterion = af.criteria.iter().find(|(col, _)| *col == f.part as u32).map(|(_, c)| c);
        if member == "on" {
            return Ok(Value::Boolean(criterion.is_some()));
        }
        let criterion = criterion.ok_or_else(|| failure(1004, "filter is not active"))?;
        match (member.as_str(), criterion) {
            ("operator", FilterCriterion::Values { .. }) => Ok(Value::Long(7)),
            ("operator", FilterCriterion::Custom { and, .. }) => Ok(Value::Long(if *and { 1 } else { 2 })),
            ("operator", FilterCriterion::Top10 { bottom, percent, .. }) => Ok(Value::Long(match (bottom, percent) {
                (false, false) => 3,
                (true, false) => 4,
                (false, true) => 5,
                (true, true) => 6,
            })),
            ("criteria1", FilterCriterion::Custom { a, .. }) => values::vba_text(&format!("{}{}", a.0, a.1)),
            ("criteria2", FilterCriterion::Custom { b: Some(b), .. }) => values::vba_text(&format!("{}{}", b.0, b.1)),
            ("criteria1", FilterCriterion::Top10 { count, .. }) => Ok(Value::Long(*count as i32)),
            ("criteria1", FilterCriterion::Values { values: items, .. }) => {
                if items.len() == 1 {
                    return values::vba_text(items.first().map(String::as_str).unwrap_or(""));
                }
                let mut array = vba_values::VbaArray::new(vba_values::ScalarKind::Variant, vec![(0, items.len() as i32 - 1)], false)
                    .map_err(|e| failure(7, e.to_string()))?;
                for (i, item) in items.iter().enumerate() {
                    array.set(i, values::vba_text(item)?).map_err(|e| failure(7, e.to_string()))?;
                }
                Ok(Value::Array(array))
            }
            _ => Err(failure(1004, "filter criterion is not available in this form")),
        }
    }
    fn validation_api(&mut self, session: &mut Session, f: &Feature, call: &HostCall) -> HostResult<Value> {
        let sheet = f.sheet.ok_or_else(|| failure(91, "validation owner deleted"))?;
        let range = f.range.as_ref().ok_or_else(|| failure(91, "validation range deleted"))?;
        let rect = range.rectangle()?;
        values::dimensions(rect)?;
        let member = call.member.to_ascii_lowercase();
        let doc = session.doc().map_err(|e| failure(91, e.to_string()))?;
        let sh = doc.wb.sheet(sheet).ok_or_else(|| failure(91, "worksheet deleted"))?;
        if member == "parent" && reads(call) {
            arguments(call, &[])?;
            return self.intern(Object::Range(sheet, range.clone()));
        }
        if member == "application" && reads(call) {
            arguments(call, &[])?;
            return self.intern(Object::Application);
        }
        let existing = sh.validations.iter().find(|v| v.ranges.iter().any(|r| r.contains(rect.start) && r.contains(rect.end))).cloned();
        let mut params = json!({"sheet":sheet,"range":rect.a1()});
        if matches!(member.as_str(), "add" | "modify" | "delete") {
            if sh.is_protected() {
                return Err(failure(1004, "validation changes require an unprotected worksheet"));
            }
            if member == "delete" {
                arguments(call, &[])?;
                params["clear"] = json!(true);
            } else {
                let args = arguments(call, &["Type", "AlertStyle", "Operator", "Formula1", "Formula2"])?;
                let kind = if absent(&args[0]) && member == "modify" {
                    existing
                        .as_ref()
                        .map(|v| match v.kind {
                            ValidationKind::Any => 0,
                            ValidationKind::Whole => 1,
                            ValidationKind::Decimal => 2,
                            ValidationKind::List => 3,
                            ValidationKind::Date => 4,
                            ValidationKind::Time => 5,
                            ValidationKind::TextLength => 6,
                            ValidationKind::Custom => 7,
                        })
                        .ok_or_else(|| failure(1004, "no existing validation"))?
                } else {
                    numeric_index(required(&args, 0)?)?
                };
                params["type"] = json!(dv_type(kind)?);
                params["errorStyle"] = json!(match if absent(&args[1]) { 1 } else { numeric_index(&args[1])? } {
                    1 => "stop",
                    2 => "warning",
                    3 => "information",
                    _ => return Err(failure(1004, "invalid AlertStyle")),
                });
                params["operator"] = json!(op(if absent(&args[2]) { 1 } else { numeric_index(&args[2])? })?);
                if !absent(&args[3]) {
                    params["formula1"] = json!(text(&args[3])?);
                }
                if !absent(&args[4]) {
                    params["formula2"] = json!(text(&args[4])?);
                }
            }
            session.execute("data.validation", params).map_err(|e| failure(1004, e.to_string()))?;
            return Ok(Value::Empty);
        }
        let validation = existing.ok_or_else(|| failure(1004, "no single validation covers the range"))?;
        if reads(call) {
            arguments(call, &[])?;
            return match member.as_str() {
                "type" => Ok(Value::Long(match validation.kind {
                    ValidationKind::Any => 0,
                    ValidationKind::Whole => 1,
                    ValidationKind::Decimal => 2,
                    ValidationKind::List => 3,
                    ValidationKind::Date => 4,
                    ValidationKind::Time => 5,
                    ValidationKind::TextLength => 6,
                    ValidationKind::Custom => 7,
                })),
                "operator" => Ok(Value::Long(op_number(validation.op))),
                "alertstyle" => Ok(Value::Long(match validation.error_style {
                    ErrorStyle::Stop => 1,
                    ErrorStyle::Warning => 2,
                    ErrorStyle::Information => 3,
                })),
                "formula1" => values::vba_text(&validation.f1),
                "formula2" => values::vba_text(validation.f2.as_deref().unwrap_or("")),
                "ignoreblank" => Ok(Value::Boolean(validation.allow_blank)),
                "incelldropdown" => Ok(Value::Boolean(validation.in_cell_dropdown)),
                "showinput" => Ok(Value::Boolean(validation.show_input)),
                "showerror" => Ok(Value::Boolean(validation.show_error)),
                "inputtitle" => values::vba_text(&validation.input_title),
                "inputmessage" => values::vba_text(&validation.input_message),
                "errortitle" => values::vba_text(&validation.error_title),
                "errormessage" => values::vba_text(&validation.error_message),
                "value" => Ok(Value::Boolean(
                    rect.iter().all(|at| gridcraft_engine::cmd::data::check_validation(&doc.wb, sheet, at, &sh.value(at).display()).is_none()),
                )),
                _ => Err(failure(438, "unsupported Validation member")),
            };
        }
        if call.kind != InvokeKind::Let {
            return Err(failure(13, "validation assignment requires Let"));
        }
        if sh.is_protected() {
            return Err(failure(1004, "validation changes require unprotected worksheet"));
        }
        let args = arguments(call, &["Value"])?;
        let value = required(&args, 0)?;
        params = self.validation_params(sheet, rect, &validation);
        let field = match member.as_str() {
            "ignoreblank" => "ignoreBlank",
            "incelldropdown" => "dropdown",
            "showinput" => "showInput",
            "showerror" => "showError",
            "inputtitle" => "inputTitle",
            "inputmessage" => "inputMessage",
            "errortitle" => "errorTitle",
            "errormessage" => "errorMessage",
            _ => return Err(failure(438, "unsupported validation assignment")),
        };
        params[field] = if matches!(member.as_str(), "ignoreblank" | "incelldropdown" | "showinput" | "showerror") {
            json!(super::address::boolean(value)?)
        } else {
            let text = string(value)?;
            if text.len() > 4096 {
                return Err(failure(7, "validation text budget"));
            }
            json!(text)
        };
        session.execute("data.validation", params).map_err(|e| failure(1004, e.to_string()))?;
        Ok(Value::Empty)
    }
    fn validation_params(&self, sheet: usize, rect: RangeRef, v: &Validation) -> Json {
        json!({"sheet":sheet,"range":rect.a1(),"type":match v.kind{ValidationKind::Any=>"any",ValidationKind::Whole=>"whole",ValidationKind::Decimal=>"decimal",ValidationKind::List=>"list",ValidationKind::Date=>"date",ValidationKind::Time=>"time",ValidationKind::TextLength=>"textLength",ValidationKind::Custom=>"custom"},"operator":op(op_number(v.op)).unwrap_or("between"),"formula1":v.f1,"formula2":v.f2,"ignoreBlank":v.allow_blank,"dropdown":v.in_cell_dropdown,"inputTitle":v.input_title,"inputMessage":v.input_message,"showInput":v.show_input,"showError":v.show_error,"errorTitle":v.error_title,"errorMessage":v.error_message,"errorStyle":match v.error_style{ErrorStyle::Stop=>"stop",ErrorStyle::Warning=>"warning",ErrorStyle::Information=>"information"}})
    }
}

pub(crate) fn descriptors() -> Vec<super::MemberDescriptor> {
    vec![
        super::MemberDescriptor {
            owner: "Worksheet",
            name: "Sort",
            signature: "Property Get Sort() As Sort",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/data_api.rs::Bridge::data_api",
        },
        super::MemberDescriptor {
            owner: "Worksheet",
            name: "AutoFilter",
            signature: "Property Get AutoFilter() As AutoFilter",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/data_api.rs::Bridge::data_api",
        },
        super::MemberDescriptor {
            owner: "Range",
            name: "Validation",
            signature: "Property Get Validation() As Validation",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/data_api.rs::Bridge::data_api",
        },
        super::MemberDescriptor {
            owner: "Sort",
            name: "SortFields",
            signature: "Property Get SortFields() As SortFields",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/data_api.rs::Bridge::data_api",
        },
        super::MemberDescriptor {
            owner: "Sort",
            name: "SetRange",
            signature: "Sub SetRange(Rng As Range)",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/data_api.rs::Bridge::data_api",
        },
        super::MemberDescriptor {
            owner: "Sort",
            name: "Apply",
            signature: "Sub Apply()",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/data_api.rs::Bridge::data_api",
        },
        super::MemberDescriptor {
            owner: "Sort",
            name: "Header",
            signature: "Property Header As Long",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/data_api.rs::Bridge::data_api",
        },
        super::MemberDescriptor {
            owner: "Sort",
            name: "MatchCase",
            signature: "Property MatchCase As Boolean",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/data_api.rs::Bridge::data_api",
        },
        super::MemberDescriptor {
            owner: "Sort",
            name: "Orientation",
            signature: "Property Orientation As Long",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/data_api.rs::Bridge::data_api",
        },
        super::MemberDescriptor {
            owner: "Sort",
            name: "Parent",
            signature: "Property Get Parent() As Object",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/data_api.rs::Bridge::data_api",
        },
        super::MemberDescriptor {
            owner: "Sort",
            name: "Application",
            signature: "Property Get Application() As Application",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/data_api.rs::Bridge::data_api",
        },
        super::MemberDescriptor {
            owner: "SortFields",
            name: "Parent",
            signature: "Property Get Parent() As Object",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/data_api.rs::Bridge::data_api",
        },
        super::MemberDescriptor {
            owner: "SortFields",
            name: "Application",
            signature: "Property Get Application() As Application",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/data_api.rs::Bridge::data_api",
        },
        super::MemberDescriptor {
            owner: "SortField",
            name: "Parent",
            signature: "Property Get Parent() As Object",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/data_api.rs::Bridge::data_api",
        },
        super::MemberDescriptor {
            owner: "SortField",
            name: "Application",
            signature: "Property Get Application() As Application",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/data_api.rs::Bridge::data_api",
        },
        super::MemberDescriptor {
            owner: "AutoFilter",
            name: "Parent",
            signature: "Property Get Parent() As Object",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/data_api.rs::Bridge::data_api",
        },
        super::MemberDescriptor {
            owner: "AutoFilter",
            name: "Application",
            signature: "Property Get Application() As Application",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/data_api.rs::Bridge::data_api",
        },
        super::MemberDescriptor {
            owner: "Filters",
            name: "Parent",
            signature: "Property Get Parent() As Object",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/data_api.rs::Bridge::data_api",
        },
        super::MemberDescriptor {
            owner: "Filters",
            name: "Application",
            signature: "Property Get Application() As Application",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/data_api.rs::Bridge::data_api",
        },
        super::MemberDescriptor {
            owner: "Filter",
            name: "Parent",
            signature: "Property Get Parent() As Object",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/data_api.rs::Bridge::data_api",
        },
        super::MemberDescriptor {
            owner: "Filter",
            name: "Application",
            signature: "Property Get Application() As Application",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/data_api.rs::Bridge::data_api",
        },
        super::MemberDescriptor {
            owner: "Validation",
            name: "Parent",
            signature: "Property Get Parent() As Object",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/data_api.rs::Bridge::data_api",
        },
        super::MemberDescriptor {
            owner: "Validation",
            name: "Application",
            signature: "Property Get Application() As Application",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/data_api.rs::Bridge::data_api",
        },
        super::MemberDescriptor {
            owner: "SortFields",
            name: "Count",
            signature: "Property Get Count() As Long",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/data_api.rs::Bridge::data_api",
        },
        super::MemberDescriptor {
            owner: "SortFields",
            name: "Item",
            signature: "Property Get Item(Index As Long) As SortField",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/data_api.rs::Bridge::data_api",
        },
        super::MemberDescriptor {
            owner: "Filters",
            name: "Count",
            signature: "Property Get Count() As Long",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/data_api.rs::Bridge::data_api",
        },
        super::MemberDescriptor {
            owner: "Filters",
            name: "Item",
            signature: "Property Get Item(Index As Long) As Filter",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/data_api.rs::Bridge::data_api",
        },
        super::MemberDescriptor {
            owner: "SortFields",
            name: "Clear",
            signature: "Sub Clear()",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/data_api.rs::Bridge::data_api",
        },
        super::MemberDescriptor {
            owner: "SortFields",
            name: "Add",
            signature: "Function Add(Key As Range, Optional SortOn As Variant, Optional Order As Variant, Optional CustomOrder As Variant, Optional DataOption As Variant) As SortField",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/data_api.rs::Bridge::data_api",
        },
        super::MemberDescriptor {
            owner: "SortFields",
            name: "Add2",
            signature: "Function Add2(Key As Range, Optional SortOn As Variant, Optional Order As Variant, Optional CustomOrder As Variant, Optional DataOption As Variant, Optional SubField As Variant) As SortField",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/data_api.rs::Bridge::data_api",
        },
        super::MemberDescriptor {
            owner: "SortField",
            name: "Delete",
            signature: "Sub Delete()",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/data_api.rs::Bridge::data_api",
        },
        super::MemberDescriptor {
            owner: "SortField",
            name: "Key",
            signature: "Property Get Key() As Range",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/data_api.rs::Bridge::data_api",
        },
        super::MemberDescriptor {
            owner: "SortField",
            name: "CustomOrder",
            signature: "Property Get CustomOrder() As String",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/data_api.rs::Bridge::data_api",
        },
        super::MemberDescriptor {
            owner: "SortField",
            name: "Order",
            signature: "Property Order As Long",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/data_api.rs::Bridge::data_api",
        },
        super::MemberDescriptor {
            owner: "SortField",
            name: "SortOn",
            signature: "Property SortOn As Long",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/data_api.rs::Bridge::data_api",
        },
        super::MemberDescriptor {
            owner: "SortField",
            name: "DataOption",
            signature: "Property DataOption As Long",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/data_api.rs::Bridge::data_api",
        },
        super::MemberDescriptor {
            owner: "Worksheet",
            name: "AutoFilterMode",
            signature: "Property AutoFilterMode As Boolean",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/data_api.rs::Bridge::data_api",
        },
        super::MemberDescriptor {
            owner: "Worksheet",
            name: "FilterMode",
            signature: "Property Get FilterMode() As Boolean",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/data_api.rs::Bridge::data_api",
        },
        super::MemberDescriptor {
            owner: "Worksheet",
            name: "ShowAllData",
            signature: "Sub ShowAllData()",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/data_api.rs::Bridge::data_api",
        },
        super::MemberDescriptor {
            owner: "AutoFilter",
            name: "Range",
            signature: "Property Get Range() As Range",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/data_api.rs::Bridge::data_api",
        },
        super::MemberDescriptor {
            owner: "AutoFilter",
            name: "Filters",
            signature: "Property Get Filters() As Filters",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/data_api.rs::Bridge::data_api",
        },
        super::MemberDescriptor {
            owner: "AutoFilter",
            name: "FilterMode",
            signature: "Property Get FilterMode() As Boolean",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/data_api.rs::Bridge::data_api",
        },
        super::MemberDescriptor {
            owner: "AutoFilter",
            name: "ShowAllData",
            signature: "Sub ShowAllData()",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/data_api.rs::Bridge::data_api",
        },
        super::MemberDescriptor {
            owner: "AutoFilter",
            name: "ApplyFilter",
            signature: "Sub ApplyFilter()",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/data_api.rs::Bridge::data_api",
        },
        super::MemberDescriptor {
            owner: "Filter",
            name: "On",
            signature: "Property Get On() As Boolean",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/data_api.rs::Bridge::data_api",
        },
        super::MemberDescriptor {
            owner: "Filter",
            name: "Criteria1",
            signature: "Property Get Criteria1() As Variant",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/data_api.rs::Bridge::data_api",
        },
        super::MemberDescriptor {
            owner: "Filter",
            name: "Criteria2",
            signature: "Property Get Criteria2() As Variant",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/data_api.rs::Bridge::data_api",
        },
        super::MemberDescriptor {
            owner: "Filter",
            name: "Operator",
            signature: "Property Get Operator() As Long",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/data_api.rs::Bridge::data_api",
        },
        super::MemberDescriptor {
            owner: "Range",
            name: "AutoFilter",
            signature: "Sub AutoFilter(Optional Field As Variant, Optional Criteria1 As Variant, Optional Operator As Variant, Optional Criteria2 As Variant, Optional VisibleDropDown As Variant)",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/data_api.rs::Bridge::data_api",
        },
        super::MemberDescriptor {
            owner: "Range",
            name: "Sort",
            signature: "Sub Sort(Key1 As Variant, Optional Order1 As Variant, Optional Key2 As Variant, Optional Type As Variant, Optional Order2 As Variant, Optional Key3 As Variant, Optional Order3 As Variant, Optional Header As Variant, Optional OrderCustom As Variant, Optional MatchCase As Variant, Optional Orientation As Variant, Optional SortMethod As Variant, Optional DataOption1 As Variant, Optional DataOption2 As Variant, Optional DataOption3 As Variant)",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/data_api.rs::Bridge::data_api",
        },
        super::MemberDescriptor {
            owner: "Range",
            name: "RemoveDuplicates",
            signature: "Sub RemoveDuplicates(Columns As Variant, Optional Header As Variant)",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/data_api.rs::Bridge::data_api",
        },
        super::MemberDescriptor {
            owner: "Range",
            name: "TextToColumns",
            signature: "Sub TextToColumns(Optional Destination As Variant, Optional DataType As Variant, Optional TextQualifier As Variant, Optional ConsecutiveDelimiter As Variant, Optional Tab As Variant, Optional Semicolon As Variant, Optional Comma As Variant, Optional Space As Variant, Optional Other As Variant, Optional OtherChar As Variant, Optional FieldInfo As Variant, Optional DecimalSeparator As Variant, Optional ThousandsSeparator As Variant, Optional TrailingMinusNumbers As Variant)",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/data_api.rs::Bridge::data_api",
        },
        super::MemberDescriptor {
            owner: "Range",
            name: "Group",
            signature: "Sub Group(Optional Start As Variant, Optional End As Variant, Optional By As Variant, Optional Periods As Variant)",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/data_api.rs::Bridge::data_api",
        },
        super::MemberDescriptor {
            owner: "Range",
            name: "Ungroup",
            signature: "Sub Ungroup(Optional Start As Variant, Optional End As Variant, Optional By As Variant, Optional Periods As Variant)",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/data_api.rs::Bridge::data_api",
        },
        super::MemberDescriptor {
            owner: "Range",
            name: "ShowDetail",
            signature: "Property Let ShowDetail(Value As Boolean)",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/data_api.rs::Bridge::data_api",
        },
        super::MemberDescriptor {
            owner: "Range",
            name: "Subtotal",
            signature: "Sub Subtotal(GroupBy As Long, Function As Long, TotalList As Variant, Optional Replace As Variant, Optional PageBreaks As Variant, Optional SummaryBelowData As Variant)",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/data_api.rs::Bridge::data_api",
        },
        super::MemberDescriptor {
            owner: "Validation",
            name: "Add",
            signature: "Sub Add(Type As Variant, Optional AlertStyle As Variant, Optional Operator As Variant, Optional Formula1 As Variant, Optional Formula2 As Variant)",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/data_api.rs::Bridge::data_api",
        },
        super::MemberDescriptor {
            owner: "Validation",
            name: "Modify",
            signature: "Sub Modify(Type As Variant, Optional AlertStyle As Variant, Optional Operator As Variant, Optional Formula1 As Variant, Optional Formula2 As Variant)",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/data_api.rs::Bridge::data_api",
        },
        super::MemberDescriptor {
            owner: "Validation",
            name: "Delete",
            signature: "Sub Delete()",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/data_api.rs::Bridge::data_api",
        },
        super::MemberDescriptor {
            owner: "Validation",
            name: "Type",
            signature: "Property Get Type() As Long",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/data_api.rs::Bridge::data_api",
        },
        super::MemberDescriptor {
            owner: "Validation",
            name: "Operator",
            signature: "Property Get Operator() As Long",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/data_api.rs::Bridge::data_api",
        },
        super::MemberDescriptor {
            owner: "Validation",
            name: "AlertStyle",
            signature: "Property Get AlertStyle() As Long",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/data_api.rs::Bridge::data_api",
        },
        super::MemberDescriptor {
            owner: "Validation",
            name: "Formula1",
            signature: "Property Get Formula1() As String",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/data_api.rs::Bridge::data_api",
        },
        super::MemberDescriptor {
            owner: "Validation",
            name: "Formula2",
            signature: "Property Get Formula2() As String",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/data_api.rs::Bridge::data_api",
        },
        super::MemberDescriptor {
            owner: "Validation",
            name: "Value",
            signature: "Property Get Value() As Boolean",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/data_api.rs::Bridge::data_api",
        },
        super::MemberDescriptor {
            owner: "Validation",
            name: "IgnoreBlank",
            signature: "Property IgnoreBlank As Boolean",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/data_api.rs::Bridge::data_api",
        },
        super::MemberDescriptor {
            owner: "Validation",
            name: "InCellDropdown",
            signature: "Property InCellDropdown As Boolean",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/data_api.rs::Bridge::data_api",
        },
        super::MemberDescriptor {
            owner: "Validation",
            name: "ShowInput",
            signature: "Property ShowInput As Boolean",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/data_api.rs::Bridge::data_api",
        },
        super::MemberDescriptor {
            owner: "Validation",
            name: "ShowError",
            signature: "Property ShowError As Boolean",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/data_api.rs::Bridge::data_api",
        },
        super::MemberDescriptor {
            owner: "Validation",
            name: "InputTitle",
            signature: "Property InputTitle As String",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/data_api.rs::Bridge::data_api",
        },
        super::MemberDescriptor {
            owner: "Validation",
            name: "InputMessage",
            signature: "Property InputMessage As String",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/data_api.rs::Bridge::data_api",
        },
        super::MemberDescriptor {
            owner: "Validation",
            name: "ErrorTitle",
            signature: "Property ErrorTitle As String",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/data_api.rs::Bridge::data_api",
        },
        super::MemberDescriptor {
            owner: "Validation",
            name: "ErrorMessage",
            signature: "Property ErrorMessage As String",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/data_api.rs::Bridge::data_api",
        },
        super::MemberDescriptor {
            owner: "Sort",
            name: "SortMethod",
            signature: "Property SortMethod As Long",
            availability: super::MemberAvailability::OutOfScope,
            implementation_symbol: "crates/vba/src/data_api.rs::Bridge::data_api",
        },
    ]
}
