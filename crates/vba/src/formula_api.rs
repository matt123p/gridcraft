//! OM07: typed formula/function projection; mathematics remains in the calculator.
use super::features::{Feature, absent, reads};
use super::objects::{arguments, numeric_index, required, string};
use super::{Bridge, HostCall, HostResult, InvokeKind, MAX_CELLS, Object, Session, Value, failure, ranges, values};
use gridcraft_engine::{
    CellContent,
    core::{CellRef, Value as CellValue},
    formula::{self, Anchor, Expr, RefKind, Reference, SheetSel},
};
use serde_json::json;
use vba_values::{ScalarKind, VbaArray};

pub(crate) fn function_name(member: &str) -> Option<String> {
    let name = member.replace('_', ".").to_ascii_uppercase();
    gridcraft_engine::calc::is_known_function(&name).then_some(name)
}
fn scalar_expr(value: &Value) -> HostResult<Expr> {
    if matches!(value, Value::Missing) {
        return Ok(Expr::Missing);
    }
    match values::content(value, false)? {
        CellContent::Value(CellValue::Empty) => Ok(Expr::Missing),
        CellContent::Value(CellValue::Number(n)) => Ok(Expr::Number(n)),
        CellContent::Value(CellValue::Text(s)) => Ok(Expr::Text(s)),
        CellContent::Value(CellValue::Bool(b)) => Ok(Expr::Bool(b)),
        CellContent::Value(CellValue::Error(e)) => Ok(Expr::Error(e)),
        _ => Err(failure(13, "scalar formula argument required")),
    }
}
pub(crate) fn result_value(value: CellValue) -> HostResult<Value> {
    let CellValue::Array(array) = value else {
        return values::from_cell(value);
    };
    if array.rows.checked_mul(array.cols).is_none_or(|n| n > MAX_CELLS) {
        return Err(failure(7, "formula result cell budget"));
    }
    let mut result =
        VbaArray::new(ScalarKind::Variant, vec![(1, array.rows as i32), (1, array.cols as i32)], false).map_err(|e| failure(7, e.to_string()))?;
    let mut payload = 0usize;
    for (i, cell) in array.data.iter().cloned().enumerate() {
        let cell = values::from_cell(cell)?;
        payload = payload.saturating_add(cell.payload_bytes() + 64);
        if payload > 8 * 1024 * 1024 {
            return Err(failure(7, "formula result payload budget"));
        }
        result.set(values::vba_matrix_index(i, array.rows, array.cols), cell).map_err(|e| failure(7, e.to_string()))?;
    }
    Ok(Value::Array(result))
}
impl Bridge {
    fn argument_expr(&self, session: &Session, value: &Value) -> HostResult<Expr> {
        if let Value::Object(reference) = value {
            let Object::Range(sheet, range) = self.object(session, *reference)? else {
                return Err(failure(13, "Range argument required"));
            };
            if range.count()? > MAX_CELLS as u64 {
                return Err(failure(7, "function reference budget"));
            }
            let name =
                session.doc().map_err(|e| failure(91, e.to_string()))?.wb.sheet(sheet).ok_or_else(|| failure(91, "worksheet deleted"))?.name.clone();
            let exprs = range
                .areas
                .iter()
                .map(|area| {
                    let anchor = |at: CellRef| Anchor { row: at.row, col: at.col, row_abs: true, col_abs: true };
                    Expr::Ref(Reference { sheet: SheetSel::Named(name.clone()), kind: RefKind::Range(anchor(area.start), anchor(area.end)) })
                })
                .collect::<Vec<_>>();
            return exprs
                .into_iter()
                .reduce(|a, b| Expr::Binary(formula::BinOp::Union, Box::new(a), Box::new(b)))
                .ok_or_else(|| failure(1004, "empty range"));
        }
        if let Value::Array(array) = value {
            value.validate().map_err(|e| failure(13, e.to_string()))?;
            if array.elements().len() > MAX_CELLS {
                return Err(failure(7, "function array budget"));
            }
            let shape = array.bounds().iter().map(|(a, b)| usize::try_from(i64::from(*b) - i64::from(*a) + 1).unwrap_or(0)).collect::<Vec<_>>();
            let cols = match shape.as_slice() {
                [n] => *n,
                [_, cols] => *cols,
                _ => return Err(failure(13, "one or two dimensional function array required")),
            };
            if cols == 0 {
                return Err(failure(13, "empty function array"));
            }
            let rows = if shape.len() == 1 { 1 } else { shape[0] };
            let data = (0..rows)
                .map(|r| {
                    (0..cols)
                        .map(|c| {
                            let index = if shape.len() == 1 { c } else { r + c * rows };
                            scalar_expr(array.elements().get(index).ok_or_else(|| failure(13, "function array shape mismatch"))?)
                        })
                        .collect::<HostResult<Vec<_>>>()
                })
                .collect::<HostResult<Vec<_>>>()?;
            return Ok(Expr::Array(data));
        }
        scalar_expr(value)
    }
    pub(crate) fn formula_api(&mut self, session: &mut Session, object: &Object, call: &HostCall) -> Option<HostResult<Value>> {
        let member = call.member.to_ascii_lowercase();
        if matches!(object, Object::Application) && member == "worksheetfunction" && reads(call) {
            return Some((|| {
                arguments(call, &[])?;
                self.feature(Feature::new("WorksheetFunction", None))
            })());
        }
        if matches!(object, Object::Application | Object::Workbook | Object::Sheet(_)) && member == "names" && reads(call) {
            return Some({
                let f = Feature::new("Names", if let Object::Sheet(s) = object { Some(*s) } else { None });
                if call.arguments.is_empty() { self.feature(f) } else { self.names_api(session, &f, call, "item") }
            });
        }
        if let Object::Feature(feature) = object {
            if let Some(sheet) = feature.sheet
                && session.doc().ok().and_then(|d| d.wb.sheet(sheet)).is_none()
            {
                return Some(Err(failure(91, "worksheet deleted")));
            }
            if feature.kind == "WorksheetFunction" {
                if reads(call) && matches!(member.as_str(), "application" | "parent") {
                    return Some(arguments(call, &[]).and_then(|_| self.intern(Object::Application)));
                }
                return Some(self.worksheet_function(session, call));
            }
            if feature.kind == "Names" {
                return Some(self.names_api(session, feature, call, &member));
            }
            if feature.kind == "Name" {
                return Some(self.name_api(session, feature, call, &member));
            }
        }
        if member == "evaluate" && reads(call) && matches!(object, Object::Application | Object::Sheet(_) | Object::Range(..)) {
            return Some((|| {
                let args = arguments(call, &["Name"])?;
                let text = string(required(&args, 0)?)?;
                if text.len() > 8192 {
                    return Err(failure(7, "Evaluate input budget"));
                }
                let doc = session.doc().map_err(|e| failure(91, e.to_string()))?;
                let (sheet, at) = match object {
                    Object::Sheet(s) => (*s, doc.selection.active),
                    Object::Range(s, r) => (*s, r.first()?.start),
                    _ => (doc.sheet_index(), doc.selection.active),
                };
                let expr = formula::parse(text.trim_start_matches('=')).map_err(|e| failure(1004, e.to_string()))?;
                match gridcraft_engine::calc::recalc::evaluate_result(&doc.wb, sheet, at, &expr) {
                    gridcraft_engine::calc::eval::Ev::V(v) => result_value(v),
                    gridcraft_engine::calc::eval::Ev::R(areas) => {
                        let owner = areas.first().map(|a| a.sheet).ok_or_else(|| failure(1004, "empty reference"))?;
                        if areas.len() > 64 || areas.iter().any(|a| a.sheet != owner) {
                            return Err(failure(1004, "Evaluate reference must belong to one worksheet"));
                        }
                        self.intern(Object::Range(
                            owner,
                            ranges::Range { areas: areas.into_iter().map(|a| a.range).collect(), unit: ranges::Unit::Cells },
                        ))
                    }
                    _ => Err(failure(1004, "Evaluate cannot return a lambda object")),
                }
            })());
        }
        if let Object::Range(sheet, range) = object
            && matches!(
                member.as_str(),
                "formula2"
                    | "formulaarray"
                    | "formular1c1"
                    | "formula2r1c1"
                    | "formulalocal"
                    | "formula2local"
                    | "formular1c1local"
                    | "formula2r1c1local"
                    | "hasformula"
            )
        {
            return Some(self.formula_property(session, *sheet, range, call, &member));
        }
        None
    }
    fn worksheet_function(&mut self, session: &Session, call: &HostCall) -> HostResult<Value> {
        if !reads(call) {
            return Err(failure(438, "WorksheetFunction is read-only"));
        }
        let name = function_name(&call.member).ok_or_else(|| failure(438, "function is not registered in GridCraft"))?;
        if call.arguments.len() > 30 {
            return Err(failure(450, "WorksheetFunction accepts at most 30 arguments"));
        }
        let names = (1..=30).map(|i| format!("Arg{i}")).collect::<Vec<_>>();
        let mut args = arguments(call, &names.iter().map(String::as_str).collect::<Vec<_>>())?;
        while args.last().is_some_and(absent) {
            args.pop();
        }
        if gridcraft_engine::functions::lookup(&name).is_some_and(|spec| !spec.accepts(args.len())) {
            return Err(failure(450, "wrong function argument count"));
        }
        let args = args.iter().map(|v| self.argument_expr(session, v)).collect::<HostResult<Vec<_>>>()?;
        let doc = session.doc().map_err(|e| failure(91, e.to_string()))?;
        let result = gridcraft_engine::calc::recalc::evaluate_expr(&doc.wb, doc.sheet_index(), doc.selection.active, &Expr::Call(name, args));
        if result.is_error() {
            return Err(failure(1004, format!("WorksheetFunction.{} returned {}", call.member, result.display())));
        }
        result_value(result)
    }
    fn names_api(&mut self, session: &mut Session, f: &Feature, call: &HostCall, member: &str) -> HostResult<Value> {
        let doc = session.doc().map_err(|e| failure(91, e.to_string()))?;
        let names = doc.wb.names.iter().filter(|n| f.sheet.is_none_or(|s| n.scope == Some(s))).collect::<Vec<_>>();
        let name_feature = |n: &gridcraft_engine::model::DefinedName| {
            let mut node = Feature::new("Name", n.scope);
            node.identity = Some(n.identity.clone());
            node
        };
        if call.kind == InvokeKind::Enumerate {
            arguments(call, &[])?;
            return self.object_array(names.iter().map(|n| name_feature(n).object()).collect());
        }
        match member {
            "count" if reads(call) => {
                arguments(call, &[])?;
                Ok(Value::Long(i32::try_from(names.len()).map_err(|_| failure(6, "Count overflow"))?))
            }
            "item" | "" if reads(call) => {
                let args = arguments(call, &["Index", "IndexLocal", "RefersTo"])?;
                if args.iter().filter(|v| !absent(v)).count() != 1 {
                    return Err(failure(1004, "specify exactly one Names.Item index"));
                }
                if !absent(&args[2]) {
                    let text = string(&args[2])?;
                    let wanted = formula::parse(text.trim_start_matches('=')).map_err(|e| failure(1004, e.to_string()))?;
                    let name = names
                        .iter()
                        .find(|n| {
                            formula::parse_name(&n.formula).is_ok_and(|expr| {
                                formula::print(&formula::adjust::shift_relative_wrapping(
                                    expr,
                                    i64::from(doc.selection.active.row),
                                    i64::from(doc.selection.active.col),
                                )) == formula::print(&wanted)
                            })
                        })
                        .ok_or_else(|| failure(1004, "no name refers to that expression"))?;
                    return self.feature(name_feature(name));
                }
                let index = if absent(&args[0]) { required(&args, 1)? } else { required(&args, 0)? };
                let name = if matches!(index, Value::String(_)) {
                    let text = string(index)?;
                    names
                        .iter()
                        .find(|n| {
                            n.name.eq_ignore_ascii_case(&text)
                                || n.scope
                                    .and_then(|s| doc.wb.sheet(s))
                                    .is_some_and(|sh| format!("{}!{}", sh.name, n.name).eq_ignore_ascii_case(&text))
                        })
                        .copied()
                } else {
                    usize::try_from(numeric_index(index)?.checked_sub(1).ok_or_else(|| failure(9, "name index"))?)
                        .ok()
                        .and_then(|i| names.get(i).copied())
                }
                .ok_or_else(|| failure(1004, "name not found"))?;
                self.feature(name_feature(name))
            }
            "add" if reads(call) => {
                let args = arguments(
                    call,
                    &[
                        "Name",
                        "RefersTo",
                        "Visible",
                        "MacroType",
                        "ShortcutKey",
                        "Category",
                        "NameLocal",
                        "RefersToLocal",
                        "CategoryLocal",
                        "RefersToR1C1",
                        "RefersToR1C1Local",
                    ],
                )?;
                if !absent(&args[3]) && numeric_index(&args[3])? != 3 || [4, 5, 8].iter().any(|i| !absent(&args[*i])) {
                    return Err(failure(1004, "macro name registration is unavailable"));
                }
                if !absent(&args[0]) && !absent(&args[6]) {
                    return Err(failure(1004, "Name and NameLocal are mutually exclusive"));
                }
                let name = string(if absent(&args[0]) { required(&args, 6)? } else { &args[0] })?;
                let (qualified, body) = gridcraft_engine::core::split_sheet(&name);
                let target = match qualified {
                    Some(name) => {
                        let sheet = doc.wb.sheet_index(&name).ok_or_else(|| failure(1004, "name scope worksheet not found"))?;
                        if f.sheet.is_some_and(|owner| owner != sheet) {
                            return Err(failure(1004, "name scope differs from its Names collection"));
                        }
                        Some(sheet)
                    }
                    None => f.sheet,
                };
                let name = body.to_string();
                let scope = target.and_then(|s| doc.wb.sheet(s)).map(|sh| sh.name.clone()).unwrap_or_else(|| "Workbook".into());
                let definitions = [1, 7, 9, 10].into_iter().filter(|i| !absent(&args[*i])).collect::<Vec<_>>();
                if definitions.len() > 1 {
                    return Err(failure(1004, "RefersTo forms are mutually exclusive"));
                }
                let mut params = json!({"name":name,"scope":scope,"hidden":!absent(&args[2])&&!super::address::boolean(&args[2])?});
                if let Some(index) = definitions.first() {
                    let raw = string(&args[*index])?;
                    let expr = self.parse_formula(
                        session,
                        target.unwrap_or(doc.sheet_index()),
                        &raw,
                        matches!(*index, 9 | 10),
                        matches!(*index, 7 | 10),
                        doc.selection.active,
                    )?;
                    params["refersTo"] = json!(format!("={}", formula::print(&expr)));
                }
                self.objects.reserve(1)?;
                session.execute("formulas.defineName", params).map_err(|e| failure(1004, e.to_string()))?;
                let name = session
                    .doc()
                    .map_err(|e| failure(91, e.to_string()))?
                    .wb
                    .names
                    .iter()
                    .find(|n| n.name.eq_ignore_ascii_case(&name) && n.scope == target)
                    .ok_or_else(|| failure(1004, "name was not created"))?;
                self.feature(name_feature(name))
            }
            "application" if reads(call) => {
                arguments(call, &[])?;
                self.intern(Object::Application)
            }
            "parent" if reads(call) => {
                arguments(call, &[])?;
                self.intern(f.sheet.map(Object::Sheet).unwrap_or(Object::Workbook))
            }
            _ => Err(failure(438, "unsupported Names member")),
        }
    }
    fn parse_formula(&self, session: &Session, sheet: usize, text: &str, r1c1: bool, local: bool, at: CellRef) -> HostResult<Expr> {
        if text.len() > 8192 {
            return Err(failure(7, "formula input budget"));
        }
        let body = text.strip_prefix('=').unwrap_or(text);
        let body = if r1c1 { r1c1_formula(body, at)? } else { body.to_string() };
        if local {
            formula::parse_input(&body, self.formula_locale, |name| session.doc().is_ok_and(|doc| doc.wb.name(name, sheet).is_some()))
        } else {
            formula::parse_name(&body)
        }
        .map_err(|e| failure(1004, e.to_string()))
    }
    fn name_api(&mut self, session: &mut Session, f: &Feature, call: &HostCall, member: &str) -> HostResult<Value> {
        let doc = session.doc().map_err(|e| failure(91, e.to_string()))?;
        let n = doc
            .wb
            .names
            .iter()
            .find(|n| f.identity.as_ref().is_some_and(|id| id.same(&n.identity)))
            .ok_or_else(|| failure(91, "name was deleted"))?
            .clone();
        let sheet = n.scope.unwrap_or(doc.sheet_index());
        let scope = n.scope.and_then(|s| doc.wb.sheet(s)).map(|s| s.name.clone()).unwrap_or_else(|| "Workbook".into());
        let at = doc.selection.active;
        let stored = formula::parse_name(&n.formula).map_err(|e| failure(1004, e.to_string()))?;
        let expanded = formula::adjust::shift_relative_wrapping(stored, i64::from(at.row), i64::from(at.col));
        if reads(call) {
            arguments(call, &[])?;
            return match member {
                "name" | "namelocal" => {
                    values::vba_text(&n.scope.and_then(|s| doc.wb.sheet(s)).map(|sh| format!("{}!{}", sh.name, n.name)).unwrap_or(n.name))
                }
                "refersto" | "referstolocal" | "referstor1c1" | "referstor1c1local" | "value" | "" => {
                    let text = if member.contains("r1c1") {
                        let text = formula::print_r1c1(&expanded, at);
                        if member.ends_with("local") { localize_r1c1(&text, self.formula_locale)? } else { text }
                    } else if member.ends_with("local") {
                        formula::print_input(&expanded, self.formula_locale, |name| doc.wb.name(name, sheet).is_some())
                            .map_err(|e| failure(1004, e.to_string()))?
                    } else {
                        formula::print(&expanded)
                    };
                    values::vba_text(&format!("={text}"))
                }
                "visible" => Ok(Value::Boolean(!n.hidden)),
                "comment" => values::vba_text(&n.comment),
                "referstorange" => {
                    let gridcraft_engine::calc::eval::Ev::R(areas) =
                        gridcraft_engine::calc::recalc::evaluate_result(&doc.wb, sheet, at, &Expr::Name(n.name))
                    else {
                        return Err(failure(1004, "name does not refer to a range"));
                    };
                    let owner = areas.first().map(|a| a.sheet).ok_or_else(|| failure(1004, "name has no reference"))?;
                    if areas.len() > 64 || areas.iter().any(|a| a.sheet != owner) {
                        return Err(failure(1004, "name reference spans worksheets"));
                    }
                    self.intern(Object::Range(
                        owner,
                        ranges::Range { areas: areas.into_iter().map(|a| a.range).collect(), unit: ranges::Unit::Cells },
                    ))
                }
                "delete" => {
                    session.execute("formulas.deleteName", json!({"name":n.name,"scope":scope})).map_err(|e| failure(1004, e.to_string()))?;
                    Ok(Value::Empty)
                }
                "parent" => self.intern(n.scope.map(Object::Sheet).unwrap_or(Object::Workbook)),
                "application" => self.intern(Object::Application),
                _ => Err(failure(438, "unsupported Name member")),
            };
        }
        if call.kind != InvokeKind::Let {
            return Err(failure(13, "Name assignment requires Let"));
        }
        let args = arguments(call, &["Value"])?;
        let value = required(&args, 0)?;
        let mut params = json!({"name":n.name,"refersTo":format!("={}",formula::print(&expanded)),"scope":scope,"oldName":n.name,"oldScope":scope,"comment":n.comment,"hidden":n.hidden});
        match member {
            "name" | "namelocal" => params["name"] = json!(string(value)?),
            "refersto" | "referstolocal" | "referstor1c1" | "referstor1c1local" | "value" => {
                let expr = self.parse_formula(session, sheet, &string(value)?, member.contains("r1c1"), member.ends_with("local"), at)?;
                params["refersTo"] = json!(format!("={}", formula::print(&expr)));
            }
            "visible" => params["hidden"] = json!(!super::address::boolean(value)?),
            "comment" => params["comment"] = json!(string(value)?),
            _ => return Err(failure(438, "unsupported Name assignment")),
        };
        session.execute("formulas.defineName", params).map_err(|e| failure(1004, e.to_string()))?;
        Ok(Value::Empty)
    }
    fn formula_property(&mut self, session: &mut Session, sheet: usize, range: &ranges::Range, call: &HostCall, member: &str) -> HostResult<Value> {
        let rect = range.rectangle()?;
        let (rows, cols) = values::dimensions(rect)?;
        let local = member.ends_with("local");
        if member == "formulaarray" {
            if reads(call) {
                arguments(call, &[])?;
                let sh = session.doc().map_err(|e| failure(91, e.to_string()))?.wb.sheet(sheet).ok_or_else(|| failure(91, "worksheet deleted"))?;
                let mut text = None;
                for at in rect.iter() {
                    let formula = sh
                        .cell(at)
                        .and_then(|c| c.formula.as_ref())
                        .filter(|f| f.array == Some(rect))
                        .or_else(|| sh.cell(rect.start).and_then(|c| c.formula.as_ref()).filter(|f| f.array == Some(rect)));
                    let Some(formula) = formula else {
                        return Ok(Value::Null);
                    };
                    let value = format!("={}", formula.text.trim_start_matches('='));
                    if text.as_ref().is_some_and(|s| s != &value) {
                        return Ok(Value::Null);
                    }
                    text = Some(value);
                }
                return values::vba_text(text.as_deref().unwrap_or(""));
            }
            if call.kind != InvokeKind::Let {
                return Err(failure(13, "FormulaArray requires Let"));
            }
            let args = arguments(call, &["Value"])?;
            let text = string(required(&args, 0)?)?;
            if !text.starts_with('=') {
                return Err(failure(1004, "FormulaArray requires formula text"));
            }
            formula::parse(text.trim_start_matches('=')).map_err(|e| failure(1004, e.to_string()))?;
            let changes = self.prepare_change(session, sheet, range.clone())?;
            session
                .execute("cell.set", json!({"sheet":sheet,"cell":rect.start.a1(),"input":text,"array":true,"arrayRange":rect.a1()}))
                .map_err(|e| failure(1004, e.to_string()))?;
            self.queue_changes(changes);
            return Ok(Value::Empty);
        }

        if member == "hasformula" {
            arguments(call, &[])?;
            let sh = session.doc().map_err(|e| failure(91, e.to_string()))?.wb.sheet(sheet).ok_or_else(|| failure(91, "worksheet deleted"))?;
            let flags = rect.iter().map(|at| sh.cell(at).is_some_and(|c| c.formula.is_some())).collect::<Vec<_>>();
            return Ok(if flags.iter().all(|b| *b) {
                Value::Boolean(true)
            } else if flags.iter().all(|b| !*b) {
                Value::Boolean(false)
            } else {
                Value::Null
            });
        }
        let r1c1 = member.contains("r1c1");
        if reads(call) {
            arguments(call, &[])?;
            if !r1c1 && !local {
                return values::read(session, sheet, rect, true, false);
            }
            let sh = session.doc().map_err(|e| failure(91, e.to_string()))?.wb.sheet(sheet).ok_or_else(|| failure(91, "worksheet deleted"))?;
            let data = rect
                .iter()
                .map(|at| match sh.cell(at).and_then(|c| c.formula.as_ref()).and_then(|f| f.expr()) {
                    Some(expr) => {
                        let text = if r1c1 {
                            {
                                let text = formula::print_r1c1(&expr, at);
                                if local { localize_r1c1(&text, self.formula_locale)? } else { text }
                            }
                        } else {
                            formula::print_input(&expr, self.formula_locale, |name| session.doc().is_ok_and(|d| d.wb.name(name, sheet).is_some()))
                                .map_err(|e| failure(1004, e.to_string()))?
                        };
                        values::vba_text(&format!("={text}"))
                    }
                    None => values::from_cell(sh.value(at)),
                })
                .collect::<HostResult<Vec<_>>>()?;
            return if rect.is_single() {
                data.into_iter().next().ok_or_else(|| failure(1004, "empty formula"))
            } else {
                let mut a =
                    VbaArray::new(ScalarKind::Variant, vec![(1, rows as i32), (1, cols as i32)], false).map_err(|e| failure(7, e.to_string()))?;
                for (i, v) in data.into_iter().enumerate() {
                    a.set(values::vba_matrix_index(i, rows, cols), v).map_err(|e| failure(7, e.to_string()))?;
                }
                Ok(Value::Array(a))
            };
        }
        if call.kind != InvokeKind::Let {
            return Err(failure(13, "formula assignment requires Let"));
        }
        let args = arguments(call, &["Value"])?;
        let value = required(&args, 0)?;
        if let Value::Array(array) = value {
            let expected = [(rows as i64), (cols as i64)];
            let shape = array.bounds().iter().map(|(a, b)| i64::from(*b) - i64::from(*a) + 1).collect::<Vec<_>>();
            if shape.as_slice() != expected {
                return Err(failure(13, "formula array dimensions do not match the destination"));
            }
        }
        if !r1c1 && !local {
            return self.write_range(session, sheet, range, value, true, false);
        }
        let mut array = VbaArray::new(ScalarKind::Variant, vec![(1, rows as i32), (1, cols as i32)], false).map_err(|e| failure(7, e.to_string()))?;
        for (i, at) in rect.iter().enumerate() {
            let input = if let Value::Array(a) = value {
                a.elements().get(values::vba_matrix_index(i, rows, cols)).ok_or_else(|| failure(13, "formula array shape mismatch"))?
            } else {
                value
            };
            let input = if let Value::String(_) = input {
                {
                    let text = string(input)?;
                    if text.starts_with('=') {
                        let text = if r1c1 { r1c1_formula(text.trim_start_matches('='), at)? } else { text.trim_start_matches('=').into() };
                        let expr = if local {
                            formula::parse_input(&text, self.formula_locale, |name| session.doc().is_ok_and(|d| d.wb.name(name, sheet).is_some()))
                        } else {
                            formula::parse(&text)
                        }
                        .map_err(|e| failure(1004, e.to_string()))?;
                        values::vba_text(&format!("={}", formula::print(&expr)))?
                    } else {
                        input.clone()
                    }
                }
            } else {
                input.clone()
            };
            array.set(values::vba_matrix_index(i, rows, cols), input).map_err(|e| failure(7, e.to_string()))?;
        }
        self.write_range(session, sheet, range, &Value::Array(array), true, false)
    }
}
/// R1C1 token conversion delegates coordinate arithmetic to the existing INDIRECT converter.
fn r1c1_formula(text: &str, at: CellRef) -> HostResult<String> {
    let mut out = String::new();
    let mut chars = text.char_indices().peekable();
    while let Some((start, c)) = chars.next() {
        if c == '"' || c == '\'' {
            out.push(c);
            while let Some((_, ch)) = chars.next() {
                out.push(ch);
                if ch == c {
                    if chars.peek().is_some_and(|(_, next)| *next == c) {
                        chars.next();
                        out.push(c);
                    } else {
                        break;
                    }
                }
            }
            continue;
        }
        if c.is_ascii_alphabetic() || c == '_' {
            let mut end = start + c.len_utf8();
            let mut brackets = 0usize;
            while let Some((i, ch)) = chars.peek().copied() {
                if ch == '[' {
                    brackets += 1;
                } else if ch == ']' {
                    brackets = brackets.saturating_sub(1);
                } else if !(ch.is_ascii_alphanumeric() || ch == '_' || ch == '.' || brackets > 0 && matches!(ch, '+' | '-')) {
                    break;
                }
                chars.next();
                end = i + ch.len_utf8();
            }
            let token = text.get(start..end).ok_or_else(|| failure(1004, "invalid formula token"))?;
            let converted = gridcraft_engine::calc::eval::r1c1_to_a1(token, at).map_err(|e| failure(1004, e.as_str()))?;
            out.push_str(converted.as_deref().unwrap_or(token));
        } else {
            out.push(c);
        }
    }
    Ok(out)
}

fn localize_r1c1(text: &str, locale: formula::FormulaLocale) -> HostResult<String> {
    if locale == formula::FormulaLocale::En {
        return Ok(text.into());
    }
    let tokens = formula::lexer::tokenize(text).map_err(|e| failure(1004, e.to_string()))?;
    let mut out = String::new();
    let mut end = 0usize;
    let mut braces = 0usize;
    for token in tokens {
        out.push_str(text.get(end..token.start).ok_or_else(|| failure(1004, "invalid formula span"))?);
        let raw = text.get(token.start..token.end).ok_or_else(|| failure(1004, "invalid formula span"))?;
        match token.tok {
            formula::lexer::Tok::Func(name) => {
                out.push_str(locale.function_name(&name));
                out.push('(');
            }
            formula::lexer::Tok::Comma => out.push(if braces > 0 { locale.array_column_separator() } else { locale.list_separator() }),
            formula::lexer::Tok::Number(_) => out.push_str(&raw.replace('.', &locale.decimal_separator().to_string())),
            formula::lexer::Tok::LBrace => {
                braces += 1;
                out.push_str(raw);
            }
            formula::lexer::Tok::RBrace => {
                braces = braces.saturating_sub(1);
                out.push_str(raw);
            }
            _ => out.push_str(raw),
        }
        end = token.end;
    }
    out.push_str(text.get(end..).ok_or_else(|| failure(1004, "invalid formula span"))?);
    Ok(out)
}

// One generic proxy supplies these documented entry names; registry lookup gates eligibility.
const FUNCTION_MEMBERS: &[(&str, &str)] = &[
    (
        "AccrInt",
        "Function AccrInt(Arg1 As Variant, Arg2 As Variant, Arg3 As Variant, Arg4 As Variant, Arg5 As Variant, Arg6 As Variant, Optional Arg7 As Variant) As Variant",
    ),
    ("AccrIntM", "Function AccrIntM(Arg1 As Variant, Arg2 As Variant, Arg3 As Variant, Arg4 As Variant, Optional Arg5 As Variant) As Variant"),
    ("Acos", "Function Acos(Arg1 As Variant) As Variant"),
    ("Acosh", "Function Acosh(Arg1 As Variant) As Variant"),
    (
        "Aggregate",
        "Function Aggregate(Arg1 As Variant, Arg2 As Variant, Arg3 As Variant, Optional Arg4 As Variant, Optional Arg5 As Variant, Optional Arg6 As Variant, Optional Arg7 As Variant, Optional Arg8 As Variant, Optional Arg9 As Variant, Optional Arg10 As Variant, Optional Arg11 As Variant, Optional Arg12 As Variant, Optional Arg13 As Variant, Optional Arg14 As Variant, Optional Arg15 As Variant, Optional Arg16 As Variant, Optional Arg17 As Variant, Optional Arg18 As Variant, Optional Arg19 As Variant, Optional Arg20 As Variant, Optional Arg21 As Variant, Optional Arg22 As Variant, Optional Arg23 As Variant, Optional Arg24 As Variant, Optional Arg25 As Variant, Optional Arg26 As Variant, Optional Arg27 As Variant, Optional Arg28 As Variant, Optional Arg29 As Variant, Optional Arg30 As Variant) As Variant",
    ),
    (
        "AmorDegrc",
        "Function AmorDegrc(Arg1 As Variant, Arg2 As Variant, Arg3 As Variant, Arg4 As Variant, Arg5 As Variant, Arg6 As Variant, Optional Arg7 As Variant) As Variant",
    ),
    (
        "AmorLinc",
        "Function AmorLinc(Arg1 As Variant, Arg2 As Variant, Arg3 As Variant, Arg4 As Variant, Arg5 As Variant, Arg6 As Variant, Optional Arg7 As Variant) As Variant",
    ),
    (
        "And",
        "Function And(Arg1 As Variant, Optional Arg2 As Variant, Optional Arg3 As Variant, Optional Arg4 As Variant, Optional Arg5 As Variant, Optional Arg6 As Variant, Optional Arg7 As Variant, Optional Arg8 As Variant, Optional Arg9 As Variant, Optional Arg10 As Variant, Optional Arg11 As Variant, Optional Arg12 As Variant, Optional Arg13 As Variant, Optional Arg14 As Variant, Optional Arg15 As Variant, Optional Arg16 As Variant, Optional Arg17 As Variant, Optional Arg18 As Variant, Optional Arg19 As Variant, Optional Arg20 As Variant, Optional Arg21 As Variant, Optional Arg22 As Variant, Optional Arg23 As Variant, Optional Arg24 As Variant, Optional Arg25 As Variant, Optional Arg26 As Variant, Optional Arg27 As Variant, Optional Arg28 As Variant, Optional Arg29 As Variant, Optional Arg30 As Variant) As Variant",
    ),
    ("Asc", "Function Asc(Arg1 As Variant) As Variant"),
    ("Asin", "Function Asin(Arg1 As Variant) As Variant"),
    ("Asinh", "Function Asinh(Arg1 As Variant) As Variant"),
    ("Atan2", "Function Atan2(Arg1 As Variant, Arg2 As Variant) As Variant"),
    ("Atanh", "Function Atanh(Arg1 As Variant) As Variant"),
    (
        "AveDev",
        "Function AveDev(Arg1 As Variant, Optional Arg2 As Variant, Optional Arg3 As Variant, Optional Arg4 As Variant, Optional Arg5 As Variant, Optional Arg6 As Variant, Optional Arg7 As Variant, Optional Arg8 As Variant, Optional Arg9 As Variant, Optional Arg10 As Variant, Optional Arg11 As Variant, Optional Arg12 As Variant, Optional Arg13 As Variant, Optional Arg14 As Variant, Optional Arg15 As Variant, Optional Arg16 As Variant, Optional Arg17 As Variant, Optional Arg18 As Variant, Optional Arg19 As Variant, Optional Arg20 As Variant, Optional Arg21 As Variant, Optional Arg22 As Variant, Optional Arg23 As Variant, Optional Arg24 As Variant, Optional Arg25 As Variant, Optional Arg26 As Variant, Optional Arg27 As Variant, Optional Arg28 As Variant, Optional Arg29 As Variant, Optional Arg30 As Variant) As Variant",
    ),
    (
        "Average",
        "Function Average(Arg1 As Variant, Optional Arg2 As Variant, Optional Arg3 As Variant, Optional Arg4 As Variant, Optional Arg5 As Variant, Optional Arg6 As Variant, Optional Arg7 As Variant, Optional Arg8 As Variant, Optional Arg9 As Variant, Optional Arg10 As Variant, Optional Arg11 As Variant, Optional Arg12 As Variant, Optional Arg13 As Variant, Optional Arg14 As Variant, Optional Arg15 As Variant, Optional Arg16 As Variant, Optional Arg17 As Variant, Optional Arg18 As Variant, Optional Arg19 As Variant, Optional Arg20 As Variant, Optional Arg21 As Variant, Optional Arg22 As Variant, Optional Arg23 As Variant, Optional Arg24 As Variant, Optional Arg25 As Variant, Optional Arg26 As Variant, Optional Arg27 As Variant, Optional Arg28 As Variant, Optional Arg29 As Variant, Optional Arg30 As Variant) As Variant",
    ),
    ("AverageIf", "Function AverageIf(Arg1 As Variant, Arg2 As Variant, Optional Arg3 As Variant) As Variant"),
    (
        "AverageIfs",
        "Function AverageIfs(Arg1 As Variant, Optional Arg2 As Variant, Optional Arg3 As Variant, Optional Arg4 As Variant, Optional Arg5 As Variant, Optional Arg6 As Variant, Optional Arg7 As Variant, Optional Arg8 As Variant, Optional Arg9 As Variant, Optional Arg10 As Variant, Optional Arg11 As Variant, Optional Arg12 As Variant, Optional Arg13 As Variant, Optional Arg14 As Variant, Optional Arg15 As Variant, Optional Arg16 As Variant, Optional Arg17 As Variant, Optional Arg18 As Variant, Optional Arg19 As Variant, Optional Arg20 As Variant, Optional Arg21 As Variant, Optional Arg22 As Variant, Optional Arg23 As Variant, Optional Arg24 As Variant, Optional Arg25 As Variant, Optional Arg26 As Variant, Optional Arg27 As Variant, Optional Arg28 As Variant, Optional Arg29 As Variant, Optional Arg30 As Variant) As Variant",
    ),
    ("BahtText", "Function BahtText(Arg1 As Variant) As Variant"),
    ("BesselI", "Function BesselI(Arg1 As Variant, Arg2 As Variant) As Variant"),
    ("BesselJ", "Function BesselJ(Arg1 As Variant, Arg2 As Variant) As Variant"),
    ("BesselK", "Function BesselK(Arg1 As Variant, Arg2 As Variant) As Variant"),
    ("BesselY", "Function BesselY(Arg1 As Variant, Arg2 As Variant) As Variant"),
    (
        "Beta_Dist",
        "Function Beta_Dist(Arg1 As Variant, Arg2 As Variant, Arg3 As Variant, Optional Arg4 As Variant, Optional Arg5 As Variant, Optional Arg6 As Variant) As Variant",
    ),
    (
        "Beta_Inv",
        "Function Beta_Inv(Arg1 As Variant, Arg2 As Variant, Arg3 As Variant, Optional Arg4 As Variant, Optional Arg5 As Variant) As Variant",
    ),
    (
        "BetaDist",
        "Function BetaDist(Arg1 As Variant, Arg2 As Variant, Arg3 As Variant, Optional Arg4 As Variant, Optional Arg5 As Variant) As Variant",
    ),
    ("BetaInv", "Function BetaInv(Arg1 As Variant, Arg2 As Variant, Arg3 As Variant, Optional Arg4 As Variant, Optional Arg5 As Variant) As Variant"),
    ("Bin2Dec", "Function Bin2Dec(Arg1 As Variant) As Variant"),
    ("Bin2Hex", "Function Bin2Hex(Arg1 As Variant, Optional Arg2 As Variant) As Variant"),
    ("Bin2Oct", "Function Bin2Oct(Arg1 As Variant, Optional Arg2 As Variant) As Variant"),
    ("Binom_Dist", "Function Binom_Dist(Arg1 As Variant, Arg2 As Variant, Arg3 As Variant, Arg4 As Variant) As Variant"),
    ("Binom_Inv", "Function Binom_Inv(Arg1 As Variant, Arg2 As Variant, Arg3 As Variant) As Variant"),
    ("BinomDist", "Function BinomDist(Arg1 As Variant, Arg2 As Variant, Arg3 As Variant, Arg4 As Variant) As Variant"),
    ("Ceiling", "Function Ceiling(Arg1 As Variant, Arg2 As Variant) As Variant"),
    ("Ceiling_Precise", "Function Ceiling_Precise(Arg1 As Variant, Optional Arg2 As Variant) As Variant"),
    ("ChiDist", "Function ChiDist(Arg1 As Variant, Arg2 As Variant) As Variant"),
    ("ChiInv", "Function ChiInv(Arg1 As Variant, Arg2 As Variant) As Variant"),
    ("ChiSq_Dist", "Function ChiSq_Dist(Arg1 As Variant, Arg2 As Variant, Optional Arg3 As Variant) As Variant"),
    ("ChiSq_Dist_RT", "Function ChiSq_Dist_RT(Arg1 As Variant, Arg2 As Variant) As Variant"),
    ("ChiSq_Inv", "Function ChiSq_Inv(Arg1 As Variant, Arg2 As Variant) As Variant"),
    ("ChiSq_Inv_RT", "Function ChiSq_Inv_RT(Arg1 As Variant, Arg2 As Variant) As Variant"),
    ("ChiSq_Test", "Function ChiSq_Test(Arg1 As Variant, Arg2 As Variant) As Variant"),
    ("ChiTest", "Function ChiTest(Arg1 As Variant, Arg2 As Variant) As Variant"),
    (
        "Choose",
        "Function Choose(Arg1 As Variant, Arg2 As Variant, Optional Arg3 As Variant, Optional Arg4 As Variant, Optional Arg5 As Variant, Optional Arg6 As Variant, Optional Arg7 As Variant, Optional Arg8 As Variant, Optional Arg9 As Variant, Optional Arg10 As Variant, Optional Arg11 As Variant, Optional Arg12 As Variant, Optional Arg13 As Variant, Optional Arg14 As Variant, Optional Arg15 As Variant, Optional Arg16 As Variant, Optional Arg17 As Variant, Optional Arg18 As Variant, Optional Arg19 As Variant, Optional Arg20 As Variant, Optional Arg21 As Variant, Optional Arg22 As Variant, Optional Arg23 As Variant, Optional Arg24 As Variant, Optional Arg25 As Variant, Optional Arg26 As Variant, Optional Arg27 As Variant, Optional Arg28 As Variant, Optional Arg29 As Variant, Optional Arg30 As Variant) As Variant",
    ),
    ("Clean", "Function Clean(Arg1 As Variant) As Variant"),
    ("Combin", "Function Combin(Arg1 As Variant, Arg2 As Variant) As Variant"),
    ("Complex", "Function Complex(Arg1 As Variant, Arg2 As Variant, Optional Arg3 As Variant) As Variant"),
    ("Confidence", "Function Confidence(Arg1 As Variant, Arg2 As Variant, Arg3 As Variant) As Variant"),
    ("Confidence_Norm", "Function Confidence_Norm(Arg1 As Variant, Arg2 As Variant, Arg3 As Variant) As Variant"),
    ("Confidence_T", "Function Confidence_T(Arg1 As Variant, Arg2 As Variant, Arg3 As Variant) As Variant"),
    ("Convert", "Function Convert(Arg1 As Variant, Arg2 As Variant, Arg3 As Variant) As Variant"),
    ("Correl", "Function Correl(Arg1 As Variant, Arg2 As Variant) As Variant"),
    ("Cosh", "Function Cosh(Arg1 As Variant) As Variant"),
    (
        "Count",
        "Function Count(Arg1 As Variant, Optional Arg2 As Variant, Optional Arg3 As Variant, Optional Arg4 As Variant, Optional Arg5 As Variant, Optional Arg6 As Variant, Optional Arg7 As Variant, Optional Arg8 As Variant, Optional Arg9 As Variant, Optional Arg10 As Variant, Optional Arg11 As Variant, Optional Arg12 As Variant, Optional Arg13 As Variant, Optional Arg14 As Variant, Optional Arg15 As Variant, Optional Arg16 As Variant, Optional Arg17 As Variant, Optional Arg18 As Variant, Optional Arg19 As Variant, Optional Arg20 As Variant, Optional Arg21 As Variant, Optional Arg22 As Variant, Optional Arg23 As Variant, Optional Arg24 As Variant, Optional Arg25 As Variant, Optional Arg26 As Variant, Optional Arg27 As Variant, Optional Arg28 As Variant, Optional Arg29 As Variant, Optional Arg30 As Variant) As Variant",
    ),
    (
        "CountA",
        "Function CountA(Arg1 As Variant, Optional Arg2 As Variant, Optional Arg3 As Variant, Optional Arg4 As Variant, Optional Arg5 As Variant, Optional Arg6 As Variant, Optional Arg7 As Variant, Optional Arg8 As Variant, Optional Arg9 As Variant, Optional Arg10 As Variant, Optional Arg11 As Variant, Optional Arg12 As Variant, Optional Arg13 As Variant, Optional Arg14 As Variant, Optional Arg15 As Variant, Optional Arg16 As Variant, Optional Arg17 As Variant, Optional Arg18 As Variant, Optional Arg19 As Variant, Optional Arg20 As Variant, Optional Arg21 As Variant, Optional Arg22 As Variant, Optional Arg23 As Variant, Optional Arg24 As Variant, Optional Arg25 As Variant, Optional Arg26 As Variant, Optional Arg27 As Variant, Optional Arg28 As Variant, Optional Arg29 As Variant, Optional Arg30 As Variant) As Variant",
    ),
    ("CountBlank", "Function CountBlank(Arg1 As Variant) As Variant"),
    ("CountIf", "Function CountIf(Arg1 As Variant, Arg2 As Variant) As Variant"),
    (
        "CountIfs",
        "Function CountIfs(Arg1 As Variant, Optional Arg2 As Variant, Optional Arg3 As Variant, Optional Arg4 As Variant, Optional Arg5 As Variant, Optional Arg6 As Variant, Optional Arg7 As Variant, Optional Arg8 As Variant, Optional Arg9 As Variant, Optional Arg10 As Variant, Optional Arg11 As Variant, Optional Arg12 As Variant, Optional Arg13 As Variant, Optional Arg14 As Variant, Optional Arg15 As Variant, Optional Arg16 As Variant, Optional Arg17 As Variant, Optional Arg18 As Variant, Optional Arg19 As Variant, Optional Arg20 As Variant, Optional Arg21 As Variant, Optional Arg22 As Variant, Optional Arg23 As Variant, Optional Arg24 As Variant, Optional Arg25 As Variant, Optional Arg26 As Variant, Optional Arg27 As Variant, Optional Arg28 As Variant, Optional Arg29 As Variant, Optional Arg30 As Variant) As Variant",
    ),
    ("CoupDayBs", "Function CoupDayBs(Arg1 As Variant, Arg2 As Variant, Arg3 As Variant, Optional Arg4 As Variant) As Variant"),
    ("CoupDays", "Function CoupDays(Arg1 As Variant, Arg2 As Variant, Arg3 As Variant, Optional Arg4 As Variant) As Variant"),
    ("CoupDaysNc", "Function CoupDaysNc(Arg1 As Variant, Arg2 As Variant, Arg3 As Variant, Optional Arg4 As Variant) As Variant"),
    ("CoupNcd", "Function CoupNcd(Arg1 As Variant, Arg2 As Variant, Arg3 As Variant, Optional Arg4 As Variant) As Variant"),
    ("CoupNum", "Function CoupNum(Arg1 As Variant, Arg2 As Variant, Arg3 As Variant, Optional Arg4 As Variant) As Variant"),
    ("CoupPcd", "Function CoupPcd(Arg1 As Variant, Arg2 As Variant, Arg3 As Variant, Optional Arg4 As Variant) As Variant"),
    ("Covar", "Function Covar(Arg1 As Variant, Arg2 As Variant) As Variant"),
    ("Covariance_P", "Function Covariance_P(Arg1 As Variant, Arg2 As Variant) As Variant"),
    ("Covariance_S", "Function Covariance_S(Arg1 As Variant, Arg2 As Variant) As Variant"),
    ("CritBinom", "Function CritBinom(Arg1 As Variant, Arg2 As Variant, Arg3 As Variant) As Variant"),
    ("CumIPmt", "Function CumIPmt(Arg1 As Variant, Arg2 As Variant, Arg3 As Variant, Arg4 As Variant, Arg5 As Variant, Arg6 As Variant) As Variant"),
    (
        "CumPrinc",
        "Function CumPrinc(Arg1 As Variant, Arg2 As Variant, Arg3 As Variant, Arg4 As Variant, Arg5 As Variant, Arg6 As Variant) As Variant",
    ),
    ("DAverage", "Function DAverage(Arg1 As Variant, Arg2 As Variant, Arg3 As Variant) As Variant"),
    ("Days360", "Function Days360(Arg1 As Variant, Optional Arg2 As Variant, Optional Arg3 As Variant) As Variant"),
    ("Db", "Function Db(Arg1 As Variant, Arg2 As Variant, Arg3 As Variant, Arg4 As Variant, Optional Arg5 As Variant) As Variant"),
    ("Dbcs", "Function Dbcs(Arg1 As Variant) As Variant"),
    ("DCount", "Function DCount(Arg1 As Variant, Arg2 As Variant, Arg3 As Variant) As Variant"),
    ("DCountA", "Function DCountA(Arg1 As Variant, Arg2 As Variant, Arg3 As Variant) As Variant"),
    ("Ddb", "Function Ddb(Arg1 As Variant, Arg2 As Variant, Arg3 As Variant, Arg4 As Variant, Optional Arg5 As Variant) As Variant"),
    ("Dec2Bin", "Function Dec2Bin(Arg1 As Variant, Optional Arg2 As Variant) As Variant"),
    ("Dec2Hex", "Function Dec2Hex(Arg1 As Variant, Optional Arg2 As Variant) As Variant"),
    ("Dec2Oct", "Function Dec2Oct(Arg1 As Variant, Optional Arg2 As Variant) As Variant"),
    ("Degrees", "Function Degrees(Arg1 As Variant) As Variant"),
    ("Delta", "Function Delta(Arg1 As Variant, Optional Arg2 As Variant) As Variant"),
    (
        "DevSq",
        "Function DevSq(Arg1 As Variant, Optional Arg2 As Variant, Optional Arg3 As Variant, Optional Arg4 As Variant, Optional Arg5 As Variant, Optional Arg6 As Variant, Optional Arg7 As Variant, Optional Arg8 As Variant, Optional Arg9 As Variant, Optional Arg10 As Variant, Optional Arg11 As Variant, Optional Arg12 As Variant, Optional Arg13 As Variant, Optional Arg14 As Variant, Optional Arg15 As Variant, Optional Arg16 As Variant, Optional Arg17 As Variant, Optional Arg18 As Variant, Optional Arg19 As Variant, Optional Arg20 As Variant, Optional Arg21 As Variant, Optional Arg22 As Variant, Optional Arg23 As Variant, Optional Arg24 As Variant, Optional Arg25 As Variant, Optional Arg26 As Variant, Optional Arg27 As Variant, Optional Arg28 As Variant, Optional Arg29 As Variant, Optional Arg30 As Variant) As Variant",
    ),
    ("DGet", "Function DGet(Arg1 As Variant, Arg2 As Variant, Arg3 As Variant) As Variant"),
    ("Disc", "Function Disc(Arg1 As Variant, Arg2 As Variant, Arg3 As Variant, Arg4 As Variant, Optional Arg5 As Variant) As Variant"),
    ("DMax", "Function DMax(Arg1 As Variant, Arg2 As Variant, Arg3 As Variant) As Variant"),
    ("DMin", "Function DMin(Arg1 As Variant, Arg2 As Variant, Arg3 As Variant) As Variant"),
    ("Dollar", "Function Dollar(Arg1 As Variant, Optional Arg2 As Variant) As Variant"),
    ("DollarDe", "Function DollarDe(Arg1 As Variant, Arg2 As Variant) As Variant"),
    ("DollarFr", "Function DollarFr(Arg1 As Variant, Arg2 As Variant) As Variant"),
    ("DProduct", "Function DProduct(Arg1 As Variant, Arg2 As Variant, Arg3 As Variant) As Variant"),
    ("DStDev", "Function DStDev(Arg1 As Variant, Arg2 As Variant, Arg3 As Variant) As Variant"),
    ("DStDevP", "Function DStDevP(Arg1 As Variant, Arg2 As Variant, Arg3 As Variant) As Variant"),
    ("DSum", "Function DSum(Arg1 As Variant, Arg2 As Variant, Arg3 As Variant) As Variant"),
    (
        "Duration",
        "Function Duration(Arg1 As Variant, Arg2 As Variant, Arg3 As Variant, Arg4 As Variant, Arg5 As Variant, Optional Arg6 As Variant) As Variant",
    ),
    ("DVar", "Function DVar(Arg1 As Variant, Arg2 As Variant, Arg3 As Variant) As Variant"),
    ("DVarP", "Function DVarP(Arg1 As Variant, Arg2 As Variant, Arg3 As Variant) As Variant"),
    ("EDate", "Function EDate(Arg1 As Variant, Arg2 As Variant) As Variant"),
    ("Effect", "Function Effect(Arg1 As Variant, Arg2 As Variant) As Variant"),
    ("EoMonth", "Function EoMonth(Arg1 As Variant, Arg2 As Variant) As Variant"),
    ("Erf", "Function Erf(Arg1 As Variant, Optional Arg2 As Variant) As Variant"),
    ("Erf_Precise", "Function Erf_Precise(Arg1 As Variant) As Variant"),
    ("ErfC", "Function ErfC(Arg1 As Variant, Optional Arg2 As Variant) As Variant"),
    ("ErfC_Precise", "Function ErfC_Precise(Arg1 As Variant) As Variant"),
    ("Even", "Function Even(Arg1 As Variant) As Variant"),
    ("Expon_Dist", "Function Expon_Dist(Arg1 As Variant, Arg2 As Variant, Arg3 As Variant) As Variant"),
    ("ExponDist", "Function ExponDist(Arg1 As Variant, Arg2 As Variant, Arg3 As Variant) As Variant"),
    ("F_Dist", "Function F_Dist(Arg1 As Variant, Arg2 As Variant, Arg3 As Variant, Optional Arg4 As Variant) As Variant"),
    ("F_Dist_RT", "Function F_Dist_RT(Arg1 As Variant, Arg2 As Variant, Arg3 As Variant) As Variant"),
    ("F_Inv", "Function F_Inv(Arg1 As Variant, Arg2 As Variant, Arg3 As Variant) As Variant"),
    ("F_Inv_RT", "Function F_Inv_RT(Arg1 As Variant, Arg2 As Variant, Arg3 As Variant) As Variant"),
    ("F_Test", "Function F_Test(Arg1 As Variant, Arg2 As Variant) As Variant"),
    ("Fact", "Function Fact(Arg1 As Variant) As Variant"),
    ("FactDouble", "Function FactDouble(Arg1 As Variant) As Variant"),
    ("FDist", "Function FDist(Arg1 As Variant, Arg2 As Variant, Arg3 As Variant) As Variant"),
    ("Find", "Function Find(Arg1 As Variant, Arg2 As Variant, Optional Arg3 As Variant) As Variant"),
    ("FindB", "Function FindB(Arg1 As Variant, Arg2 As Variant, Optional Arg3 As Variant) As Variant"),
    ("FInv", "Function FInv(Arg1 As Variant, Arg2 As Variant, Arg3 As Variant) As Variant"),
    ("Fisher", "Function Fisher(Arg1 As Variant) As Variant"),
    ("FisherInv", "Function FisherInv(Arg1 As Variant) As Variant"),
    ("Fixed", "Function Fixed(Arg1 As Variant, Optional Arg2 As Variant, Optional Arg3 As Variant) As Variant"),
    ("Floor", "Function Floor(Arg1 As Variant, Arg2 As Variant) As Variant"),
    ("Floor_Precise", "Function Floor_Precise(Arg1 As Variant, Optional Arg2 As Variant) As Variant"),
    ("Forecast", "Function Forecast(Arg1 As Variant, Arg2 As Variant, Arg3 As Variant) As Variant"),
    ("Frequency", "Function Frequency(Arg1 As Variant, Arg2 As Variant) As Variant"),
    ("FTest", "Function FTest(Arg1 As Variant, Arg2 As Variant) As Variant"),
    ("Fv", "Function Fv(Arg1 As Variant, Arg2 As Variant, Arg3 As Variant, Optional Arg4 As Variant, Optional Arg5 As Variant) As Variant"),
    ("FVSchedule", "Function FVSchedule(Arg1 As Variant, Arg2 As Variant) As Variant"),
    ("Gamma_Dist", "Function Gamma_Dist(Arg1 As Variant, Arg2 As Variant, Arg3 As Variant, Arg4 As Variant) As Variant"),
    ("Gamma_Inv", "Function Gamma_Inv(Arg1 As Variant, Arg2 As Variant, Arg3 As Variant) As Variant"),
    ("GammaDist", "Function GammaDist(Arg1 As Variant, Arg2 As Variant, Arg3 As Variant, Arg4 As Variant) As Variant"),
    ("GammaInv", "Function GammaInv(Arg1 As Variant, Arg2 As Variant, Arg3 As Variant) As Variant"),
    ("GammaLn", "Function GammaLn(Arg1 As Variant) As Variant"),
    ("GammaLn_Precise", "Function GammaLn_Precise(Arg1 As Variant) As Variant"),
    (
        "Gcd",
        "Function Gcd(Arg1 As Variant, Optional Arg2 As Variant, Optional Arg3 As Variant, Optional Arg4 As Variant, Optional Arg5 As Variant, Optional Arg6 As Variant, Optional Arg7 As Variant, Optional Arg8 As Variant, Optional Arg9 As Variant, Optional Arg10 As Variant, Optional Arg11 As Variant, Optional Arg12 As Variant, Optional Arg13 As Variant, Optional Arg14 As Variant, Optional Arg15 As Variant, Optional Arg16 As Variant, Optional Arg17 As Variant, Optional Arg18 As Variant, Optional Arg19 As Variant, Optional Arg20 As Variant, Optional Arg21 As Variant, Optional Arg22 As Variant, Optional Arg23 As Variant, Optional Arg24 As Variant, Optional Arg25 As Variant, Optional Arg26 As Variant, Optional Arg27 As Variant, Optional Arg28 As Variant, Optional Arg29 As Variant, Optional Arg30 As Variant) As Variant",
    ),
    (
        "GeoMean",
        "Function GeoMean(Arg1 As Variant, Optional Arg2 As Variant, Optional Arg3 As Variant, Optional Arg4 As Variant, Optional Arg5 As Variant, Optional Arg6 As Variant, Optional Arg7 As Variant, Optional Arg8 As Variant, Optional Arg9 As Variant, Optional Arg10 As Variant, Optional Arg11 As Variant, Optional Arg12 As Variant, Optional Arg13 As Variant, Optional Arg14 As Variant, Optional Arg15 As Variant, Optional Arg16 As Variant, Optional Arg17 As Variant, Optional Arg18 As Variant, Optional Arg19 As Variant, Optional Arg20 As Variant, Optional Arg21 As Variant, Optional Arg22 As Variant, Optional Arg23 As Variant, Optional Arg24 As Variant, Optional Arg25 As Variant, Optional Arg26 As Variant, Optional Arg27 As Variant, Optional Arg28 As Variant, Optional Arg29 As Variant, Optional Arg30 As Variant) As Variant",
    ),
    ("GeStep", "Function GeStep(Arg1 As Variant, Optional Arg2 As Variant) As Variant"),
    ("Growth", "Function Growth(Arg1 As Variant, Optional Arg2 As Variant, Optional Arg3 As Variant, Optional Arg4 As Variant) As Variant"),
    (
        "HarMean",
        "Function HarMean(Arg1 As Variant, Optional Arg2 As Variant, Optional Arg3 As Variant, Optional Arg4 As Variant, Optional Arg5 As Variant, Optional Arg6 As Variant, Optional Arg7 As Variant, Optional Arg8 As Variant, Optional Arg9 As Variant, Optional Arg10 As Variant, Optional Arg11 As Variant, Optional Arg12 As Variant, Optional Arg13 As Variant, Optional Arg14 As Variant, Optional Arg15 As Variant, Optional Arg16 As Variant, Optional Arg17 As Variant, Optional Arg18 As Variant, Optional Arg19 As Variant, Optional Arg20 As Variant, Optional Arg21 As Variant, Optional Arg22 As Variant, Optional Arg23 As Variant, Optional Arg24 As Variant, Optional Arg25 As Variant, Optional Arg26 As Variant, Optional Arg27 As Variant, Optional Arg28 As Variant, Optional Arg29 As Variant, Optional Arg30 As Variant) As Variant",
    ),
    ("Hex2Bin", "Function Hex2Bin(Arg1 As Variant, Optional Arg2 As Variant) As Variant"),
    ("Hex2Dec", "Function Hex2Dec(Arg1 As Variant) As Variant"),
    ("Hex2Oct", "Function Hex2Oct(Arg1 As Variant, Optional Arg2 As Variant) As Variant"),
    ("HLookup", "Function HLookup(Arg1 As Variant, Arg2 As Variant, Arg3 As Variant, Optional Arg4 As Variant) As Variant"),
    (
        "HypGeom_Dist",
        "Function HypGeom_Dist(Arg1 As Variant, Arg2 As Variant, Arg3 As Variant, Arg4 As Variant, Optional Arg5 As Variant) As Variant",
    ),
    ("HypGeomDist", "Function HypGeomDist(Arg1 As Variant, Arg2 As Variant, Arg3 As Variant, Arg4 As Variant) As Variant"),
    ("IfError", "Function IfError(Arg1 As Variant, Arg2 As Variant) As Variant"),
    ("ImAbs", "Function ImAbs(Arg1 As Variant) As Variant"),
    ("Imaginary", "Function Imaginary(Arg1 As Variant) As Variant"),
    ("ImArgument", "Function ImArgument(Arg1 As Variant) As Variant"),
    ("ImConjugate", "Function ImConjugate(Arg1 As Variant) As Variant"),
    ("ImCos", "Function ImCos(Arg1 As Variant) As Variant"),
    ("ImDiv", "Function ImDiv(Arg1 As Variant, Arg2 As Variant) As Variant"),
    ("ImExp", "Function ImExp(Arg1 As Variant) As Variant"),
    ("ImLn", "Function ImLn(Arg1 As Variant) As Variant"),
    ("ImLog10", "Function ImLog10(Arg1 As Variant) As Variant"),
    ("ImLog2", "Function ImLog2(Arg1 As Variant) As Variant"),
    ("ImPower", "Function ImPower(Arg1 As Variant, Arg2 As Variant) As Variant"),
    (
        "ImProduct",
        "Function ImProduct(Arg1 As Variant, Optional Arg2 As Variant, Optional Arg3 As Variant, Optional Arg4 As Variant, Optional Arg5 As Variant, Optional Arg6 As Variant, Optional Arg7 As Variant, Optional Arg8 As Variant, Optional Arg9 As Variant, Optional Arg10 As Variant, Optional Arg11 As Variant, Optional Arg12 As Variant, Optional Arg13 As Variant, Optional Arg14 As Variant, Optional Arg15 As Variant, Optional Arg16 As Variant, Optional Arg17 As Variant, Optional Arg18 As Variant, Optional Arg19 As Variant, Optional Arg20 As Variant, Optional Arg21 As Variant, Optional Arg22 As Variant, Optional Arg23 As Variant, Optional Arg24 As Variant, Optional Arg25 As Variant, Optional Arg26 As Variant, Optional Arg27 As Variant, Optional Arg28 As Variant, Optional Arg29 As Variant, Optional Arg30 As Variant) As Variant",
    ),
    ("ImReal", "Function ImReal(Arg1 As Variant) As Variant"),
    ("ImSin", "Function ImSin(Arg1 As Variant) As Variant"),
    ("ImSqrt", "Function ImSqrt(Arg1 As Variant) As Variant"),
    ("ImSub", "Function ImSub(Arg1 As Variant, Arg2 As Variant) As Variant"),
    (
        "ImSum",
        "Function ImSum(Arg1 As Variant, Optional Arg2 As Variant, Optional Arg3 As Variant, Optional Arg4 As Variant, Optional Arg5 As Variant, Optional Arg6 As Variant, Optional Arg7 As Variant, Optional Arg8 As Variant, Optional Arg9 As Variant, Optional Arg10 As Variant, Optional Arg11 As Variant, Optional Arg12 As Variant, Optional Arg13 As Variant, Optional Arg14 As Variant, Optional Arg15 As Variant, Optional Arg16 As Variant, Optional Arg17 As Variant, Optional Arg18 As Variant, Optional Arg19 As Variant, Optional Arg20 As Variant, Optional Arg21 As Variant, Optional Arg22 As Variant, Optional Arg23 As Variant, Optional Arg24 As Variant, Optional Arg25 As Variant, Optional Arg26 As Variant, Optional Arg27 As Variant, Optional Arg28 As Variant, Optional Arg29 As Variant, Optional Arg30 As Variant) As Variant",
    ),
    ("Index", "Function Index(Arg1 As Variant, Arg2 As Variant, Optional Arg3 As Variant, Optional Arg4 As Variant) As Variant"),
    ("Intercept", "Function Intercept(Arg1 As Variant, Arg2 As Variant) As Variant"),
    ("IntRate", "Function IntRate(Arg1 As Variant, Arg2 As Variant, Arg3 As Variant, Arg4 As Variant, Optional Arg5 As Variant) As Variant"),
    (
        "Ipmt",
        "Function Ipmt(Arg1 As Variant, Arg2 As Variant, Arg3 As Variant, Arg4 As Variant, Optional Arg5 As Variant, Optional Arg6 As Variant) As Variant",
    ),
    ("Irr", "Function Irr(Arg1 As Variant, Optional Arg2 As Variant) As Variant"),
    ("IsErr", "Function IsErr(Arg1 As Variant) As Variant"),
    ("IsError", "Function IsError(Arg1 As Variant) As Variant"),
    ("IsEven", "Function IsEven(Arg1 As Variant) As Variant"),
    ("IsLogical", "Function IsLogical(Arg1 As Variant) As Variant"),
    ("IsNA", "Function IsNA(Arg1 As Variant) As Variant"),
    ("IsNonText", "Function IsNonText(Arg1 As Variant) As Variant"),
    ("IsNumber", "Function IsNumber(Arg1 As Variant) As Variant"),
    ("ISO_Ceiling", "Function ISO_Ceiling(Arg1 As Variant, Optional Arg2 As Variant) As Variant"),
    ("IsOdd", "Function IsOdd(Arg1 As Variant) As Variant"),
    ("Ispmt", "Function Ispmt(Arg1 As Variant, Arg2 As Variant, Arg3 As Variant, Arg4 As Variant) As Variant"),
    ("IsText", "Function IsText(Arg1 As Variant) As Variant"),
    (
        "Kurt",
        "Function Kurt(Arg1 As Variant, Optional Arg2 As Variant, Optional Arg3 As Variant, Optional Arg4 As Variant, Optional Arg5 As Variant, Optional Arg6 As Variant, Optional Arg7 As Variant, Optional Arg8 As Variant, Optional Arg9 As Variant, Optional Arg10 As Variant, Optional Arg11 As Variant, Optional Arg12 As Variant, Optional Arg13 As Variant, Optional Arg14 As Variant, Optional Arg15 As Variant, Optional Arg16 As Variant, Optional Arg17 As Variant, Optional Arg18 As Variant, Optional Arg19 As Variant, Optional Arg20 As Variant, Optional Arg21 As Variant, Optional Arg22 As Variant, Optional Arg23 As Variant, Optional Arg24 As Variant, Optional Arg25 As Variant, Optional Arg26 As Variant, Optional Arg27 As Variant, Optional Arg28 As Variant, Optional Arg29 As Variant, Optional Arg30 As Variant) As Variant",
    ),
    ("Large", "Function Large(Arg1 As Variant, Arg2 As Variant) As Variant"),
    (
        "Lcm",
        "Function Lcm(Arg1 As Variant, Optional Arg2 As Variant, Optional Arg3 As Variant, Optional Arg4 As Variant, Optional Arg5 As Variant, Optional Arg6 As Variant, Optional Arg7 As Variant, Optional Arg8 As Variant, Optional Arg9 As Variant, Optional Arg10 As Variant, Optional Arg11 As Variant, Optional Arg12 As Variant, Optional Arg13 As Variant, Optional Arg14 As Variant, Optional Arg15 As Variant, Optional Arg16 As Variant, Optional Arg17 As Variant, Optional Arg18 As Variant, Optional Arg19 As Variant, Optional Arg20 As Variant, Optional Arg21 As Variant, Optional Arg22 As Variant, Optional Arg23 As Variant, Optional Arg24 As Variant, Optional Arg25 As Variant, Optional Arg26 As Variant, Optional Arg27 As Variant, Optional Arg28 As Variant, Optional Arg29 As Variant, Optional Arg30 As Variant) As Variant",
    ),
    ("LinEst", "Function LinEst(Arg1 As Variant, Optional Arg2 As Variant, Optional Arg3 As Variant, Optional Arg4 As Variant) As Variant"),
    ("Ln", "Function Ln(Arg1 As Variant) As Variant"),
    ("Log", "Function Log(Arg1 As Variant, Optional Arg2 As Variant) As Variant"),
    ("Log10", "Function Log10(Arg1 As Variant) As Variant"),
    ("LogEst", "Function LogEst(Arg1 As Variant, Optional Arg2 As Variant, Optional Arg3 As Variant, Optional Arg4 As Variant) As Variant"),
    ("LogInv", "Function LogInv(Arg1 As Variant, Arg2 As Variant, Arg3 As Variant) As Variant"),
    ("LogNorm_Dist", "Function LogNorm_Dist(Arg1 As Variant, Arg2 As Variant, Arg3 As Variant, Optional Arg4 As Variant) As Variant"),
    ("LogNorm_Inv", "Function LogNorm_Inv(Arg1 As Variant, Arg2 As Variant, Arg3 As Variant) As Variant"),
    ("LogNormDist", "Function LogNormDist(Arg1 As Variant, Arg2 As Variant, Arg3 As Variant) As Variant"),
    ("Lookup", "Function Lookup(Arg1 As Variant, Arg2 As Variant, Optional Arg3 As Variant) As Variant"),
    ("Match", "Function Match(Arg1 As Variant, Arg2 As Variant, Optional Arg3 As Variant) As Variant"),
    (
        "Max",
        "Function Max(Optional Arg1 As Variant, Optional Arg2 As Variant, Optional Arg3 As Variant, Optional Arg4 As Variant, Optional Arg5 As Variant, Optional Arg6 As Variant, Optional Arg7 As Variant, Optional Arg8 As Variant, Optional Arg9 As Variant, Optional Arg10 As Variant, Optional Arg11 As Variant, Optional Arg12 As Variant, Optional Arg13 As Variant, Optional Arg14 As Variant, Optional Arg15 As Variant, Optional Arg16 As Variant, Optional Arg17 As Variant, Optional Arg18 As Variant, Optional Arg19 As Variant, Optional Arg20 As Variant, Optional Arg21 As Variant, Optional Arg22 As Variant, Optional Arg23 As Variant, Optional Arg24 As Variant, Optional Arg25 As Variant, Optional Arg26 As Variant, Optional Arg27 As Variant, Optional Arg28 As Variant, Optional Arg29 As Variant, Optional Arg30 As Variant) As Variant",
    ),
    ("MDeterm", "Function MDeterm(Arg1 As Variant) As Variant"),
    (
        "MDuration",
        "Function MDuration(Arg1 As Variant, Arg2 As Variant, Arg3 As Variant, Arg4 As Variant, Arg5 As Variant, Optional Arg6 As Variant) As Variant",
    ),
    (
        "Median",
        "Function Median(Arg1 As Variant, Optional Arg2 As Variant, Optional Arg3 As Variant, Optional Arg4 As Variant, Optional Arg5 As Variant, Optional Arg6 As Variant, Optional Arg7 As Variant, Optional Arg8 As Variant, Optional Arg9 As Variant, Optional Arg10 As Variant, Optional Arg11 As Variant, Optional Arg12 As Variant, Optional Arg13 As Variant, Optional Arg14 As Variant, Optional Arg15 As Variant, Optional Arg16 As Variant, Optional Arg17 As Variant, Optional Arg18 As Variant, Optional Arg19 As Variant, Optional Arg20 As Variant, Optional Arg21 As Variant, Optional Arg22 As Variant, Optional Arg23 As Variant, Optional Arg24 As Variant, Optional Arg25 As Variant, Optional Arg26 As Variant, Optional Arg27 As Variant, Optional Arg28 As Variant, Optional Arg29 As Variant, Optional Arg30 As Variant) As Variant",
    ),
    (
        "Min",
        "Function Min(Optional Arg1 As Variant, Optional Arg2 As Variant, Optional Arg3 As Variant, Optional Arg4 As Variant, Optional Arg5 As Variant, Optional Arg6 As Variant, Optional Arg7 As Variant, Optional Arg8 As Variant, Optional Arg9 As Variant, Optional Arg10 As Variant, Optional Arg11 As Variant, Optional Arg12 As Variant, Optional Arg13 As Variant, Optional Arg14 As Variant, Optional Arg15 As Variant, Optional Arg16 As Variant, Optional Arg17 As Variant, Optional Arg18 As Variant, Optional Arg19 As Variant, Optional Arg20 As Variant, Optional Arg21 As Variant, Optional Arg22 As Variant, Optional Arg23 As Variant, Optional Arg24 As Variant, Optional Arg25 As Variant, Optional Arg26 As Variant, Optional Arg27 As Variant, Optional Arg28 As Variant, Optional Arg29 As Variant, Optional Arg30 As Variant) As Variant",
    ),
    ("MInverse", "Function MInverse(Arg1 As Variant) As Variant"),
    ("MIrr", "Function MIrr(Arg1 As Variant, Arg2 As Variant, Arg3 As Variant) As Variant"),
    ("MMult", "Function MMult(Arg1 As Variant, Optional Arg2 As Variant) As Variant"),
    (
        "Mode",
        "Function Mode(Arg1 As Variant, Optional Arg2 As Variant, Optional Arg3 As Variant, Optional Arg4 As Variant, Optional Arg5 As Variant, Optional Arg6 As Variant, Optional Arg7 As Variant, Optional Arg8 As Variant, Optional Arg9 As Variant, Optional Arg10 As Variant, Optional Arg11 As Variant, Optional Arg12 As Variant, Optional Arg13 As Variant, Optional Arg14 As Variant, Optional Arg15 As Variant, Optional Arg16 As Variant, Optional Arg17 As Variant, Optional Arg18 As Variant, Optional Arg19 As Variant, Optional Arg20 As Variant, Optional Arg21 As Variant, Optional Arg22 As Variant, Optional Arg23 As Variant, Optional Arg24 As Variant, Optional Arg25 As Variant, Optional Arg26 As Variant, Optional Arg27 As Variant, Optional Arg28 As Variant, Optional Arg29 As Variant, Optional Arg30 As Variant) As Variant",
    ),
    (
        "Mode_Mult",
        "Function Mode_Mult(Arg1 As Variant, Optional Arg2 As Variant, Optional Arg3 As Variant, Optional Arg4 As Variant, Optional Arg5 As Variant, Optional Arg6 As Variant, Optional Arg7 As Variant, Optional Arg8 As Variant, Optional Arg9 As Variant, Optional Arg10 As Variant, Optional Arg11 As Variant, Optional Arg12 As Variant, Optional Arg13 As Variant, Optional Arg14 As Variant, Optional Arg15 As Variant, Optional Arg16 As Variant, Optional Arg17 As Variant, Optional Arg18 As Variant, Optional Arg19 As Variant, Optional Arg20 As Variant, Optional Arg21 As Variant, Optional Arg22 As Variant, Optional Arg23 As Variant, Optional Arg24 As Variant, Optional Arg25 As Variant, Optional Arg26 As Variant, Optional Arg27 As Variant, Optional Arg28 As Variant, Optional Arg29 As Variant, Optional Arg30 As Variant) As Variant",
    ),
    (
        "Mode_Sngl",
        "Function Mode_Sngl(Arg1 As Variant, Optional Arg2 As Variant, Optional Arg3 As Variant, Optional Arg4 As Variant, Optional Arg5 As Variant, Optional Arg6 As Variant, Optional Arg7 As Variant, Optional Arg8 As Variant, Optional Arg9 As Variant, Optional Arg10 As Variant, Optional Arg11 As Variant, Optional Arg12 As Variant, Optional Arg13 As Variant, Optional Arg14 As Variant, Optional Arg15 As Variant, Optional Arg16 As Variant, Optional Arg17 As Variant, Optional Arg18 As Variant, Optional Arg19 As Variant, Optional Arg20 As Variant, Optional Arg21 As Variant, Optional Arg22 As Variant, Optional Arg23 As Variant, Optional Arg24 As Variant, Optional Arg25 As Variant, Optional Arg26 As Variant, Optional Arg27 As Variant, Optional Arg28 As Variant, Optional Arg29 As Variant, Optional Arg30 As Variant) As Variant",
    ),
    ("MRound", "Function MRound(Arg1 As Variant, Arg2 As Variant) As Variant"),
    (
        "MultiNomial",
        "Function MultiNomial(Arg1 As Variant, Optional Arg2 As Variant, Optional Arg3 As Variant, Optional Arg4 As Variant, Optional Arg5 As Variant, Optional Arg6 As Variant, Optional Arg7 As Variant, Optional Arg8 As Variant, Optional Arg9 As Variant, Optional Arg10 As Variant, Optional Arg11 As Variant, Optional Arg12 As Variant, Optional Arg13 As Variant, Optional Arg14 As Variant, Optional Arg15 As Variant, Optional Arg16 As Variant, Optional Arg17 As Variant, Optional Arg18 As Variant, Optional Arg19 As Variant, Optional Arg20 As Variant, Optional Arg21 As Variant, Optional Arg22 As Variant, Optional Arg23 As Variant, Optional Arg24 As Variant, Optional Arg25 As Variant, Optional Arg26 As Variant, Optional Arg27 As Variant, Optional Arg28 As Variant, Optional Arg29 As Variant, Optional Arg30 As Variant) As Variant",
    ),
    ("NegBinom_Dist", "Function NegBinom_Dist(Arg1 As Variant, Arg2 As Variant, Arg3 As Variant, Optional Arg4 As Variant) As Variant"),
    ("NegBinomDist", "Function NegBinomDist(Arg1 As Variant, Arg2 As Variant, Arg3 As Variant) As Variant"),
    ("NetworkDays", "Function NetworkDays(Arg1 As Variant, Arg2 As Variant, Optional Arg3 As Variant) As Variant"),
    (
        "NetworkDays_Intl",
        "Function NetworkDays_Intl(Arg1 As Variant, Arg2 As Variant, Optional Arg3 As Variant, Optional Arg4 As Variant) As Variant",
    ),
    ("Nominal", "Function Nominal(Arg1 As Variant, Arg2 As Variant) As Variant"),
    ("Norm_Dist", "Function Norm_Dist(Arg1 As Variant, Arg2 As Variant, Arg3 As Variant, Arg4 As Variant) As Variant"),
    ("Norm_Inv", "Function Norm_Inv(Arg1 As Variant, Arg2 As Variant, Arg3 As Variant) As Variant"),
    ("Norm_S_Dist", "Function Norm_S_Dist(Arg1 As Variant, Optional Arg2 As Variant) As Variant"),
    ("Norm_S_Inv", "Function Norm_S_Inv(Arg1 As Variant) As Variant"),
    ("NormDist", "Function NormDist(Arg1 As Variant, Arg2 As Variant, Arg3 As Variant, Arg4 As Variant) As Variant"),
    ("NormInv", "Function NormInv(Arg1 As Variant, Arg2 As Variant, Arg3 As Variant) As Variant"),
    ("NormSDist", "Function NormSDist(Arg1 As Variant) As Variant"),
    ("NormSInv", "Function NormSInv(Arg1 As Variant) As Variant"),
    ("NPer", "Function NPer(Arg1 As Variant, Arg2 As Variant, Arg3 As Variant, Optional Arg4 As Variant, Optional Arg5 As Variant) As Variant"),
    (
        "Npv",
        "Function Npv(Arg1 As Variant, Arg2 As Variant, Optional Arg3 As Variant, Optional Arg4 As Variant, Optional Arg5 As Variant, Optional Arg6 As Variant, Optional Arg7 As Variant, Optional Arg8 As Variant, Optional Arg9 As Variant, Optional Arg10 As Variant, Optional Arg11 As Variant, Optional Arg12 As Variant, Optional Arg13 As Variant, Optional Arg14 As Variant, Optional Arg15 As Variant, Optional Arg16 As Variant, Optional Arg17 As Variant, Optional Arg18 As Variant, Optional Arg19 As Variant, Optional Arg20 As Variant, Optional Arg21 As Variant, Optional Arg22 As Variant, Optional Arg23 As Variant, Optional Arg24 As Variant, Optional Arg25 As Variant, Optional Arg26 As Variant, Optional Arg27 As Variant, Optional Arg28 As Variant, Optional Arg29 As Variant, Optional Arg30 As Variant) As Variant",
    ),
    ("Oct2Bin", "Function Oct2Bin(Arg1 As Variant, Optional Arg2 As Variant) As Variant"),
    ("Oct2Dec", "Function Oct2Dec(Arg1 As Variant) As Variant"),
    ("Oct2Hex", "Function Oct2Hex(Arg1 As Variant, Optional Arg2 As Variant) As Variant"),
    ("Odd", "Function Odd(Arg1 As Variant) As Variant"),
    (
        "OddFPrice",
        "Function OddFPrice(Arg1 As Variant, Arg2 As Variant, Arg3 As Variant, Arg4 As Variant, Arg5 As Variant, Arg6 As Variant, Arg7 As Variant, Arg8 As Variant, Optional Arg9 As Variant) As Variant",
    ),
    (
        "OddFYield",
        "Function OddFYield(Arg1 As Variant, Arg2 As Variant, Arg3 As Variant, Arg4 As Variant, Arg5 As Variant, Arg6 As Variant, Arg7 As Variant, Arg8 As Variant, Optional Arg9 As Variant) As Variant",
    ),
    (
        "OddLPrice",
        "Function OddLPrice(Arg1 As Variant, Arg2 As Variant, Arg3 As Variant, Arg4 As Variant, Arg5 As Variant, Arg6 As Variant, Arg7 As Variant, Optional Arg8 As Variant) As Variant",
    ),
    (
        "OddLYield",
        "Function OddLYield(Arg1 As Variant, Arg2 As Variant, Arg3 As Variant, Arg4 As Variant, Arg5 As Variant, Arg6 As Variant, Arg7 As Variant, Optional Arg8 As Variant) As Variant",
    ),
    (
        "Or",
        "Function Or(Arg1 As Variant, Optional Arg2 As Variant, Optional Arg3 As Variant, Optional Arg4 As Variant, Optional Arg5 As Variant, Optional Arg6 As Variant, Optional Arg7 As Variant, Optional Arg8 As Variant, Optional Arg9 As Variant, Optional Arg10 As Variant, Optional Arg11 As Variant, Optional Arg12 As Variant, Optional Arg13 As Variant, Optional Arg14 As Variant, Optional Arg15 As Variant, Optional Arg16 As Variant, Optional Arg17 As Variant, Optional Arg18 As Variant, Optional Arg19 As Variant, Optional Arg20 As Variant, Optional Arg21 As Variant, Optional Arg22 As Variant, Optional Arg23 As Variant, Optional Arg24 As Variant, Optional Arg25 As Variant, Optional Arg26 As Variant, Optional Arg27 As Variant, Optional Arg28 As Variant, Optional Arg29 As Variant, Optional Arg30 As Variant) As Variant",
    ),
    ("Pearson", "Function Pearson(Arg1 As Variant, Arg2 As Variant) As Variant"),
    ("Percentile", "Function Percentile(Arg1 As Variant, Arg2 As Variant) As Variant"),
    ("Percentile_Exc", "Function Percentile_Exc(Arg1 As Variant, Arg2 As Variant) As Variant"),
    ("Percentile_Inc", "Function Percentile_Inc(Arg1 As Variant, Arg2 As Variant) As Variant"),
    ("PercentRank", "Function PercentRank(Arg1 As Variant, Arg2 As Variant, Optional Arg3 As Variant) As Variant"),
    ("PercentRank_Exc", "Function PercentRank_Exc(Arg1 As Variant, Arg2 As Variant, Optional Arg3 As Variant) As Variant"),
    ("PercentRank_Inc", "Function PercentRank_Inc(Arg1 As Variant, Arg2 As Variant, Optional Arg3 As Variant) As Variant"),
    ("Permut", "Function Permut(Arg1 As Variant, Arg2 As Variant) As Variant"),
    ("Phonetic", "Function Phonetic(Arg1 As Variant) As Variant"),
    ("Pi", "Function Pi() As Variant"),
    ("Pmt", "Function Pmt(Arg1 As Variant, Arg2 As Variant, Arg3 As Variant, Optional Arg4 As Variant, Optional Arg5 As Variant) As Variant"),
    ("Poisson", "Function Poisson(Arg1 As Variant, Arg2 As Variant, Arg3 As Variant) As Variant"),
    ("Poisson_Dist", "Function Poisson_Dist(Arg1 As Variant, Arg2 As Variant, Arg3 As Variant) As Variant"),
    ("Power", "Function Power(Arg1 As Variant, Arg2 As Variant) As Variant"),
    (
        "Ppmt",
        "Function Ppmt(Arg1 As Variant, Arg2 As Variant, Arg3 As Variant, Arg4 As Variant, Optional Arg5 As Variant, Optional Arg6 As Variant) As Variant",
    ),
    (
        "Price",
        "Function Price(Arg1 As Variant, Arg2 As Variant, Arg3 As Variant, Arg4 As Variant, Arg5 As Variant, Arg6 As Variant, Optional Arg7 As Variant) As Variant",
    ),
    ("PriceDisc", "Function PriceDisc(Arg1 As Variant, Arg2 As Variant, Arg3 As Variant, Arg4 As Variant, Optional Arg5 As Variant) As Variant"),
    (
        "PriceMat",
        "Function PriceMat(Arg1 As Variant, Arg2 As Variant, Arg3 As Variant, Arg4 As Variant, Arg5 As Variant, Optional Arg6 As Variant) As Variant",
    ),
    ("Prob", "Function Prob(Arg1 As Variant, Arg2 As Variant, Arg3 As Variant, Optional Arg4 As Variant) As Variant"),
    (
        "Product",
        "Function Product(Arg1 As Variant, Optional Arg2 As Variant, Optional Arg3 As Variant, Optional Arg4 As Variant, Optional Arg5 As Variant, Optional Arg6 As Variant, Optional Arg7 As Variant, Optional Arg8 As Variant, Optional Arg9 As Variant, Optional Arg10 As Variant, Optional Arg11 As Variant, Optional Arg12 As Variant, Optional Arg13 As Variant, Optional Arg14 As Variant, Optional Arg15 As Variant, Optional Arg16 As Variant, Optional Arg17 As Variant, Optional Arg18 As Variant, Optional Arg19 As Variant, Optional Arg20 As Variant, Optional Arg21 As Variant, Optional Arg22 As Variant, Optional Arg23 As Variant, Optional Arg24 As Variant, Optional Arg25 As Variant, Optional Arg26 As Variant, Optional Arg27 As Variant, Optional Arg28 As Variant, Optional Arg29 As Variant, Optional Arg30 As Variant) As Variant",
    ),
    ("Proper", "Function Proper(Arg1 As Variant) As Variant"),
    ("Pv", "Function Pv(Arg1 As Variant, Arg2 As Variant, Arg3 As Variant, Optional Arg4 As Variant, Optional Arg5 As Variant) As Variant"),
    ("Quartile", "Function Quartile(Arg1 As Variant, Arg2 As Variant) As Variant"),
    ("Quartile_Exc", "Function Quartile_Exc(Arg1 As Variant, Arg2 As Variant) As Variant"),
    ("Quartile_Inc", "Function Quartile_Inc(Arg1 As Variant, Arg2 As Variant) As Variant"),
    ("Quotient", "Function Quotient(Arg1 As Variant, Arg2 As Variant) As Variant"),
    ("Radians", "Function Radians(Arg1 As Variant) As Variant"),
    ("RandBetween", "Function RandBetween(Arg1 As Variant, Arg2 As Variant) As Variant"),
    ("Rank", "Function Rank(Arg1 As Variant, Arg2 As Variant, Optional Arg3 As Variant) As Variant"),
    ("Rank_Avg", "Function Rank_Avg(Arg1 As Variant, Arg2 As Variant, Optional Arg3 As Variant) As Variant"),
    ("Rank_Eq", "Function Rank_Eq(Arg1 As Variant, Arg2 As Variant, Optional Arg3 As Variant) As Variant"),
    (
        "Rate",
        "Function Rate(Arg1 As Variant, Arg2 As Variant, Arg3 As Variant, Optional Arg4 As Variant, Optional Arg5 As Variant, Optional Arg6 As Variant) As Variant",
    ),
    ("Received", "Function Received(Arg1 As Variant, Arg2 As Variant, Arg3 As Variant, Arg4 As Variant, Optional Arg5 As Variant) As Variant"),
    ("Replace", "Function Replace(Arg1 As Variant, Arg2 As Variant, Arg3 As Variant, Arg4 As Variant) As Variant"),
    ("ReplaceB", "Function ReplaceB(Arg1 As Variant, Arg2 As Variant, Arg3 As Variant, Arg4 As Variant) As Variant"),
    ("Rept", "Function Rept(Arg1 As Variant, Arg2 As Variant) As Variant"),
    ("Roman", "Function Roman(Arg1 As Variant, Optional Arg2 As Variant) As Variant"),
    ("Round", "Function Round(Arg1 As Variant, Arg2 As Variant) As Variant"),
    ("RoundDown", "Function RoundDown(Arg1 As Variant, Arg2 As Variant) As Variant"),
    ("RoundUp", "Function RoundUp(Arg1 As Variant, Arg2 As Variant) As Variant"),
    ("RSq", "Function RSq(Arg1 As Variant, Arg2 As Variant) As Variant"),
    (
        "RTD",
        "Function RTD(progID As Variant, server As Variant, topic1 As Variant, Optional topic2 As Variant, Optional topic3 As Variant, Optional topic4 As Variant, Optional topic5 As Variant, Optional topic6 As Variant, Optional topic7 As Variant, Optional topic8 As Variant, Optional topic9 As Variant, Optional topic10 As Variant, Optional topic11 As Variant, Optional topic12 As Variant, Optional topic13 As Variant, Optional topic14 As Variant, Optional topic15 As Variant, Optional topic16 As Variant, Optional topic17 As Variant, Optional topic18 As Variant, Optional topic19 As Variant, Optional topic20 As Variant, Optional topic21 As Variant, Optional topic22 As Variant, Optional topic23 As Variant, Optional topic24 As Variant, Optional topic25 As Variant, Optional topic26 As Variant, Optional topic27 As Variant, Optional topic28 As Variant) As Variant",
    ),
    ("Search", "Function Search(Arg1 As Variant, Arg2 As Variant, Optional Arg3 As Variant) As Variant"),
    ("SearchB", "Function SearchB(Arg1 As Variant, Arg2 As Variant, Optional Arg3 As Variant) As Variant"),
    ("SeriesSum", "Function SeriesSum(Arg1 As Variant, Arg2 As Variant, Arg3 As Variant, Arg4 As Variant) As Variant"),
    ("Sinh", "Function Sinh(Arg1 As Variant) As Variant"),
    (
        "Skew",
        "Function Skew(Arg1 As Variant, Optional Arg2 As Variant, Optional Arg3 As Variant, Optional Arg4 As Variant, Optional Arg5 As Variant, Optional Arg6 As Variant, Optional Arg7 As Variant, Optional Arg8 As Variant, Optional Arg9 As Variant, Optional Arg10 As Variant, Optional Arg11 As Variant, Optional Arg12 As Variant, Optional Arg13 As Variant, Optional Arg14 As Variant, Optional Arg15 As Variant, Optional Arg16 As Variant, Optional Arg17 As Variant, Optional Arg18 As Variant, Optional Arg19 As Variant, Optional Arg20 As Variant, Optional Arg21 As Variant, Optional Arg22 As Variant, Optional Arg23 As Variant, Optional Arg24 As Variant, Optional Arg25 As Variant, Optional Arg26 As Variant, Optional Arg27 As Variant, Optional Arg28 As Variant, Optional Arg29 As Variant, Optional Arg30 As Variant) As Variant",
    ),
    ("Sln", "Function Sln(Arg1 As Variant, Arg2 As Variant, Arg3 As Variant) As Variant"),
    ("Slope", "Function Slope(Arg1 As Variant, Arg2 As Variant) As Variant"),
    ("Small", "Function Small(Arg1 As Variant, Arg2 As Variant) As Variant"),
    ("SqrtPi", "Function SqrtPi(Arg1 As Variant) As Variant"),
    ("Standardize", "Function Standardize(Arg1 As Variant, Arg2 As Variant, Arg3 As Variant) As Variant"),
    (
        "StDev",
        "Function StDev(Arg1 As Variant, Optional Arg2 As Variant, Optional Arg3 As Variant, Optional Arg4 As Variant, Optional Arg5 As Variant, Optional Arg6 As Variant, Optional Arg7 As Variant, Optional Arg8 As Variant, Optional Arg9 As Variant, Optional Arg10 As Variant, Optional Arg11 As Variant, Optional Arg12 As Variant, Optional Arg13 As Variant, Optional Arg14 As Variant, Optional Arg15 As Variant, Optional Arg16 As Variant, Optional Arg17 As Variant, Optional Arg18 As Variant, Optional Arg19 As Variant, Optional Arg20 As Variant, Optional Arg21 As Variant, Optional Arg22 As Variant, Optional Arg23 As Variant, Optional Arg24 As Variant, Optional Arg25 As Variant, Optional Arg26 As Variant, Optional Arg27 As Variant, Optional Arg28 As Variant, Optional Arg29 As Variant, Optional Arg30 As Variant) As Variant",
    ),
    (
        "StDev_P",
        "Function StDev_P(Arg1 As Variant, Optional Arg2 As Variant, Optional Arg3 As Variant, Optional Arg4 As Variant, Optional Arg5 As Variant, Optional Arg6 As Variant, Optional Arg7 As Variant, Optional Arg8 As Variant, Optional Arg9 As Variant, Optional Arg10 As Variant, Optional Arg11 As Variant, Optional Arg12 As Variant, Optional Arg13 As Variant, Optional Arg14 As Variant, Optional Arg15 As Variant, Optional Arg16 As Variant, Optional Arg17 As Variant, Optional Arg18 As Variant, Optional Arg19 As Variant, Optional Arg20 As Variant, Optional Arg21 As Variant, Optional Arg22 As Variant, Optional Arg23 As Variant, Optional Arg24 As Variant, Optional Arg25 As Variant, Optional Arg26 As Variant, Optional Arg27 As Variant, Optional Arg28 As Variant, Optional Arg29 As Variant, Optional Arg30 As Variant) As Variant",
    ),
    (
        "StDev_S",
        "Function StDev_S(Arg1 As Variant, Optional Arg2 As Variant, Optional Arg3 As Variant, Optional Arg4 As Variant, Optional Arg5 As Variant, Optional Arg6 As Variant, Optional Arg7 As Variant, Optional Arg8 As Variant, Optional Arg9 As Variant, Optional Arg10 As Variant, Optional Arg11 As Variant, Optional Arg12 As Variant, Optional Arg13 As Variant, Optional Arg14 As Variant, Optional Arg15 As Variant, Optional Arg16 As Variant, Optional Arg17 As Variant, Optional Arg18 As Variant, Optional Arg19 As Variant, Optional Arg20 As Variant, Optional Arg21 As Variant, Optional Arg22 As Variant, Optional Arg23 As Variant, Optional Arg24 As Variant, Optional Arg25 As Variant, Optional Arg26 As Variant, Optional Arg27 As Variant, Optional Arg28 As Variant, Optional Arg29 As Variant, Optional Arg30 As Variant) As Variant",
    ),
    (
        "StDevP",
        "Function StDevP(Arg1 As Variant, Optional Arg2 As Variant, Optional Arg3 As Variant, Optional Arg4 As Variant, Optional Arg5 As Variant, Optional Arg6 As Variant, Optional Arg7 As Variant, Optional Arg8 As Variant, Optional Arg9 As Variant, Optional Arg10 As Variant, Optional Arg11 As Variant, Optional Arg12 As Variant, Optional Arg13 As Variant, Optional Arg14 As Variant, Optional Arg15 As Variant, Optional Arg16 As Variant, Optional Arg17 As Variant, Optional Arg18 As Variant, Optional Arg19 As Variant, Optional Arg20 As Variant, Optional Arg21 As Variant, Optional Arg22 As Variant, Optional Arg23 As Variant, Optional Arg24 As Variant, Optional Arg25 As Variant, Optional Arg26 As Variant, Optional Arg27 As Variant, Optional Arg28 As Variant, Optional Arg29 As Variant, Optional Arg30 As Variant) As Variant",
    ),
    ("StEyx", "Function StEyx(Arg1 As Variant, Arg2 As Variant) As Variant"),
    ("Substitute", "Function Substitute(Arg1 As Variant, Arg2 As Variant, Arg3 As Variant, Optional Arg4 As Variant) As Variant"),
    (
        "Subtotal",
        "Function Subtotal(Arg1 As Variant, Arg2 As Variant, Optional Arg3 As Variant, Optional Arg4 As Variant, Optional Arg5 As Variant, Optional Arg6 As Variant, Optional Arg7 As Variant, Optional Arg8 As Variant, Optional Arg9 As Variant, Optional Arg10 As Variant, Optional Arg11 As Variant, Optional Arg12 As Variant, Optional Arg13 As Variant, Optional Arg14 As Variant, Optional Arg15 As Variant, Optional Arg16 As Variant, Optional Arg17 As Variant, Optional Arg18 As Variant, Optional Arg19 As Variant, Optional Arg20 As Variant, Optional Arg21 As Variant, Optional Arg22 As Variant, Optional Arg23 As Variant, Optional Arg24 As Variant, Optional Arg25 As Variant, Optional Arg26 As Variant, Optional Arg27 As Variant, Optional Arg28 As Variant, Optional Arg29 As Variant, Optional Arg30 As Variant) As Variant",
    ),
    (
        "Sum",
        "Function Sum(Arg1 As Variant, Optional Arg2 As Variant, Optional Arg3 As Variant, Optional Arg4 As Variant, Optional Arg5 As Variant, Optional Arg6 As Variant, Optional Arg7 As Variant, Optional Arg8 As Variant, Optional Arg9 As Variant, Optional Arg10 As Variant, Optional Arg11 As Variant, Optional Arg12 As Variant, Optional Arg13 As Variant, Optional Arg14 As Variant, Optional Arg15 As Variant, Optional Arg16 As Variant, Optional Arg17 As Variant, Optional Arg18 As Variant, Optional Arg19 As Variant, Optional Arg20 As Variant, Optional Arg21 As Variant, Optional Arg22 As Variant, Optional Arg23 As Variant, Optional Arg24 As Variant, Optional Arg25 As Variant, Optional Arg26 As Variant, Optional Arg27 As Variant, Optional Arg28 As Variant, Optional Arg29 As Variant, Optional Arg30 As Variant) As Variant",
    ),
    ("SumIf", "Function SumIf(Arg1 As Variant, Arg2 As Variant, Optional Arg3 As Variant) As Variant"),
    (
        "SumIfs",
        "Function SumIfs(Arg1 As Variant, Arg2 As Variant, Arg3 As Variant, Optional Arg4 As Variant, Optional Arg5 As Variant, Optional Arg6 As Variant, Optional Arg7 As Variant, Optional Arg8 As Variant, Optional Arg9 As Variant, Optional Arg10 As Variant, Optional Arg11 As Variant, Optional Arg12 As Variant, Optional Arg13 As Variant, Optional Arg14 As Variant, Optional Arg15 As Variant, Optional Arg16 As Variant, Optional Arg17 As Variant, Optional Arg18 As Variant, Optional Arg19 As Variant, Optional Arg20 As Variant, Optional Arg21 As Variant, Optional Arg22 As Variant, Optional Arg23 As Variant, Optional Arg24 As Variant, Optional Arg25 As Variant, Optional Arg26 As Variant, Optional Arg27 As Variant, Optional Arg28 As Variant, Optional Arg29 As Variant, Optional Arg30 As Variant) As Variant",
    ),
    (
        "SumProduct",
        "Function SumProduct(Arg1 As Variant, Optional Arg2 As Variant, Optional Arg3 As Variant, Optional Arg4 As Variant, Optional Arg5 As Variant, Optional Arg6 As Variant, Optional Arg7 As Variant, Optional Arg8 As Variant, Optional Arg9 As Variant, Optional Arg10 As Variant, Optional Arg11 As Variant, Optional Arg12 As Variant, Optional Arg13 As Variant, Optional Arg14 As Variant, Optional Arg15 As Variant, Optional Arg16 As Variant, Optional Arg17 As Variant, Optional Arg18 As Variant, Optional Arg19 As Variant, Optional Arg20 As Variant, Optional Arg21 As Variant, Optional Arg22 As Variant, Optional Arg23 As Variant, Optional Arg24 As Variant, Optional Arg25 As Variant, Optional Arg26 As Variant, Optional Arg27 As Variant, Optional Arg28 As Variant, Optional Arg29 As Variant, Optional Arg30 As Variant) As Variant",
    ),
    (
        "SumSq",
        "Function SumSq(Arg1 As Variant, Optional Arg2 As Variant, Optional Arg3 As Variant, Optional Arg4 As Variant, Optional Arg5 As Variant, Optional Arg6 As Variant, Optional Arg7 As Variant, Optional Arg8 As Variant, Optional Arg9 As Variant, Optional Arg10 As Variant, Optional Arg11 As Variant, Optional Arg12 As Variant, Optional Arg13 As Variant, Optional Arg14 As Variant, Optional Arg15 As Variant, Optional Arg16 As Variant, Optional Arg17 As Variant, Optional Arg18 As Variant, Optional Arg19 As Variant, Optional Arg20 As Variant, Optional Arg21 As Variant, Optional Arg22 As Variant, Optional Arg23 As Variant, Optional Arg24 As Variant, Optional Arg25 As Variant, Optional Arg26 As Variant, Optional Arg27 As Variant, Optional Arg28 As Variant, Optional Arg29 As Variant, Optional Arg30 As Variant) As Variant",
    ),
    ("SumX2MY2", "Function SumX2MY2(Arg1 As Variant, Arg2 As Variant) As Variant"),
    ("SumX2PY2", "Function SumX2PY2(Arg1 As Variant, Arg2 As Variant) As Variant"),
    ("SumXMY2", "Function SumXMY2(Arg1 As Variant, Arg2 As Variant) As Variant"),
    ("Syd", "Function Syd(Arg1 As Variant, Arg2 As Variant, Arg3 As Variant, Arg4 As Variant) As Variant"),
    ("T_Dist", "Function T_Dist(Arg1 As Variant, Arg2 As Variant, Arg3 As Variant) As Variant"),
    ("T_Dist_2T", "Function T_Dist_2T(Arg1 As Variant, Arg2 As Variant) As Variant"),
    ("T_Dist_RT", "Function T_Dist_RT(Arg1 As Variant, Arg2 As Variant) As Variant"),
    ("T_Inv", "Function T_Inv(Arg1 As Variant, Arg2 As Variant) As Variant"),
    ("T_Inv_2T", "Function T_Inv_2T(Arg1 As Variant, Arg2 As Variant) As Variant"),
    ("T_Test", "Function T_Test(Arg1 As Variant, Arg2 As Variant, Arg3 As Variant, Arg4 As Variant) As Variant"),
    ("Tanh", "Function Tanh(Arg1 As Variant) As Variant"),
    ("TBillEq", "Function TBillEq(Arg1 As Variant, Arg2 As Variant, Optional Arg3 As Variant) As Variant"),
    ("TBillPrice", "Function TBillPrice(Arg1 As Variant, Arg2 As Variant, Optional Arg3 As Variant) As Variant"),
    ("TBillYield", "Function TBillYield(Arg1 As Variant, Arg2 As Variant, Optional Arg3 As Variant) As Variant"),
    ("TDist", "Function TDist(Arg1 As Variant, Arg2 As Variant, Arg3 As Variant) As Variant"),
    ("Text", "Function Text(Arg1 As Variant, Arg2 As Variant) As Variant"),
    ("TInv", "Function TInv(Arg1 As Variant, Arg2 As Variant) As Variant"),
    ("Transpose", "Function Transpose(Arg1 As Variant) As Variant"),
    ("Trend", "Function Trend(Arg1 As Variant, Optional Arg2 As Variant, Optional Arg3 As Variant, Optional Arg4 As Variant) As Variant"),
    ("Trim", "Function Trim(Arg1 As Variant) As Variant"),
    ("TrimMean", "Function TrimMean(Arg1 As Variant, Arg2 As Variant) As Variant"),
    ("TTest", "Function TTest(Arg1 As Variant, Arg2 As Variant, Arg3 As Variant, Arg4 As Variant) As Variant"),
    ("USDollar", "Function USDollar(Arg1 As Variant, Arg2 As Variant) As Variant"),
    (
        "Var",
        "Function Var(Arg1 As Variant, Optional Arg2 As Variant, Optional Arg3 As Variant, Optional Arg4 As Variant, Optional Arg5 As Variant, Optional Arg6 As Variant, Optional Arg7 As Variant, Optional Arg8 As Variant, Optional Arg9 As Variant, Optional Arg10 As Variant, Optional Arg11 As Variant, Optional Arg12 As Variant, Optional Arg13 As Variant, Optional Arg14 As Variant, Optional Arg15 As Variant, Optional Arg16 As Variant, Optional Arg17 As Variant, Optional Arg18 As Variant, Optional Arg19 As Variant, Optional Arg20 As Variant, Optional Arg21 As Variant, Optional Arg22 As Variant, Optional Arg23 As Variant, Optional Arg24 As Variant, Optional Arg25 As Variant, Optional Arg26 As Variant, Optional Arg27 As Variant, Optional Arg28 As Variant, Optional Arg29 As Variant, Optional Arg30 As Variant) As Variant",
    ),
    (
        "Var_P",
        "Function Var_P(Arg1 As Variant, Optional Arg2 As Variant, Optional Arg3 As Variant, Optional Arg4 As Variant, Optional Arg5 As Variant, Optional Arg6 As Variant, Optional Arg7 As Variant, Optional Arg8 As Variant, Optional Arg9 As Variant, Optional Arg10 As Variant, Optional Arg11 As Variant, Optional Arg12 As Variant, Optional Arg13 As Variant, Optional Arg14 As Variant, Optional Arg15 As Variant, Optional Arg16 As Variant, Optional Arg17 As Variant, Optional Arg18 As Variant, Optional Arg19 As Variant, Optional Arg20 As Variant, Optional Arg21 As Variant, Optional Arg22 As Variant, Optional Arg23 As Variant, Optional Arg24 As Variant, Optional Arg25 As Variant, Optional Arg26 As Variant, Optional Arg27 As Variant, Optional Arg28 As Variant, Optional Arg29 As Variant, Optional Arg30 As Variant) As Variant",
    ),
    (
        "Var_S",
        "Function Var_S(Arg1 As Variant, Optional Arg2 As Variant, Optional Arg3 As Variant, Optional Arg4 As Variant, Optional Arg5 As Variant, Optional Arg6 As Variant, Optional Arg7 As Variant, Optional Arg8 As Variant, Optional Arg9 As Variant, Optional Arg10 As Variant, Optional Arg11 As Variant, Optional Arg12 As Variant, Optional Arg13 As Variant, Optional Arg14 As Variant, Optional Arg15 As Variant, Optional Arg16 As Variant, Optional Arg17 As Variant, Optional Arg18 As Variant, Optional Arg19 As Variant, Optional Arg20 As Variant, Optional Arg21 As Variant, Optional Arg22 As Variant, Optional Arg23 As Variant, Optional Arg24 As Variant, Optional Arg25 As Variant, Optional Arg26 As Variant, Optional Arg27 As Variant, Optional Arg28 As Variant, Optional Arg29 As Variant, Optional Arg30 As Variant) As Variant",
    ),
    (
        "VarP",
        "Function VarP(Arg1 As Variant, Optional Arg2 As Variant, Optional Arg3 As Variant, Optional Arg4 As Variant, Optional Arg5 As Variant, Optional Arg6 As Variant, Optional Arg7 As Variant, Optional Arg8 As Variant, Optional Arg9 As Variant, Optional Arg10 As Variant, Optional Arg11 As Variant, Optional Arg12 As Variant, Optional Arg13 As Variant, Optional Arg14 As Variant, Optional Arg15 As Variant, Optional Arg16 As Variant, Optional Arg17 As Variant, Optional Arg18 As Variant, Optional Arg19 As Variant, Optional Arg20 As Variant, Optional Arg21 As Variant, Optional Arg22 As Variant, Optional Arg23 As Variant, Optional Arg24 As Variant, Optional Arg25 As Variant, Optional Arg26 As Variant, Optional Arg27 As Variant, Optional Arg28 As Variant, Optional Arg29 As Variant, Optional Arg30 As Variant) As Variant",
    ),
    (
        "Vdb",
        "Function Vdb(Arg1 As Variant, Arg2 As Variant, Arg3 As Variant, Arg4 As Variant, Arg5 As Variant, Optional Arg6 As Variant, Optional Arg7 As Variant) As Variant",
    ),
    ("VLookup", "Function VLookup(Arg1 As Variant, Arg2 As Variant, Arg3 As Variant, Optional Arg4 As Variant) As Variant"),
    ("Weekday", "Function Weekday(Arg1 As Variant, Optional Arg2 As Variant) As Variant"),
    ("WeekNum", "Function WeekNum(Arg1 As Variant, Optional Arg2 As Variant) As Variant"),
    ("Weibull", "Function Weibull(Arg1 As Variant, Arg2 As Variant, Arg3 As Variant, Arg4 As Variant) As Variant"),
    ("Weibull_Dist", "Function Weibull_Dist(Arg1 As Variant, Arg2 As Variant, Arg3 As Variant, Arg4 As Variant) As Variant"),
    ("WorkDay", "Function WorkDay(Arg1 As Variant, Arg2 As Variant, Optional Arg3 As Variant) As Variant"),
    ("WorkDay_Intl", "Function WorkDay_Intl(Arg1 As Variant, Arg2 As Variant, Optional Arg3 As Variant, Optional Arg4 As Variant) As Variant"),
    ("Xirr", "Function Xirr(Arg1 As Variant, Arg2 As Variant, Optional Arg3 As Variant) As Variant"),
    ("Xnpv", "Function Xnpv(Arg1 As Variant, Arg2 As Variant) As Variant"),
    ("YearFrac", "Function YearFrac(Arg1 As Variant, Arg2 As Variant, Optional Arg3 As Variant) As Variant"),
    ("YieldDisc", "Function YieldDisc(Arg1 As Variant, Arg2 As Variant, Arg3 As Variant, Arg4 As Variant, Optional Arg5 As Variant) As Variant"),
    (
        "YieldMat",
        "Function YieldMat(Arg1 As Variant, Arg2 As Variant, Arg3 As Variant, Arg4 As Variant, Arg5 As Variant, Optional Arg6 As Variant) As Variant",
    ),
    ("Z_Test", "Function Z_Test(Arg1 As Variant, Arg2 As Variant, Optional Arg3 As Variant) As Variant"),
    ("ZTest", "Function ZTest(Arg1 As Variant, Arg2 As Variant, Optional Arg3 As Variant) As Variant"),
];
pub(crate) fn descriptors() -> Vec<super::MemberDescriptor> {
    let mut out = vec![
        descriptor("WorksheetFunction", "Parent", "Property Get Parent() As Application"),
        descriptor("WorksheetFunction", "Application", "Property Get Application() As Application"),
        descriptor("Name", "NameLocal", "Property NameLocal As String"),
        descriptor("Name", "RefersToLocal", "Property RefersToLocal As String"),
        descriptor("Name", "RefersToR1C1Local", "Property RefersToR1C1Local As String"),
        descriptor("Name", "Value", "Property Value As String"),
        descriptor("Range", "FormulaArray", "Property FormulaArray As Variant"),
        descriptor("Application", "WorksheetFunction", "Property Get WorksheetFunction() As WorksheetFunction"),
        descriptor("Application", "Evaluate", "Function Evaluate(Name As Variant) As Variant"),
        descriptor("Worksheet", "Evaluate", "Function Evaluate(Name As Variant) As Variant"),
        descriptor("Range", "Evaluate", "Function Evaluate(Name As Variant) As Variant"),
        descriptor("Application", "Names", "Property Get Names() As Names"),
        descriptor("Workbook", "Names", "Property Get Names() As Names"),
        descriptor("Worksheet", "Names", "Property Get Names() As Names"),
        descriptor("Names", "Count", "Property Get Count() As Long"),
        descriptor(
            "Names",
            "Item",
            "Property Get Item(Optional Index As Variant, Optional IndexLocal As Variant, Optional RefersTo As Variant) As Name",
        ),
        descriptor(
            "Names",
            "Add",
            "Function Add(Optional Name As Variant, Optional RefersTo As Variant, Optional Visible As Variant, Optional MacroType As Variant, Optional ShortcutKey As Variant, Optional Category As Variant, Optional NameLocal As Variant, Optional RefersToLocal As Variant, Optional CategoryLocal As Variant, Optional RefersToR1C1 As Variant, Optional RefersToR1C1Local As Variant) As Name",
        ),
        descriptor("Names", "Parent", "Property Get Parent() As Object"),
        descriptor("Names", "Application", "Property Get Application() As Application"),
        descriptor("Name", "Parent", "Property Get Parent() As Object"),
        descriptor("Name", "Application", "Property Get Application() As Application"),
        descriptor("Name", "Name", "Property Name As String"),
        descriptor("Name", "RefersTo", "Property RefersTo As String"),
        descriptor("Name", "Visible", "Property Visible As Boolean"),
        descriptor("Name", "Comment", "Property Comment As String"),
        descriptor("Name", "RefersToRange", "Property Get RefersToRange() As Range"),
        descriptor("Name", "RefersToR1C1", "Property RefersToR1C1 As String"),
        descriptor("Name", "Delete", "Sub Delete()"),
        descriptor("Range", "Formula2", "Property Formula2 As Variant"),
        descriptor("Range", "FormulaR1C1", "Property FormulaR1C1 As Variant"),
        descriptor("Range", "Formula2R1C1", "Property Formula2R1C1 As Variant"),
        descriptor("Range", "FormulaLocal", "Property FormulaLocal As Variant"),
        descriptor("Range", "Formula2Local", "Property Formula2Local As Variant"),
        descriptor("Range", "FormulaR1C1Local", "Property FormulaR1C1Local As Variant"),
        descriptor("Range", "Formula2R1C1Local", "Property Formula2R1C1Local As Variant"),
        descriptor("Range", "HasFormula", "Property Get HasFormula() As Variant"),
    ];
    out.extend(FUNCTION_MEMBERS.iter().map(|(name, signature)| {
        let mut member = descriptor("WorksheetFunction", name, signature);
        if function_name(name).is_none() {
            member.availability = super::MemberAvailability::OutOfScope;
        }
        member
    }));
    out
}
fn descriptor(owner: &'static str, name: &'static str, signature: &'static str) -> super::MemberDescriptor {
    super::MemberDescriptor {
        owner,
        name,
        signature,
        availability: super::MemberAvailability::Implemented,
        implementation_symbol: "crates/vba/src/formula_api.rs::Bridge::formula_api",
    }
}
