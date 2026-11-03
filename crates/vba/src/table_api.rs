//! OM18: table graph and lifecycle over native table IDs, identities and commands.
use super::features::{Feature, absent, reads};
use super::objects::{arguments, numeric_index, required, string};
use super::{Bridge, HostCall, HostResult, InvokeKind, Object, Session, Value, failure, values};
use gridcraft_engine::{
    StructuralEdit,
    core::{CellRef, RangeRef},
    model::{Table, TotalsFn},
};
use serde_json::json;
impl Bridge {
    pub(crate) fn table_node(sheet: usize, table: &Table) -> Feature {
        let mut f = Feature::new("ListObject", Some(sheet));
        f.key = u64::from(table.id);
        f.identity = Some(table.identity.clone());
        f
    }
    pub(crate) fn table_api(&mut self, session: &mut Session, object: &Object, call: &HostCall) -> Option<HostResult<Value>> {
        if let Object::Sheet(sheet) = object
            && call.member.eq_ignore_ascii_case("ListObjects")
            && reads(call)
        {
            return Some({
                let f = Feature::new("ListObjects", Some(*sheet));
                if call.arguments.is_empty() { self.feature(f) } else { self.table_dispatch(session, &f, call) }
            });
        }
        if let Object::Range(sheet, range) = object
            && call.member.eq_ignore_ascii_case("ListObject")
            && reads(call)
        {
            return Some((|| {
                arguments(call, &[])?;
                let sh = session.doc().map_err(|e| failure(91, e.to_string()))?.wb.sheet(*sheet).ok_or_else(|| failure(91, "worksheet deleted"))?;
                let rect = range.rectangle()?;
                let table =
                    sh.tables.iter().find(|t| t.range.contains_range(&rect)).ok_or_else(|| failure(1004, "range does not belong to a table"))?;
                self.feature(Self::table_node(*sheet, table))
            })());
        }
        let Object::Feature(f) = object else {
            return None;
        };
        if !matches!(f.kind, "ListObjects" | "ListObject" | "ListRows" | "ListRow" | "ListColumns" | "ListColumn") {
            return None;
        }
        Some(self.table_dispatch(session, f, call))
    }
    fn table_dispatch(&mut self, session: &mut Session, f: &Feature, call: &HostCall) -> HostResult<Value> {
        let sheet = f.sheet.ok_or_else(|| failure(91, "table owner deleted"))?;
        let doc = session.doc().map_err(|e| failure(91, e.to_string()))?;
        let sh = doc.wb.sheet(sheet).ok_or_else(|| failure(91, "worksheet deleted"))?;
        let member = call.member.to_ascii_lowercase();
        if member == "application" && reads(call) {
            arguments(call, &[])?;
            return self.intern(Object::Application);
        }
        if f.kind == "ListObjects" {
            if call.kind == InvokeKind::Enumerate {
                arguments(call, &[])?;
                return self.object_array(sh.tables.iter().map(|t| Self::table_node(sheet, t).object()).collect());
            }
            match member.as_str() {
                "count" if reads(call) => {
                    arguments(call, &[])?;
                    return Ok(Value::Long(sh.tables.len() as i32));
                }
                "parent" if reads(call) => {
                    arguments(call, &[])?;
                    return self.intern(Object::Sheet(sheet));
                }
                "item" | "" | "listobjects" if reads(call) => {
                    let args = arguments(call, &["Index"])?;
                    let value = required(&args, 0)?;
                    let table = if matches!(value, Value::String(_)) {
                        let name = string(value)?;
                        sh.tables.iter().find(|t| t.name.eq_ignore_ascii_case(&name))
                    } else {
                        usize::try_from(numeric_index(value)?.checked_sub(1).ok_or_else(|| failure(9, "table index"))?)
                            .ok()
                            .and_then(|i| sh.tables.get(i))
                    }
                    .ok_or_else(|| failure(9, "table not found"))?;
                    return self.feature(Self::table_node(sheet, table));
                }
                "add" => return self.add_table(session, sheet, call),
                _ => return Err(failure(438, "unsupported ListObjects member")),
            }
        }
        let table = sh.tables.iter().find(|t| u64::from(t.id) == f.key).ok_or_else(|| failure(91, "table was deleted"))?.clone();
        if f.kind != "ListColumn" && f.identity.as_ref().is_some_and(|id| !id.same(&table.identity)) {
            return Err(failure(91, "table identity is stale"));
        }
        if matches!(f.kind, "ListRows" | "ListColumns") {
            let rows = f.kind == "ListRows";
            let count = if rows { table.data_range().map(|r| r.height() as usize).unwrap_or(0) } else { table.columns.len() };
            if call.kind == InvokeKind::Enumerate {
                arguments(call, &[])?;
                if count > super::MAX_OBJECTS {
                    return Err(failure(7, "table enumeration budget"));
                }
                return self.object_array((0..count).map(|i| Self::table_child(sheet, &table, rows, i).object()).collect());
            }
            if member == "count" && reads(call) {
                arguments(call, &[])?;
                return Ok(Value::Long(count as i32));
            }
            if member == "parent" && reads(call) {
                arguments(call, &[])?;
                return self.feature(Self::table_node(sheet, &table));
            }
            if member == "add" {
                let args = arguments(call, if rows { &["Position", "AlwaysInsert"] } else { &["Position"] })?;
                let position = if absent(&args[0]) {
                    count + 1
                } else {
                    usize::try_from(numeric_index(&args[0])?).map_err(|_| failure(1004, "invalid table position"))?
                };
                if position == 0 || position > count + 1 {
                    return Err(failure(1004, "invalid table position"));
                }
                let always = if rows && !absent(&args[1]) { super::address::boolean(&args[1])? } else { true };
                self.objects.reserve(1)?;
                self.table_mutate(session, sheet, &table, if rows { "insertRow" } else { "insertColumn" }, position, always)?;
                let table = session
                    .doc()
                    .map_err(|e| failure(91, e.to_string()))?
                    .wb
                    .sheet(sheet)
                    .and_then(|sh| sh.tables.iter().find(|t| t.id == table.id))
                    .ok_or_else(|| failure(91, "table disappeared"))?;
                return self.feature(Self::table_child(sheet, table, rows, position - 1));
            }
            if matches!(member.as_str(), "item" | "") && reads(call) {
                let args = arguments(call, &["Index"])?;
                let value = required(&args, 0)?;
                let index = if !rows && matches!(value, Value::String(_)) {
                    let name = string(value)?;
                    table.column_index(&name).ok_or_else(|| failure(9, "table column not found"))?
                } else {
                    usize::try_from(numeric_index(value)?.checked_sub(1).ok_or_else(|| failure(9, "table index"))?)
                        .map_err(|_| failure(9, "table index"))?
                };
                if index >= count {
                    return Err(failure(9, "table index"));
                }
                return self.feature(Self::table_child(sheet, &table, rows, index));
            }
            return Err(failure(438, "unsupported table collection member"));
        }
        if matches!(f.kind, "ListRow" | "ListColumn") {
            let row = f.kind == "ListRow";
            let index = if row {
                let at = f.range.as_ref().ok_or_else(|| failure(91, "row deleted"))?.first()?.start.row;
                let data = table.data_range().filter(|r| (r.start.row..=r.end.row).contains(&at)).ok_or_else(|| failure(91, "table row deleted"))?;
                (at - data.start.row) as usize
            } else {
                table
                    .columns
                    .iter()
                    .position(|c| f.identity.as_ref().is_some_and(|id| id.same(&c.identity)))
                    .ok_or_else(|| failure(91, "table column deleted"))?
            };
            if member == "parent" && reads(call) {
                arguments(call, &[])?;
                return self.feature(Self::table_node(sheet, &table));
            }
            if member == "index" && reads(call) {
                arguments(call, &[])?;
                return Ok(Value::Long(index as i32 + 1));
            }
            if member == "range" && reads(call) {
                arguments(call, &[])?;
                let rect = if row {
                    let data = table.data_range().ok_or_else(|| failure(91, "row deleted"))?;
                    RangeRef::new(
                        CellRef::new(data.start.row + index as u32, table.range.start.col),
                        CellRef::new(data.start.row + index as u32, table.range.end.col),
                    )
                } else {
                    RangeRef::new(
                        CellRef::new(table.range.start.row, table.range.start.col + index as u32),
                        CellRef::new(table.range.end.row, table.range.start.col + index as u32),
                    )
                };
                return self.intern(Object::Range(sheet, rect.into()));
            }
            if member == "delete" {
                arguments(call, &[])?;
                return self.table_mutate(session, sheet, &table, if row { "deleteRow" } else { "deleteColumn" }, index + 1, true);
            }
            if !row {
                let column = table.columns.get(index).ok_or_else(|| failure(91, "column deleted"))?;
                if reads(call) {
                    arguments(call, &[])?;
                    return match member.as_str() {
                        "name" => values::vba_text(&column.name),
                        "databodyrange" | "total" => {
                            let area = if member == "databodyrange" { table.data_range() } else { table.totals_range() };
                            match area {
                                Some(r) => self.intern(Object::Range(
                                    sheet,
                                    RangeRef::new(
                                        CellRef::new(r.start.row, table.range.start.col + index as u32),
                                        CellRef::new(r.end.row, table.range.start.col + index as u32),
                                    )
                                    .into(),
                                )),
                                None => Ok(Value::Nothing),
                            }
                        }
                        "totalscalculation" => Ok(Value::Long(total_number(column.totals))),
                        _ => Err(failure(438, "unsupported ListColumn member")),
                    };
                }
                if call.kind == InvokeKind::Let {
                    let args = arguments(call, &["Value"])?;
                    let value = required(&args, 0)?;
                    if sh.is_protected() {
                        return Err(failure(1004, "table changes require unprotected worksheet"));
                    }
                    match member.as_str() {
                        "name" => {
                            let name = string(value)?;
                            if name.is_empty() || name.chars().count() > 255 {
                                return Err(failure(1004, "invalid column name"));
                            }
                            if table.columns.iter().enumerate().any(|(i, c)| i != index && c.name.eq_ignore_ascii_case(&name)) {
                                return Err(failure(1004, "table column name collision"));
                            }
                            return self.write_range(
                                session,
                                sheet,
                                &RangeRef::cell(CellRef::new(table.range.start.row, table.range.start.col + index as u32)).into(),
                                &values::vba_text(&name)?,
                                false,
                                false,
                            );
                        }
                        "totalscalculation" => {
                            let function = total_function(numeric_index(value)?)?;
                            let event = if let Some(r) = table.totals_range() {
                                self.prepare_change(session, sheet, r.into())?
                            } else {
                                super::PreparedChanges::default()
                            };
                            session
                                .execute("table.totalFunction", json!({"table":table.name,"column":column.name,"function":function}))
                                .map_err(|e| failure(1004, e.to_string()))?;
                            self.queue_changes(event);
                            return Ok(Value::Empty);
                        }
                        _ => {}
                    }
                }
            }
            return Err(failure(438, "unsupported table row/column member"));
        }
        if reads(call) {
            match member.as_str() {
                "name" | "displayname" => {
                    arguments(call, &[])?;
                    return values::vba_text(&table.name);
                }
                "parent" => {
                    arguments(call, &[])?;
                    return self.intern(Object::Sheet(sheet));
                }
                "listrows" | "listcolumns" => {
                    let mut node = Self::table_node(sheet, &table);
                    node.kind = if member == "listrows" { "ListRows" } else { "ListColumns" };
                    if call.arguments.is_empty() {
                        return self.feature(node);
                    }
                    let mut call = call.clone();
                    call.member = "Item".into();
                    return self.table_dispatch(session, &node, &call);
                }
                "range" | "databodyrange" | "headerrowrange" | "totalsrowrange" => {
                    arguments(call, &[])?;
                    return match match member.as_str() {
                        "databodyrange" => table.data_range(),
                        "headerrowrange" => table.header_range(),
                        "totalsrowrange" => table.totals_range(),
                        _ => Some(table.range),
                    } {
                        Some(r) => self.intern(Object::Range(sheet, r.into())),
                        None => Ok(Value::Nothing),
                    };
                }
                "tableStyle" | "tablestyle" => {
                    arguments(call, &[])?;
                    return values::vba_text(&table.style);
                }
                "showtotals"
                | "showheaders"
                | "showautofilter"
                | "showtablestylerowstripes"
                | "showtablestylecolumnstripes"
                | "showtablestylefirstcolumn"
                | "showtablestylelastcolumn" => {
                    arguments(call, &[])?;
                    return Ok(Value::Boolean(match member.as_str() {
                        "showtotals" => table.totals_row,
                        "showheaders" => table.header_row,
                        "showautofilter" => table.filter_button,
                        "showtablestylerowstripes" => table.banded_rows,
                        "showtablestylecolumnstripes" => table.banded_cols,
                        "showtablestylefirstcolumn" => table.first_col,
                        _ => table.last_col,
                    }));
                }
                "autofilter" => {
                    arguments(call, &[])?;
                    let mut node = Self::table_node(sheet, &table);
                    node.kind = "AutoFilter";
                    return self.feature(node);
                }
                "sort" => {
                    arguments(call, &[])?;
                    self.sort_state(session, sheet, u64::from(table.id))?;
                    let mut node = Self::table_node(sheet, &table);
                    node.kind = "Sort";
                    return self.feature(node);
                }
                "sourcetype" => {
                    arguments(call, &[])?;
                    return Ok(Value::Long(if table.connection_id.is_some() { 0 } else { 1 }));
                }
                "resize" => {
                    let args = arguments(call, &["Range"])?;
                    let reference = self.range_argument(session, sheet, required(&args, 0)?)?;
                    let Object::Range(_, range) = self.object(session, reference)? else {
                        return Err(failure(13, "Range required"));
                    };
                    let rect = range.rectangle()?;
                    values::dimensions(rect)?;
                    if sh.is_protected() {
                        return Err(failure(1004, "table changes require unprotected worksheet"));
                    }
                    session.execute("table.resize", json!({"table":table.name,"range":rect.a1()})).map_err(|e| failure(1004, e.to_string()))?;
                    return Ok(Value::Empty);
                }
                "unlist" => {
                    arguments(call, &[])?;
                    if sh.is_protected() {
                        return Err(failure(1004, "table changes require unprotected worksheet"));
                    }
                    session.execute("table.convertToRange", json!({"table":table.name})).map_err(|e| failure(1004, e.to_string()))?;
                    return Ok(Value::Empty);
                }
                "delete" => {
                    arguments(call, &[])?;
                    values::dimensions(table.range)?;
                    let changes = self.prepare_change(session, sheet, table.range.into())?;
                    session.execute("table.delete", json!({"table":table.name})).map_err(|e| failure(1004, e.to_string()))?;
                    self.queue_changes(changes);
                    return Ok(Value::Empty);
                }
                _ => {}
            }
        }
        if call.kind != InvokeKind::Let {
            return Err(failure(438, "unsupported ListObject member"));
        }
        if sh.is_protected() {
            return Err(failure(1004, "table changes require unprotected worksheet"));
        }
        let args = arguments(call, &["Value"])?;
        let value = required(&args, 0)?;
        let mut params = json!({"table":table.name});
        let command = match member.as_str() {
            "name" | "displayname" => {
                params["name"] = json!(string(value)?);
                "table.rename"
            }
            "tablestyle" => {
                let name = string(value)?;
                if !gridcraft_engine::tables::gallery().contains(&name) {
                    return Err(failure(1004, "unknown native table style"));
                }
                params["style"] = json!(name);
                "table.style"
            }
            "showtotals"
            | "showautofilter"
            | "showtablestylerowstripes"
            | "showtablestylecolumnstripes"
            | "showtablestylefirstcolumn"
            | "showtablestylelastcolumn" => {
                params["on"] = json!(super::address::boolean(value)?);
                match member.as_str() {
                    "showtotals" => "table.totalRow",
                    "showautofilter" => "table.headerRow",
                    "showtablestylerowstripes" => "table.bandedRows",
                    "showtablestylecolumnstripes" => "table.bandedColumns",
                    "showtablestylefirstcolumn" => "table.firstColumn",
                    _ => "table.lastColumn",
                }
            }
            _ => return Err(failure(438, "unsupported ListObject assignment")),
        };
        if member == "showtotals" && super::address::boolean(value)? == table.totals_row {
            return Ok(Value::Empty);
        }
        let changes = if member == "showtotals" {
            let total = if table.totals_row {
                table.totals_range().ok_or_else(|| failure(1004, "total range unavailable"))?
            } else {
                let row = table
                    .range
                    .end
                    .row
                    .checked_add(1)
                    .filter(|r| *r < gridcraft_engine::core::MAX_ROWS)
                    .ok_or_else(|| failure(1004, "worksheet boundary"))?;
                let range = RangeRef::new(CellRef::new(row, table.range.start.col), CellRef::new(row, table.range.end.col));
                if range.iter().any(|at| sh.cell(at).is_some_and(|c| c.formula.is_some() || !c.value.is_empty())) {
                    return Err(failure(1004, "total row would overwrite data; insert a row first"));
                }
                range
            };
            self.prepare_change(session, sheet, total.into())?
        } else {
            super::PreparedChanges::default()
        };
        session.execute(command, params).map_err(|e| failure(1004, e.to_string()))?;
        self.queue_changes(changes);
        Ok(Value::Empty)
    }
    fn table_child(sheet: usize, table: &Table, row: bool, index: usize) -> Feature {
        let mut f = Self::table_node(sheet, table);
        f.kind = if row { "ListRow" } else { "ListColumn" };
        f.part = index as i32;
        if row {
            if let Some(data) = table.data_range() {
                f.range = Some(
                    RangeRef::new(
                        CellRef::new(data.start.row + index as u32, table.range.start.col),
                        CellRef::new(data.start.row + index as u32, table.range.end.col),
                    )
                    .into(),
                );
            }
        } else {
            f.identity = table.columns.get(index).map(|c| c.identity.clone());
        }
        f
    }
    fn add_table(&mut self, session: &mut Session, sheet: usize, call: &HostCall) -> HostResult<Value> {
        let args = arguments(call, &["SourceType", "Source", "LinkSource", "XlListObjectHasHeaders", "Destination", "TableStyleName"])?;
        if !absent(&args[0]) && numeric_index(&args[0])? != 1 || !absent(&args[2]) {
            return Err(failure(1004, "only local range-backed tables are supported"));
        }
        let source = if absent(&args[1]) {
            let doc = session.doc().map_err(|e| failure(91, e.to_string()))?;
            let sh = doc.wb.sheet(sheet).ok_or_else(|| failure(91, "worksheet deleted"))?;
            let at = if sheet == doc.sheet_index() { doc.selection.active } else { CellRef::new(0, 0) };
            let rect = gridcraft_engine::selection::current_region(sh, at);
            self.intern(Object::Range(sheet, rect.into()))?
        } else {
            args[1].clone()
        };
        let reference = self.range_argument(session, sheet, &source)?;
        let Object::Range(_, source) = self.object(session, reference)? else {
            return Err(failure(13, "Range required"));
        };
        let rect = source.rectangle()?;
        values::dimensions(rect)?;
        let sh = session.doc().map_err(|e| failure(91, e.to_string()))?.wb.sheet(sheet).ok_or_else(|| failure(91, "worksheet deleted"))?;
        if sh.is_protected() {
            return Err(failure(1004, "table creation requires unprotected worksheet"));
        }
        let headers = if absent(&args[3]) { 0 } else { numeric_index(&args[3])? };
        if !(0..=2).contains(&headers) {
            return Err(failure(1004, "invalid header setting"));
        }
        let header = if headers == 0 {
            rect.height() > 1 && (rect.start.col..=rect.end.col).all(|col| sh.value(CellRef::new(rect.start.row, col)).is_text())
        } else {
            headers == 1
        };
        let mut params = json!({"sheet":sheet,"range":rect.a1(),"header":header});
        if !absent(&args[5]) {
            let style = string(&args[5])?;
            if !gridcraft_engine::tables::gallery().contains(&style) {
                return Err(failure(1004, "unknown native table style"));
            }
            params["style"] = json!(style);
        }
        let impact = if header {
            rect
        } else {
            RangeRef::new(
                rect.start,
                CellRef::new(
                    rect.end
                        .row
                        .checked_add(1)
                        .filter(|r| *r < gridcraft_engine::core::MAX_ROWS)
                        .ok_or_else(|| failure(1004, "worksheet boundary"))?,
                    rect.end.col,
                ),
            )
        };
        values::dimensions(impact)?;
        let changes = self.prepare_change(session, sheet, impact.into())?;
        self.objects.reserve(1)?;
        let rebased = if !header {
            Some(
                self.objects.rebase(
                    self.context,
                    sheet,
                    &StructuralEdit::InsertDown
                        .reference_edit(RangeRef::new(rect.start, CellRef::new(rect.start.row, rect.end.col)))
                        .map_err(|e| failure(1004, e.to_string()))?,
                )?,
            )
        } else {
            None
        };
        let result = session.execute("insert.table", params).map_err(|e| failure(1004, e.to_string()))?;
        if let Some(rebased) = rebased {
            self.objects.replace(rebased)?;
        }
        self.range_epochs.insert(self.context, session.doc().map_err(|e| failure(91, e.to_string()))?.range_epoch);
        self.queue_changes(changes);
        let name = result.get("table").and_then(serde_json::Value::as_str).ok_or_else(|| failure(1004, "table result missing"))?;
        let table = session
            .doc()
            .map_err(|e| failure(91, e.to_string()))?
            .wb
            .sheet(sheet)
            .and_then(|sh| sh.tables.iter().find(|t| t.name == name))
            .ok_or_else(|| failure(91, "table disappeared"))?;
        self.feature(Self::table_node(sheet, table))
    }
    fn table_mutate(
        &mut self,
        session: &mut Session,
        sheet: usize,
        table: &Table,
        operation: &str,
        position: usize,
        always: bool,
    ) -> HostResult<Value> {
        let row = operation.ends_with("Row");
        let insert = operation.starts_with("insert");
        let count = if row { table.data_range().map(|r| r.height()).unwrap_or(0) } else { table.range.width() };
        let coordinate =
            if row { table.range.start.row + u32::from(table.header_row) + position as u32 - 1 } else { table.range.start.col + position as u32 - 1 };
        let rect = if row {
            RangeRef::new(CellRef::new(coordinate, table.range.start.col), CellRef::new(coordinate, table.range.end.col))
        } else {
            RangeRef::new(CellRef::new(table.range.start.row, coordinate), CellRef::new(table.range.end.row, coordinate))
        };
        let shift = !(row
            && insert
            && !always
            && position as u32 == count + 1
            && !table.totals_row
            && rect.iter().all(|at| {
                session
                    .doc()
                    .ok()
                    .and_then(|d| d.wb.sheet(sheet))
                    .and_then(|sh| sh.cell(at))
                    .is_none_or(|c| c.formula.is_none() && c.value.is_empty())
            }));
        let native = match (row, insert) {
            (true, true) => StructuralEdit::InsertDown,
            (true, false) => StructuralEdit::DeleteUp,
            (false, true) => StructuralEdit::InsertRight,
            (false, false) => StructuralEdit::DeleteLeft,
        };
        let rebased = if shift {
            Some(self.objects.rebase(self.context, sheet, &native.reference_edit(rect).map_err(|e| failure(1004, e.to_string()))?)?)
        } else {
            None
        };
        let changes = self.prepare_change(session, sheet, rect.into())?;
        session
            .execute(&format!("table.{operation}"), json!({"table":table.name,"position":position,"alwaysInsert":always}))
            .map_err(|e| failure(1004, e.to_string()))?;
        if let Some(rebased) = rebased {
            self.objects.replace(rebased)?;
        }
        self.range_epochs.insert(self.context, session.doc().map_err(|e| failure(91, e.to_string()))?.range_epoch);
        self.queue_changes(changes);
        Ok(Value::Empty)
    }
}
fn total_number(t: TotalsFn) -> i32 {
    match t {
        TotalsFn::None => 0,
        TotalsFn::Sum => 1,
        TotalsFn::Average => 2,
        TotalsFn::Count => 3,
        TotalsFn::CountNums => 4,
        TotalsFn::Min => 5,
        TotalsFn::Max => 6,
        TotalsFn::StdDev => 7,
        TotalsFn::Var => 8,
        TotalsFn::Custom => 9,
    }
}
fn total_function(n: i32) -> HostResult<&'static str> {
    match n {
        0 => Ok("none"),
        1 => Ok("sum"),
        2 => Ok("average"),
        3 => Ok("count"),
        4 => Ok("countNums"),
        5 => Ok("min"),
        6 => Ok("max"),
        7 => Ok("stdDev"),
        8 => Ok("var"),
        _ => Err(failure(1004, "unsupported totals function")),
    }
}

pub(crate) fn descriptors() -> Vec<super::MemberDescriptor> {
    vec![
        super::MemberDescriptor {
            owner: "ListObject",
            name: "Sort",
            signature: "Property Get Sort() As Sort",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/table_api.rs::Bridge::table_api",
        },
        super::MemberDescriptor {
            owner: "Worksheet",
            name: "ListObjects",
            signature: "Property Get ListObjects() As ListObjects",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/table_api.rs::Bridge::table_api",
        },
        super::MemberDescriptor {
            owner: "Range",
            name: "ListObject",
            signature: "Property Get ListObject() As ListObject",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/table_api.rs::Bridge::table_api",
        },
        super::MemberDescriptor {
            owner: "ListObjects",
            name: "Parent",
            signature: "Property Get Parent() As Object",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/table_api.rs::Bridge::table_api",
        },
        super::MemberDescriptor {
            owner: "ListObjects",
            name: "Application",
            signature: "Property Get Application() As Application",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/table_api.rs::Bridge::table_api",
        },
        super::MemberDescriptor {
            owner: "ListObject",
            name: "Parent",
            signature: "Property Get Parent() As Object",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/table_api.rs::Bridge::table_api",
        },
        super::MemberDescriptor {
            owner: "ListObject",
            name: "Application",
            signature: "Property Get Application() As Application",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/table_api.rs::Bridge::table_api",
        },
        super::MemberDescriptor {
            owner: "ListRows",
            name: "Parent",
            signature: "Property Get Parent() As Object",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/table_api.rs::Bridge::table_api",
        },
        super::MemberDescriptor {
            owner: "ListRows",
            name: "Application",
            signature: "Property Get Application() As Application",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/table_api.rs::Bridge::table_api",
        },
        super::MemberDescriptor {
            owner: "ListRow",
            name: "Parent",
            signature: "Property Get Parent() As Object",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/table_api.rs::Bridge::table_api",
        },
        super::MemberDescriptor {
            owner: "ListRow",
            name: "Application",
            signature: "Property Get Application() As Application",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/table_api.rs::Bridge::table_api",
        },
        super::MemberDescriptor {
            owner: "ListColumns",
            name: "Parent",
            signature: "Property Get Parent() As Object",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/table_api.rs::Bridge::table_api",
        },
        super::MemberDescriptor {
            owner: "ListColumns",
            name: "Application",
            signature: "Property Get Application() As Application",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/table_api.rs::Bridge::table_api",
        },
        super::MemberDescriptor {
            owner: "ListColumn",
            name: "Parent",
            signature: "Property Get Parent() As Object",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/table_api.rs::Bridge::table_api",
        },
        super::MemberDescriptor {
            owner: "ListColumn",
            name: "Application",
            signature: "Property Get Application() As Application",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/table_api.rs::Bridge::table_api",
        },
        super::MemberDescriptor {
            owner: "ListObjects",
            name: "Count",
            signature: "Property Get Count() As Long",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/table_api.rs::Bridge::table_api",
        },
        super::MemberDescriptor {
            owner: "ListObjects",
            name: "Item",
            signature: "Property Get Item(Index As Variant) As ListObject",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/table_api.rs::Bridge::table_api",
        },
        super::MemberDescriptor {
            owner: "ListRows",
            name: "Count",
            signature: "Property Get Count() As Long",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/table_api.rs::Bridge::table_api",
        },
        super::MemberDescriptor {
            owner: "ListRows",
            name: "Item",
            signature: "Property Get Item(Index As Variant) As ListRow",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/table_api.rs::Bridge::table_api",
        },
        super::MemberDescriptor {
            owner: "ListColumns",
            name: "Count",
            signature: "Property Get Count() As Long",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/table_api.rs::Bridge::table_api",
        },
        super::MemberDescriptor {
            owner: "ListColumns",
            name: "Item",
            signature: "Property Get Item(Index As Variant) As ListColumn",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/table_api.rs::Bridge::table_api",
        },
        super::MemberDescriptor {
            owner: "ListObjects",
            name: "Add",
            signature: "Function Add(Optional SourceType As Variant, Optional Source As Variant, Optional LinkSource As Variant, Optional XlListObjectHasHeaders As Variant, Optional Destination As Variant, Optional TableStyleName As Variant) As ListObject",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/table_api.rs::Bridge::table_api",
        },
        super::MemberDescriptor {
            owner: "ListRows",
            name: "Add",
            signature: "Function Add(Optional Position As Variant, Optional AlwaysInsert As Variant) As ListRow",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/table_api.rs::Bridge::table_api",
        },
        super::MemberDescriptor {
            owner: "ListColumns",
            name: "Add",
            signature: "Function Add(Optional Position As Variant) As ListColumn",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/table_api.rs::Bridge::table_api",
        },
        super::MemberDescriptor {
            owner: "ListRow",
            name: "Range",
            signature: "Property Get Range() As Range",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/table_api.rs::Bridge::table_api",
        },
        super::MemberDescriptor {
            owner: "ListRow",
            name: "Index",
            signature: "Property Get Index() As Long",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/table_api.rs::Bridge::table_api",
        },
        super::MemberDescriptor {
            owner: "ListRow",
            name: "Delete",
            signature: "Sub Delete()",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/table_api.rs::Bridge::table_api",
        },
        super::MemberDescriptor {
            owner: "ListColumn",
            name: "Range",
            signature: "Property Get Range() As Range",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/table_api.rs::Bridge::table_api",
        },
        super::MemberDescriptor {
            owner: "ListColumn",
            name: "Index",
            signature: "Property Get Index() As Long",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/table_api.rs::Bridge::table_api",
        },
        super::MemberDescriptor {
            owner: "ListColumn",
            name: "Delete",
            signature: "Sub Delete()",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/table_api.rs::Bridge::table_api",
        },
        super::MemberDescriptor {
            owner: "ListColumn",
            name: "Name",
            signature: "Property Name As String",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/table_api.rs::Bridge::table_api",
        },
        super::MemberDescriptor {
            owner: "ListColumn",
            name: "DataBodyRange",
            signature: "Property Get DataBodyRange() As Range",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/table_api.rs::Bridge::table_api",
        },
        super::MemberDescriptor {
            owner: "ListColumn",
            name: "Total",
            signature: "Property Get Total() As Range",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/table_api.rs::Bridge::table_api",
        },
        super::MemberDescriptor {
            owner: "ListColumn",
            name: "TotalsCalculation",
            signature: "Property TotalsCalculation As Long",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/table_api.rs::Bridge::table_api",
        },
        super::MemberDescriptor {
            owner: "ListObject",
            name: "Range",
            signature: "Property Get Range() As Range",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/table_api.rs::Bridge::table_api",
        },
        super::MemberDescriptor {
            owner: "ListObject",
            name: "DataBodyRange",
            signature: "Property Get DataBodyRange() As Range",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/table_api.rs::Bridge::table_api",
        },
        super::MemberDescriptor {
            owner: "ListObject",
            name: "HeaderRowRange",
            signature: "Property Get HeaderRowRange() As Range",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/table_api.rs::Bridge::table_api",
        },
        super::MemberDescriptor {
            owner: "ListObject",
            name: "TotalsRowRange",
            signature: "Property Get TotalsRowRange() As Range",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/table_api.rs::Bridge::table_api",
        },
        super::MemberDescriptor {
            owner: "ListObject",
            name: "ListRows",
            signature: "Property Get ListRows() As ListRows",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/table_api.rs::Bridge::table_api",
        },
        super::MemberDescriptor {
            owner: "ListObject",
            name: "ListColumns",
            signature: "Property Get ListColumns() As ListColumns",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/table_api.rs::Bridge::table_api",
        },
        super::MemberDescriptor {
            owner: "ListObject",
            name: "AutoFilter",
            signature: "Property Get AutoFilter() As AutoFilter",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/table_api.rs::Bridge::table_api",
        },
        super::MemberDescriptor {
            owner: "ListObject",
            name: "SourceType",
            signature: "Property Get SourceType() As Long",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/table_api.rs::Bridge::table_api",
        },
        super::MemberDescriptor {
            owner: "ListObject",
            name: "ShowHeaders",
            signature: "Property Get ShowHeaders() As Boolean",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/table_api.rs::Bridge::table_api",
        },
        super::MemberDescriptor {
            owner: "ListObject",
            name: "Name",
            signature: "Property Name As String",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/table_api.rs::Bridge::table_api",
        },
        super::MemberDescriptor {
            owner: "ListObject",
            name: "DisplayName",
            signature: "Property DisplayName As String",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/table_api.rs::Bridge::table_api",
        },
        super::MemberDescriptor {
            owner: "ListObject",
            name: "TableStyle",
            signature: "Property TableStyle As Variant",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/table_api.rs::Bridge::table_api",
        },
        super::MemberDescriptor {
            owner: "ListObject",
            name: "ShowTotals",
            signature: "Property ShowTotals As Boolean",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/table_api.rs::Bridge::table_api",
        },
        super::MemberDescriptor {
            owner: "ListObject",
            name: "ShowAutoFilter",
            signature: "Property ShowAutoFilter As Boolean",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/table_api.rs::Bridge::table_api",
        },
        super::MemberDescriptor {
            owner: "ListObject",
            name: "ShowTableStyleRowStripes",
            signature: "Property ShowTableStyleRowStripes As Boolean",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/table_api.rs::Bridge::table_api",
        },
        super::MemberDescriptor {
            owner: "ListObject",
            name: "ShowTableStyleColumnStripes",
            signature: "Property ShowTableStyleColumnStripes As Boolean",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/table_api.rs::Bridge::table_api",
        },
        super::MemberDescriptor {
            owner: "ListObject",
            name: "ShowTableStyleFirstColumn",
            signature: "Property ShowTableStyleFirstColumn As Boolean",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/table_api.rs::Bridge::table_api",
        },
        super::MemberDescriptor {
            owner: "ListObject",
            name: "ShowTableStyleLastColumn",
            signature: "Property ShowTableStyleLastColumn As Boolean",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/table_api.rs::Bridge::table_api",
        },
        super::MemberDescriptor {
            owner: "ListObject",
            name: "Resize",
            signature: "Sub Resize(Range As Range)",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/table_api.rs::Bridge::table_api",
        },
        super::MemberDescriptor {
            owner: "ListObject",
            name: "Unlist",
            signature: "Sub Unlist()",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/table_api.rs::Bridge::table_api",
        },
        super::MemberDescriptor {
            owner: "ListObject",
            name: "Delete",
            signature: "Sub Delete()",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/table_api.rs::Bridge::table_api",
        },
    ]
}
