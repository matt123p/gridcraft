//! Host-neutral compositions of existing table resize/convert and structural cell commands.
use crate::{
    EngineError, Result, Session, StructuralEdit,
    core::{CellRef, MAX_COLS, MAX_ROWS, RangeRef},
};
use serde_json::{Value as Json, json};
pub(crate) fn mutate(s: &mut Session, p: &Json, operation: &str) -> Result<Json> {
    let bad = |message: &str| EngineError::BadParams { cmd: format!("table.{operation}"), msg: message.into() };
    let name = p.get("table").and_then(Json::as_str).ok_or_else(|| bad("missing table"))?;
    let doc = s.doc()?;
    let (sheet, index) = doc.wb.table(name).ok_or_else(|| bad("table not found"))?;
    let sh = doc.wb.sheet(sheet).ok_or_else(|| bad("worksheet deleted"))?;
    let table = sh.tables.get(index).ok_or_else(|| bad("table deleted"))?.clone();
    if sh.is_protected() {
        return Err(bad("table changes require an unprotected worksheet"));
    }
    if table.range.count() > 65536 {
        return Err(bad("table cell budget"));
    }
    if table.connection_id.is_some() || table.table_type.is_some() {
        return Err(bad("provider-backed table changes are unavailable"));
    }
    if operation.starts_with("insert") {
        let grown = if operation.ends_with("Row") {
            u64::from(table.range.height() + 1) * u64::from(table.range.width())
        } else {
            u64::from(table.range.width() + 1) * u64::from(table.range.height())
        };
        if grown > 65536 {
            return Err(bad("grown table cell budget"));
        }
    }
    if operation == "delete" {
        s.execute("table.convertToRange", json!({"table":name}))?;
        s.execute("edit.clearAll", json!({"sheet":sheet,"range":table.range.a1()}))?;
        return Ok(json!({"deleted":true}));
    }
    let row = operation.ends_with("Row");
    let insert = operation.starts_with("insert");
    let count = if row { table.data_range().map(|r| r.height()).unwrap_or(0) } else { table.range.width() };
    let position = match p.get("position") {
        None => {
            if insert {
                count + 1
            } else {
                1
            }
        }
        Some(value) => value.as_u64().and_then(|n| u32::try_from(n).ok()).ok_or_else(|| bad("invalid position"))?,
    };
    if position < 1 || position > count + u32::from(insert) {
        return Err(bad("position outside table"));
    }
    let coordinate = if row {
        table.range.start.row.checked_add(u32::from(table.header_row)).and_then(|n| n.checked_add(position - 1))
    } else {
        table.range.start.col.checked_add(position - 1)
    }
    .ok_or_else(|| bad("coordinate overflow"))?;
    if coordinate >= if row { MAX_ROWS } else { MAX_COLS } {
        return Err(bad("table is at worksheet boundary"));
    }
    let rect = if row {
        RangeRef::new(CellRef::new(coordinate, table.range.start.col), CellRef::new(coordinate, table.range.end.col))
    } else {
        RangeRef::new(CellRef::new(table.range.start.row, coordinate), CellRef::new(table.range.end.row, coordinate))
    };
    let append = position == count + 1 && (!row || !table.totals_row);
    let shift = !(insert
        && row
        && append
        && p.get("alwaysInsert").and_then(Json::as_bool) == Some(false)
        && rect.iter().all(|at| sh.cell(at).is_none_or(|c| c.formula.is_none() && c.value.is_empty())));
    if insert && (if row { table.range.end.row + 1 >= MAX_ROWS } else { table.range.end.col + 1 >= MAX_COLS }) {
        return Err(bad("table cannot grow past worksheet"));
    }
    if shift {
        s.edit_range_structure(
            sheet,
            rect,
            match (row, insert) {
                (true, true) => StructuralEdit::InsertDown,
                (true, false) => StructuralEdit::DeleteUp,
                (false, true) => StructuralEdit::InsertRight,
                (false, false) => StructuralEdit::DeleteLeft,
            },
        )?;
    }
    if insert && append {
        let end =
            if row { CellRef::new(table.range.end.row + 1, table.range.end.col) } else { CellRef::new(table.range.end.row, table.range.end.col + 1) };
        s.execute("table.resize", json!({"table":name,"range":RangeRef::new(table.range.start,end).a1()}))?;
    }
    Ok(json!({"range":rect.a1(),"shift":shift,"position":position}))
}
