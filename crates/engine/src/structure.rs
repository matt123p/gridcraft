//! Native structural commands. Serialization is deferred to successful journaling.
use crate::{
    EngineError, Result, Session, cmd,
    core::{MAX_COLS, MAX_ROWS, RangeRef},
    formula::adjust::{Axis, Edit},
};
use serde_json::{Value as Json, json};
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum StructuralEdit {
    InsertRows,
    InsertColumns,
    DeleteRows,
    DeleteColumns,
    InsertDown,
    InsertRight,
    DeleteUp,
    DeleteLeft,
}
impl StructuralEdit {
    fn command(self) -> &'static str {
        use StructuralEdit::*;
        match self {
            InsertRows => "home.insertRows",
            InsertColumns => "home.insertColumns",
            DeleteRows => "home.deleteRows",
            DeleteColumns => "home.deleteColumns",
            InsertDown | InsertRight => "home.insertCells",
            DeleteUp | DeleteLeft => "home.deleteCells",
        }
    }
    pub fn reference_edit(self, r: RangeRef) -> Result<Edit> {
        if !r.start.is_valid() || !r.end.is_valid() || r.start.row > r.end.row || r.start.col > r.end.col {
            return Err(EngineError::BadParams { cmd: self.command().into(), msg: "invalid structural reference rectangle".into() });
        }
        use StructuralEdit::*;
        Ok(match self {
            InsertRows => Edit::Insert { axis: Axis::Rows, at: r.start.row, count: r.height() },
            InsertColumns => Edit::Insert { axis: Axis::Cols, at: r.start.col, count: r.width() },
            DeleteRows => Edit::Delete { axis: Axis::Rows, at: r.start.row, count: r.height() },
            DeleteColumns => Edit::Delete { axis: Axis::Cols, at: r.start.col, count: r.width() },
            InsertDown | InsertRight => Edit::InsertCells { axis: if self == InsertRight { Axis::Cols } else { Axis::Rows }, range: r },
            DeleteUp | DeleteLeft => Edit::DeleteCells { axis: if self == DeleteLeft { Axis::Cols } else { Axis::Rows }, range: r },
        })
    }
}
impl Session {
    /// Execute the same structural implementation as the ribbon/control command.
    pub fn edit_range_structure(&mut self, sheet: usize, range: RangeRef, operation: StructuralEdit) -> Result<Json> {
        let bad = || EngineError::BadParams { cmd: operation.command().into(), msg: "invalid structural rectangle or sheet".into() };
        if !range.start.is_valid()
            || !range.end.is_valid()
            || range.start.row > range.end.row
            || range.start.col > range.end.col
            || self.doc()?.wb.sheet(sheet).is_none()
        {
            return Err(bad());
        }
        // Whole-grid insertion and unbounded format inheritance cannot be useful.
        if matches!(operation, StructuralEdit::InsertRows) && range.height() == MAX_ROWS
            || matches!(operation, StructuralEdit::InsertColumns) && range.width() == MAX_COLS
        {
            return Err(bad());
        }
        self.execute_with(
            operation.command(),
            || {
                std::borrow::Cow::Owned(match operation {
                    StructuralEdit::InsertRows | StructuralEdit::DeleteRows => {
                        json!({"sheet":sheet,"rows":RangeRef::rows(range.start.row,range.end.row).a1()})
                    }
                    StructuralEdit::InsertColumns | StructuralEdit::DeleteColumns => {
                        json!({"sheet":sheet,"cols":RangeRef::cols(range.start.col,range.end.col).a1()})
                    }
                    StructuralEdit::InsertDown => json!({"sheet":sheet,"range":range.a1(),"shift":"down"}),
                    StructuralEdit::InsertRight => json!({"sheet":sheet,"range":range.a1(),"shift":"right"}),
                    StructuralEdit::DeleteUp => json!({"sheet":sheet,"range":range.a1(),"shift":"up"}),
                    StructuralEdit::DeleteLeft => json!({"sheet":sheet,"range":range.a1(),"shift":"left"}),
                })
            },
            |session, _| cmd::sheet::structure(session, sheet, range, operation),
        )
    }
}
