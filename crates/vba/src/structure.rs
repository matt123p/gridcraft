//! Excel arguments end here; the engine owns protection, movement, undo and recalculation.
use super::objects::{arguments, numeric_index};
use super::ranges::Range;
use super::*;
use gridcraft_engine::StructuralEdit;
impl Bridge {
    pub(crate) fn structural(&mut self, session: &mut Session, sheet: usize, range: &Range, call: &HostCall) -> HostResult<Value> {
        let insert = call.member.eq_ignore_ascii_case("Insert");
        let args = arguments(call, if insert { &["Shift", "CopyOrigin"] } else { &["Shift"] })?;
        let r = range.rectangle()?;
        if r.count() == 17_179_869_184 {
            return Err(failure(1004, "cannot insert/delete the entire worksheet"));
        }
        if let Some(v) = args.get(1).filter(|v| !matches!(v, Value::Missing))
            && numeric_index(v)? != 0
        {
            return Err(failure(1004, "CopyOrigin from below/right is unsupported"));
        }
        let shift = args.first().filter(|v| !matches!(v, Value::Missing)).map(numeric_index).transpose()?;
        let operation = match (insert, r.is_full_rows(), r.is_full_cols(), shift) {
            (true, true, _, None | Some(-4121)) => StructuralEdit::InsertRows,
            (false, true, _, None | Some(-4162)) => StructuralEdit::DeleteRows,
            (true, _, true, None | Some(-4161)) => StructuralEdit::InsertColumns,
            (false, _, true, None | Some(-4159)) => StructuralEdit::DeleteColumns,
            (true, false, false, Some(-4121)) => StructuralEdit::InsertDown,
            (true, false, false, Some(-4161)) => StructuralEdit::InsertRight,
            (false, false, false, Some(-4162)) => StructuralEdit::DeleteUp,
            (false, false, false, Some(-4159)) => StructuralEdit::DeleteLeft,
            _ => return Err(failure(1004, "a supported explicit Shift is required for partial-cell edits")),
        };
        // Bound line-format inheritance and reference maintenance before mutation.
        let lines = if r.is_full_rows() {
            r.height()
        } else if r.is_full_cols() {
            r.width()
        } else {
            1
        };
        if lines > MAX_CELLS as u32 {
            return Err(failure(7, "structural line budget"));
        }
        let edit = operation.reference_edit(r).map_err(|e| failure(1004, e.to_string()))?;
        let rebased = self.objects.rebase(self.context, sheet, &edit)?;
        session.edit_range_structure(sheet, r, operation).map_err(|e| failure(1004, e.to_string()))?;
        self.objects.replace(rebased)?;
        self.range_epochs.insert(self.context, session.doc().map_err(|_| failure(91, "workbook closed"))?.range_epoch);
        Ok(Value::Empty)
    }
}
