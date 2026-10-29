//! Native structural calls and recorded command envelopes must share behavior.
use gridcraft_engine::{
    Session, StructuralEdit,
    core::{CellRef, RangeRef, Value},
};
use serde_json::json;
type R = Result<(), Box<dyn std::error::Error>>;
#[test]
fn typed_structural_commands_match_envelopes_and_replay_other_sheet() -> R {
    let mut s = Session::new();
    s.new_workbook();
    s.execute("home.insertSheet", json!({"name":"Other"}))?;
    s.execute("sheet.activate", json!({"sheet":0}))?;
    s.execute("cell.set", json!({"sheet":1,"cell":"B2","input":"11"}))?;
    s.execute("cell.set", json!({"sheet":0,"cell":"C1","input":"=Other!B2*2"}))?;
    let before = s.doc()?.wb.clone();
    s.edit_range_structure(1, RangeRef::rows(1, 1), StructuralEdit::InsertRows)?;
    assert_eq!(s.doc()?.wb.sheet(1).ok_or("sheet")?.value(CellRef::new(2, 1)), Value::Number(11.));
    assert_eq!(s.doc()?.wb.sheet(0).ok_or("sheet")?.value(CellRef::new(0, 2)), Value::Number(22.));
    let after = s.doc()?.wb.clone();
    let recorded = s.journal.last().cloned().ok_or("journal")?;
    assert_eq!(recorded.0, "home.insertRows");
    assert_eq!(recorded.1["sheet"], 1);
    s.execute("edit.undo", json!({}))?;
    assert_eq!(*before, *s.doc()?.wb);
    s.execute(&recorded.0, recorded.1)?;
    assert_eq!(*after, *s.doc()?.wb);
    Ok(())
}
#[test]
fn native_structure_validates_sheet_bounds_and_protection_before_effects() -> R {
    let mut s = Session::new();
    s.new_workbook();
    s.execute("review.protectSheet", json!({}))?;
    let before = s.doc()?.wb.clone();
    let undo = s.doc()?.undo.len();
    let journal = s.journal.len();
    for op in [StructuralEdit::InsertRows, StructuralEdit::DeleteRows, StructuralEdit::InsertDown, StructuralEdit::DeleteLeft] {
        assert!(s.edit_range_structure(0, RangeRef::cell(CellRef::new(1, 1)), op).is_err());
    }
    assert!(StructuralEdit::InsertRows.reference_edit(RangeRef { start: CellRef::new(9, 9), end: CellRef::new(0, 0) }).is_err());
    assert!(s.edit_range_structure(usize::MAX, RangeRef::default(), StructuralEdit::InsertRows).is_err());
    assert!(s.edit_range_structure(0, RangeRef { start: CellRef::new(9, 9), end: CellRef::new(0, 0) }, StructuralEdit::DeleteUp).is_err());
    assert_eq!(*before, *s.doc()?.wb);
    assert_eq!(s.doc()?.undo.len(), undo);
    assert_eq!(s.journal.len(), journal);
    Ok(())
}
