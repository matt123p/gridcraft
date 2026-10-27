use gridcraft_engine::{
    Session,
    core::{CellRef, Value},
};
use serde_json::json;

#[test]
fn assignment_is_inert_undoable_and_round_trips_escaped_drawingml_and_json() {
    let mut s = Session::new();
    s.new_workbook();
    let id = s.execute("insert.shape", json!({"text":"Action"})).unwrap()["shape"].clone();
    let entry = "[0]!Module.Procedure\"&<>";
    s.execute("shape.setMacro", json!({"id":id,"entry":entry})).unwrap();
    assert_eq!(s.doc().unwrap().wb.active().unwrap().shapes[0].on_action.as_deref(), Some(entry));
    assert_eq!(s.doc().unwrap().wb.active().unwrap().value(CellRef::new(0, 0)), Value::Empty);
    let bytes = gridcraft_xlsx::write_xlsx(&s.doc().unwrap().wb).unwrap();
    let (wb, _) = gridcraft_xlsx::read_xlsx(&bytes).unwrap();
    assert_eq!(wb.active().unwrap().shapes[0].on_action.as_deref(), Some(entry));
    let encoded = serde_json::to_value(&wb).unwrap();
    let decoded: gridcraft_engine::model::Workbook = serde_json::from_value(encoded).unwrap();
    assert_eq!(decoded.active().unwrap().shapes[0].on_action.as_deref(), Some(entry));
    s.execute("edit.undo", json!({})).unwrap();
    assert!(s.doc().unwrap().wb.active().unwrap().shapes[0].on_action.is_none());
    s.execute("edit.redo", json!({})).unwrap();
    s.execute("shape.setMacro", json!({"id":id,"entry":null})).unwrap();
    assert!(s.doc().unwrap().wb.active().unwrap().shapes[0].on_action.is_none());
}
#[test]
fn malformed_or_protected_assignments_fail_without_modifying_the_shape() {
    let mut s = Session::new();
    s.new_workbook();
    let id = s.execute("insert.shape", json!({})).unwrap()["shape"].clone();
    let before = s.doc().unwrap().wb.clone();
    for p in [
        json!({}),
        json!({"id":u64::MAX,"entry":"Fill"}),
        json!({"id":id,"entry":23}),
        json!({"id":id,"entry":"x".repeat(257)}),
        json!({"id":id,"entry":"Fill\n"}),
    ] {
        assert!(s.execute("shape.setMacro", p).is_err());
        assert_eq!(s.doc().unwrap().wb, before);
    }
    s.execute("review.protectSheet", json!({})).unwrap();
    let protected = s.doc().unwrap().wb.clone();
    assert!(s.execute("shape.setMacro", json!({"id":id,"entry":"Fill"})).is_err());
    assert_eq!(s.doc().unwrap().wb, protected);
}
