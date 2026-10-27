use gridcraft_engine::{
    CellContent, MAX_CONTENT_BYTES, MAX_CONTENT_CELLS, Session,
    core::{CellError, CellRef, RangeRef, Value},
};
use serde_json::json;

fn session() -> Session {
    let mut s = Session::new();
    s.new_workbook();
    s
}
fn range(text: &str) -> RangeRef {
    RangeRef::parse(text).unwrap()
}
fn value(s: &Session, at: &str) -> Value {
    s.doc().unwrap().wb.sheet(0).unwrap().value(CellRef::parse(at).unwrap())
}
fn constant(value: Value) -> CellContent {
    CellContent::Value(value)
}

#[test]
fn native_contents_preserve_exact_constants_and_formats_with_undo_and_recording() {
    let mut s = session();
    s.execute("home.numberFormat", json!({"range":"A1:F1","format":"@"})).unwrap();
    let contents = vec![
        constant(Value::text("=literal")),
        constant(Value::text("'quoted")),
        constant(Value::text("00123")),
        constant(Value::Error(CellError::Div0)),
        constant(Value::Bool(true)),
        constant(Value::Empty),
    ];
    let before = s.doc().unwrap().wb.clone();
    let undo = s.doc().unwrap().undo.len();
    s.start_recording();
    s.set_range_contents(0, range("A1:F1"), &contents).unwrap();
    let recording = s.stop_recording().unwrap();
    assert_eq!(recording.steps.len(), 1);
    let step = recording.steps.first().unwrap();
    assert_eq!(step.id, "range.setContents");
    assert_eq!(s.doc().unwrap().undo.len(), undo + 1);
    for (at, content) in range("A1:F1").iter().zip(&contents) {
        let CellContent::Value(expected) = content else { panic!("constant") };
        let d = s.doc().unwrap();
        let sh = d.wb.sheet(0).unwrap();
        assert_eq!(sh.value(at), *expected);
        assert_eq!(d.wb.styles.get(sh.style_id(at)).num_fmt.as_str(), "@");
    }
    let after = s.doc().unwrap().wb.clone();
    s.execute("edit.undo", json!({})).unwrap();
    assert_eq!(*s.doc().unwrap().wb, *before);
    s.execute("edit.redo", json!({})).unwrap();
    assert_eq!(*s.doc().unwrap().wb, *after);
    s.execute("edit.undo", json!({})).unwrap();
    s.execute(&step.id, step.params.clone()).unwrap();
    assert_eq!(*s.doc().unwrap().wb, *after);
}

#[test]
fn explicit_formulas_ignore_text_format_recalculate_and_fail_atomically() {
    let mut s = session();
    s.execute("home.numberFormat", json!({"range":"A1:C1","format":"@"})).unwrap();
    s.set_range_contents(0, range("A1:C1"), &[constant(Value::Number(5.0)), CellContent::Formula("=A1*2".into()), constant(Value::text("=A1*2"))])
        .unwrap();
    assert_eq!(value(&s, "B1"), Value::Number(10.0));
    assert_eq!(value(&s, "C1"), Value::text("=A1*2"));
    let before = s.doc().unwrap().wb.clone();
    let undo = s.doc().unwrap().undo.len();
    let journal = s.journal.len();
    assert!(s.set_range_contents(0, range("A1:B1"), &[constant(Value::Number(99.0)), CellContent::Formula("=1+*2".into())]).is_err());
    assert_eq!(*s.doc().unwrap().wb, *before);
    assert_eq!(s.doc().unwrap().undo.len(), undo);
    assert_eq!(s.journal.len(), journal);
    s.set_range_contents(0, range("A1"), &[constant(Value::Number(7.0))]).unwrap();
    assert_eq!(value(&s, "B1"), Value::Number(14.0));
    let unicode = "é".repeat(5000);
    s.set_range_contents(0, range("C1"), &[CellContent::Formula(format!("=\"{unicode}\""))]).unwrap();
    assert_eq!(value(&s, "C1"), Value::text(unicode));
}

#[test]
fn native_and_wire_paths_have_same_budgets_and_refuse_hostile_input_without_effects() {
    let mut s = session();
    let before = s.doc().unwrap().wb.clone();
    let inputs = [
        json!({}),
        json!({"range":"A1","contents":[{"value":{"t":"Array","v":[]}}]}),
        json!({"range":"A1","contents":[{"value":{"t":"Empty","v":null}}]}),
        json!({"range":"A1","contents":[{"value":{"t":"Text","v":3}}]}),
        json!({"range":"A1","contents":[{"value":{"t":"Empty"},"formula":"=1"}]}),
        json!({"range":"A1","sheet":true,"contents":[{"value":{"t":"Empty"}}]}),
        json!({"range":"A1","sheet":4294967296u64,"contents":[{"value":{"t":"Empty"}}]}),
        json!({"range":"A1","contents":[{"value":{"t":"Error","v":"not-an-error"}}]}),
        json!({"range":"A1,B1","contents":[{"value":{"t":"Empty"}}]}),
    ];
    for p in inputs {
        assert!(s.execute("range.setContents", p).is_err());
        assert_eq!(*s.doc().unwrap().wb, *before);
    }
    for n in [f64::NAN, f64::INFINITY, f64::NEG_INFINITY] {
        assert!(s.set_range_contents(0, range("A1"), &[constant(Value::Number(n))]).is_err());
    }
    let array = gridcraft_engine::core::Array::new(1, 1, vec![Value::Number(1.0)]).unwrap();
    assert!(s.set_range_contents(0, range("A1"), &[constant(Value::from(array))]).is_err());
    assert!(s.set_range_contents(usize::MAX, range("A1"), &[constant(Value::Empty)]).is_err());
    let invalid = RangeRef { start: CellRef::new(1, 0), end: CellRef::new(0, 0) };
    assert!(s.set_range_contents(0, invalid, &[constant(Value::Empty)]).is_err());
    let invalid = RangeRef::cell(CellRef::new(u32::MAX, u32::MAX));
    assert!(s.set_range_contents(0, invalid, &[constant(Value::Empty)]).is_err());
    assert!(s.set_range_contents(0, range("A1:B1"), &[constant(Value::Empty)]).is_err());
    assert!(s.set_range_contents(0, range("A1:A65537"), &vec![constant(Value::Empty); MAX_CONTENT_CELLS + 1]).is_err());
    let oversized = vec![constant(Value::text("x".repeat(MAX_CONTENT_BYTES)))];
    assert!(s.set_range_contents(0, range("A1"), &oversized).is_err());
    assert!(s.execute("range.setContents", json!({"range":"A1","contents":oversized})).is_err());
    assert_eq!(*s.doc().unwrap().wb, *before);
    let text = "x".repeat(MAX_CONTENT_BYTES - 64);
    let boundary = vec![constant(Value::text(text))];
    s.set_range_contents(0, range("A1"), &boundary).unwrap();
    let native = s.doc().unwrap().wb.clone();
    s.execute("range.setContents", json!({"range":"A1","contents":boundary})).unwrap();
    assert_eq!(*s.doc().unwrap().wb, *native);
    let text = "x".repeat(gridcraft_engine::formula::MAX_FORMULA_LEN - 2);
    let formula = [CellContent::Formula(format!("=\"{text}\""))];
    s.set_range_contents(0, range("B1"), &formula).unwrap();
    assert_eq!(value(&s, "B1"), Value::text(text));
    let oversized = [CellContent::Formula(format!("=\"{}\"", "x".repeat(gridcraft_engine::formula::MAX_FORMULA_LEN)))];
    assert!(s.set_range_contents(0, range("B1"), &oversized).is_err());
    assert!(s.execute("range.setContents", json!({"range":"B1","contents":oversized})).is_err());
}

#[test]
fn protected_cells_refuse_the_entire_native_or_wire_write_and_groups_are_retained() {
    let mut s = session();
    s.execute("review.protectSheet", json!({})).unwrap();
    let before = s.doc().unwrap().wb.clone();
    let contents = [constant(Value::Number(1.0)), constant(Value::text("x"))];
    assert!(s.set_range_contents(0, range("A1:B1"), &contents).is_err());
    assert!(s.execute("range.setContents", json!({"range":"A1:B1","contents":contents})).is_err());
    assert_eq!(*s.doc().unwrap().wb, *before);
    s.execute("review.unprotectSheet", json!({})).unwrap();
    let undo = s.doc().unwrap().undo.len();
    let token = s.group_begin("Native host").unwrap();
    s.set_range_contents(0, range("A1:B1"), &contents).unwrap();
    assert!(s.execute("edit.undo", json!({})).is_err());
    s.group_end(token).unwrap();
    assert_eq!(s.doc().unwrap().undo.len(), undo + 1);
    s.execute("edit.undo", json!({})).unwrap();
    assert!(value(&s, "A1").is_empty());
}

#[test]
fn wire_sheet_resolution_matches_native_and_rejects_conflicting_qualified_ranges() {
    let mut s = session();
    s.execute("home.insertSheet", json!({})).unwrap();
    let name = s.doc().unwrap().wb.sheet(0).unwrap().name.clone();
    let contents = [constant(Value::Number(21.0))];
    s.set_range_contents(0, range("B2"), &contents).unwrap();
    let before = s.doc().unwrap().wb.clone();
    s.execute("range.setContents", json!({"sheet":name,"range":"B2","contents":contents})).unwrap();
    assert_eq!(*s.doc().unwrap().wb, *before);
    assert!(s.execute("range.setContents", json!({"sheet":1,"range":format!("'{name}'!B2"),"contents":contents})).is_err());
    assert_eq!(*s.doc().unwrap().wb, *before);
}
