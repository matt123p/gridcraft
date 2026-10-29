use gridcraft_engine::{
    CellContent, Session, StyleChange,
    core::{CellRef, RangeRef, Value},
};
use serde_json::json;
type R = Result<(), Box<dyn std::error::Error>>;
#[test]
fn native_styles_and_typed_formats_replay_and_undo_atomically() -> R {
    let mut s = Session::new();
    s.new_workbook();
    let r = RangeRef::cell(CellRef::default());
    s.set_range_contents_with_formats(0, r, &[CellContent::Value(Value::Number(45000.))], &[Some("m/d/yyyy".into())])?;
    s.set_range_style(0, r, StyleChange::Bold(true))?;
    let mut replay = Session::new();
    replay.new_workbook();
    for (id, params) in &s.journal {
        replay.execute(id, params.clone())?;
    }
    assert_eq!(*s.doc()?.wb, *replay.doc()?.wb);
    s.execute("edit.undo", json!({}))?;
    let sh = s.doc()?.wb.sheet(0).ok_or("sheet")?;
    let style = s.doc()?.wb.styles.get(sh.style_id(CellRef::default()));
    assert!(!style.font.bold);
    assert_eq!(style.num_fmt.as_str(), "m/d/yyyy");
    s.execute("edit.undo", json!({}))?;
    assert!(s.doc()?.wb.sheet(0).ok_or("sheet")?.value(CellRef::default()).is_empty());
    let before = s.doc()?.wb.clone();
    assert!(s.execute("range.setContents", json!({"sheet":0,"range":"A1","contents":[{"value":{"t":"Number","v":1}}],"formats":["0","0"]})).is_err());
    assert_eq!(*before, *s.doc()?.wb);
    assert!(s.set_range_style(0, r, StyleChange::FontSize(f32::NAN)).is_err());
    assert_eq!(*before, *s.doc()?.wb);
    s.execute("review.protectSheet", json!({}))?;
    let before = s.doc()?.wb.clone();
    assert!(s.set_range_style(0, r, StyleChange::Bold(true)).is_err());
    assert_eq!(*before, *s.doc()?.wb);
    Ok(())
}
