use gridcraft_engine::{Session, model::Visibility};
use serde_json::json;

type R = Result<(), Box<dyn std::error::Error>>;

#[test]
fn hidden_sheet_delete_keeps_the_last_visible_sheet_and_is_undoable() -> R {
    let mut session = Session::new();
    session.new_workbook();
    session.execute("home.insertSheet", json!({"name":"Hidden"}))?;
    session.set_worksheet_visibility(1, Visibility::VeryHidden)?;
    session.execute("home.deleteSheet", json!({"sheet":"Hidden"}))?;
    assert_eq!(session.doc()?.wb.sheets.len(), 1);
    assert_eq!(session.doc()?.wb.sheet(0).ok_or("sheet")?.visibility, Visibility::Visible);
    assert!(session.execute("home.deleteSheet", json!({"sheet":0})).is_err());
    session.execute("edit.undo", json!({}))?;
    assert_eq!(session.doc()?.wb.sheet(1).ok_or("hidden sheet")?.visibility, Visibility::VeryHidden);
    Ok(())
}

#[test]
fn explicit_new_sheet_count_is_replayable_and_invalid_count_has_no_effect() -> R {
    let mut session = Session::new();
    session.prefs.sheets_in_new_workbook = 4;
    session.execute("file.new", json!({"sheets":1}))?;
    assert_eq!(session.doc()?.wb.sheets.len(), 1);
    assert_eq!(session.prefs.sheets_in_new_workbook, 4);
    let count = session.documents().len();
    for invalid in [json!(0), json!(256), json!(-1), json!(1.5), json!("1")] {
        assert!(session.execute("file.new", json!({"sheets":invalid})).is_err());
        assert_eq!(session.documents().len(), count);
    }
    session.execute("file.new", json!({}))?;
    assert_eq!(session.doc()?.wb.sheets.len(), 4);
    Ok(())
}
