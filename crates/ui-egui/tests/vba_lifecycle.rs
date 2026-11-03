#![cfg(feature = "vba")]
use gridcraft_engine::{
    Session,
    core::{CellRef, Value},
};
use gridcraft_ui_egui::{SaveAnswer, SheetApp};
use serde_json::json;
type R = Result<(), Box<dyn std::error::Error>>;
fn setup(source: &str) -> Result<SheetApp, Box<dyn std::error::Error>> {
    let mut session = Session::new();
    session.new_workbook();
    std::sync::Arc::make_mut(&mut session.doc_mut()?.wb).code_name = Some("BookCode".into());
    let mut app = SheetApp::new(session, Default::default());
    app.run("vba.open", json!({}))?;
    let editor = app.vba.as_mut().ok_or("editor")?;
    editor.ide.add_document_module("BookCode.cls", "BookCode", &format!("Attribute VB_Name = \"BookCode\"\n{source}"))?;
    for m in &mut editor.ide.modules {
        m.saved_revision = m.revision;
    }
    editor.bridge.configure_document(&mut editor.ide, &app.session.doc()?.wb)?;
    editor.bridge.consent(true);
    Ok(app)
}
fn pump(app: &mut SheetApp) {
    let ctx = egui::Context::default();
    for _ in 0..80 {
        app.logic(&ctx);
    }
}
fn file(name: &str) -> String {
    std::env::temp_dir().join(format!("gridcraft-ui-life-{}-{name}.xlsx", std::process::id())).to_string_lossy().into_owned()
}
#[test]
fn native_save_cancels_before_write_and_reset_discards_parked_commit() -> R {
    let file = file("cancel");
    let _ = std::fs::remove_file(&file);
    let mut app = setup("Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, ByRef Cancel As Boolean)\nCancel = True\nEnd Sub\n")?;
    assert_eq!(app.run("file.saveAs", json!({"path":file}))?["pending"], "vba-lifecycle");
    pump(&mut app);
    assert!(!std::path::Path::new(&file).exists());
    let mut app = setup("Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, ByRef Cancel As Boolean)\nStop\nEnd Sub\n")?;
    app.run("file.saveAs", json!({"path":file}))?;
    pump(&mut app);
    assert_eq!(app.vba.as_ref().ok_or("editor")?.ide.mode(), gridcraft_vba::Mode::Break);
    assert!(!std::path::Path::new(&file).exists());
    app.run("vba.reset", json!({}))?;
    pump(&mut app);
    assert!(!std::path::Path::new(&file).exists());
    Ok(())
}
#[test]
fn native_before_and_after_save_surround_the_physical_write() -> R {
    let file = file("ordered");
    let mut app = setup(
        "Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, ByRef Cancel As Boolean)\nMe.Worksheets(1).Range(\"A1\").Value2 = 41\nEnd Sub\nPrivate Sub Workbook_AfterSave(ByVal Success As Boolean)\nMe.Worksheets(1).Range(\"B1\").Value2 = Success\nEnd Sub\n",
    )?;
    app.run("file.saveAs", json!({"path":file}))?;
    pump(&mut app);
    let (wb, _) = gridcraft_xlsx::read_xlsx(&std::fs::read(&file)?)?;
    assert_eq!(wb.sheet(0).ok_or("sheet")?.value(CellRef::default()), Value::Number(41.0));
    assert_eq!(wb.sheet(0).ok_or("sheet")?.value(CellRef::new(0, 1)), Value::Empty);
    assert_eq!(app.session.doc()?.wb.sheet(0).ok_or("sheet")?.value(CellRef::new(0, 1)), Value::Bool(true));
    std::fs::remove_file(file)?;
    Ok(())
}
#[test]
fn dirty_close_decision_keeps_single_before_close_and_save_continues_asynchronously() -> R {
    let file = file("close");
    let mut app = setup(
        "Private Sub Workbook_BeforeClose(ByRef Cancel As Boolean)\nMe.Worksheets(1).Range(\"A1\").Value2 = Me.Worksheets(1).Range(\"A1\").Value2 + 1\nEnd Sub\nPrivate Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, ByRef Cancel As Boolean)\nEnd Sub\n",
    )?;
    app.session.execute("file.saveAs", json!({"path":file}))?;
    let doc = app.session.doc()?.uid;
    app.run("file.close", json!({"document":true}))?;
    pump(&mut app);
    assert!(app.session.document(doc).is_some());
    assert!(app.answer_save_changes(SaveAnswer::Save, false, &egui::Context::default()));
    pump(&mut app);
    assert!(app.session.document(doc).is_none());
    let (wb, _) = gridcraft_xlsx::read_xlsx(&std::fs::read(&file)?)?;
    assert_eq!(wb.sheet(0).ok_or("sheet")?.value(CellRef::default()), Value::Number(1.0));
    std::fs::remove_file(file)?;
    Ok(())
}

#[test]
fn native_quit_obeys_before_close_cancel_even_for_clean_documents() -> R {
    let mut app = setup("Private Sub Workbook_BeforeClose(ByRef Cancel As Boolean)\nCancel = True\nEnd Sub\n")?;
    let owner = app.session.doc()?.uid;
    app.request_quit(&egui::Context::default());
    pump(&mut app);
    assert!(app.session.document(owner).is_some());
    assert_eq!(app.run("vba.status", json!({}))?["lifecycle"]["state"], "cancelled");
    Ok(())
}
#[test]
fn closing_one_of_two_windows_does_not_fire_before_close() -> R {
    let mut app = setup("Private Sub Workbook_BeforeClose(ByRef Cancel As Boolean)\nCancel = True\nEnd Sub\n")?;
    let doc = app.session.doc()?.uid;
    app.run("view.newWindow", json!({}))?;
    let id = app.session.active_window().ok_or("window")?;
    assert_eq!(app.run("window.close", json!({"window":id}))?["closed"], true);
    pump(&mut app);
    assert!(app.session.document(doc).is_some());
    assert_eq!(app.session.windows().len(), 1);
    Ok(())
}

#[test]
fn native_cancel_outputs_survive_queued_change_callbacks() -> R {
    let file = file("change-cancel");
    let _ = std::fs::remove_file(&file);
    let mut app = setup(
        "Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, ByRef Cancel As Boolean)\nMe.Worksheets(1).Range(\"A1\").Value2 = 1\nCancel = True\nEnd Sub\n",
    )?;
    std::sync::Arc::make_mut(&mut app.session.doc_mut()?.wb).sheet_mut(0).ok_or("sheet")?.code_name = Some("SheetCode".into());
    let e = app.vba.as_mut().ok_or("editor")?;
    e.ide.add_document_module(
        "SheetCode.cls",
        "SheetCode",
        "Attribute VB_Name = \"SheetCode\"\nPrivate Sub Worksheet_Change(ByVal Target As Range)\nDebug.Print \"changed\"\nEnd Sub\n",
    )?;
    e.bridge.configure_document(&mut e.ide, &app.session.doc()?.wb)?;
    app.run("file.saveAs", json!({"path":file}))?;
    pump(&mut app);
    assert!(!std::path::Path::new(&file).exists());
    assert_eq!(app.vba.as_ref().ok_or("editor")?.ide.output(), ["changed"]);
    Ok(())
}
#[test]
fn native_unsaved_close_uses_picker_once_and_does_not_repeat_before_save() -> R {
    let file = file("unsaved-close");
    let chosen = file.clone();
    let mut app = setup(
        "Private Sub Workbook_BeforeClose(ByRef Cancel As Boolean)\nEnd Sub\nPrivate Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, ByRef Cancel As Boolean)\nMe.Worksheets(1).Range(\"A1\").Value2 = Me.Worksheets(1).Range(\"A1\").Value2 + 1\nEnd Sub\n",
    )?;
    app.services.pick_save = Some(Box::new(move |_| Some(chosen.clone())));
    app.session.execute("cell.set", json!({"cell":"C1","input":"1"}))?;
    let owner = app.session.doc()?.uid;
    app.run("file.close", json!({"document":true}))?;
    pump(&mut app);
    app.answer_save_changes(SaveAnswer::Save, false, &egui::Context::default());
    pump(&mut app);
    assert!(app.session.document(owner).is_none());
    let (wb, _) = gridcraft_xlsx::read_xlsx(&std::fs::read(&file)?)?;
    assert_eq!(wb.sheet(0).ok_or("sheet")?.value(CellRef::default()), Value::Number(1.0));
    std::fs::remove_file(file)?;
    Ok(())
}

#[test]
fn native_io_failure_retains_failed_status_after_after_save_false() -> R {
    let mut app = setup("Private Sub Workbook_AfterSave(ByVal Success As Boolean)\nMe.Worksheets(1).Range(\"B1\").Value2 = Success\nEnd Sub\n")?;
    let file = std::env::temp_dir().join(format!("gridcraft-ui-no-parent-{}/book.xlsx", std::process::id()));
    app.run("file.saveAs", json!({"path":file}))?;
    pump(&mut app);
    assert_eq!(app.session.doc()?.wb.sheet(0).ok_or("sheet")?.value(CellRef::new(0, 1)), Value::Bool(false));
    assert_eq!(app.run("vba.status", json!({}))?["lifecycle"]["state"], "failed");
    Ok(())
}

#[test]
fn native_save_keeps_target_identity_when_before_save_activates_another_book() -> R {
    let file = file("activation");
    let mut app = setup(
        "Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, ByRef Cancel As Boolean)\nMe.Worksheets(1).Range(\"A1\").Value2 = 7\nWorkbooks(2).Activate\nEnd Sub\n",
    )?;
    let owner = app.session.doc()?.uid;
    app.session.new_workbook();
    let other = app.session.doc()?.uid;
    let owner_index = app.session.document_index(owner).ok_or("owner")?;
    app.session.set_active(owner_index);
    app.run("file.saveAs", json!({"path":file}))?;
    pump(&mut app);
    assert_eq!(app.session.doc()?.uid, other);
    assert!(app.session.doc()?.path.is_none());
    assert_eq!(app.session.document(owner).ok_or("owner")?.path.as_deref(), Some(file.as_str()));
    let (wb, _) = gridcraft_xlsx::read_xlsx(&std::fs::read(&file)?)?;
    assert_eq!(wb.sheet(0).ok_or("sheet")?.value(CellRef::default()), Value::Number(7.0));
    std::fs::remove_file(file)?;
    Ok(())
}

#[test]
fn native_dirty_close_decision_is_bound_to_its_dialog_owner_after_activation() -> R {
    let mut app = setup("Private Sub Workbook_BeforeClose(ByRef Cancel As Boolean)\nEnd Sub\n")?;
    app.session.execute("cell.set", json!({"cell":"A1","input":"1"}))?;
    let owner = app.session.doc()?.uid;
    app.run("file.close", json!({"document":true}))?;
    pump(&mut app);
    app.session.new_workbook();
    let other = app.session.doc()?.uid;
    app.answer_save_changes(SaveAnswer::DontSave, false, &egui::Context::default());
    pump(&mut app);
    assert!(app.session.document(owner).is_none());
    assert_eq!(app.session.doc()?.uid, other);
    Ok(())
}
