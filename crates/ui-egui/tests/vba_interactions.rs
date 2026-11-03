#![cfg(feature = "vba")]
use gridcraft_engine::{
    Session,
    core::{CellRef, Value},
};
use gridcraft_ui_egui::{SheetApp, theme};
use gridcraft_vba::Mode;
use serde_json::json;
use std::{cell::RefCell, rc::Rc};
type R = Result<(), Box<dyn std::error::Error>>;
fn app(source: &str) -> Result<SheetApp, Box<dyn std::error::Error>> {
    let mut s = Session::new();
    s.new_workbook();
    let mut a = SheetApp::new(s, Default::default());
    a.run("vba.open", json!({"source":source}))?;
    Ok(a)
}
fn control(a: &mut SheetApp, id: &str, p: serde_json::Value) -> Result<serde_json::Value, String> {
    match gridcraft_ui_egui::control::handle(a, &egui::Context::default(), id, &p) {
        gridcraft_ui_egui::control::Outcome::Done(v) if v["ok"] == true => Ok(v["result"].clone()),
        gridcraft_ui_egui::control::Outcome::Done(v) => Err(v["error"].as_str().unwrap_or("control error").into()),
        _ => Err("unexpected screenshot".into()),
    }
}
fn pump(a: &mut SheetApp) {
    let ctx = egui::Context::default();
    for _ in 0..100 {
        gridcraft_ui_egui::vba::tick(a, &ctx);
    }
}
fn complete(a: &SheetApp) -> R {
    let e = a.vba.as_ref().ok_or("editor")?;
    assert_eq!(e.ide.mode(), Mode::Completed, "{:?}", e.ide.diagnostics);
    Ok(())
}
#[test]
fn existing_column_width_dialog_waits_for_real_confirmation_and_keeps_failed_edits_open() -> R {
    let mut a = app("Sub Main()\nRange(\"A1\").Value2=Application.Dialogs(xlDialogColumnWidth).Show\nEnd Sub")?;
    a.run("vba.run", json!({}))?;
    pump(&mut a);
    assert_eq!(a.vba.as_ref().ok_or("editor")?.ide.mode(), Mode::Waiting);
    assert!(!a.vba.as_ref().ok_or("editor")?.visible);
    assert_eq!(a.dialog.as_ref().ok_or("dialog")?.name(), "columnWidth");
    assert!(a.run("cell.set", json!({"cell":"B2","input":"blocked"})).is_err());
    control(&mut a, "ui.dialog.set", json!({"field":"chars","value":-10}))?;
    assert!(control(&mut a, "ui.dialog.confirm", json!({})).is_err());
    pump(&mut a);
    assert!(a.dialog.is_some());
    control(&mut a, "ui.dialog.set", json!({"field":"chars","value":20}))?;
    control(&mut a, "ui.dialog.confirm", json!({}))?;
    pump(&mut a);
    complete(&a)?;
    assert!(a.dialog.is_none());
    assert!(a.vba.as_ref().ok_or("editor")?.visible);
    assert_eq!(a.session.doc()?.wb.sheet(0).ok_or("sheet")?.value(CellRef::new(0, 0)), Value::Bool(true));
    a.run("edit.undo", json!({}))?;
    assert!(a.session.doc()?.wb.sheet(0).ok_or("sheet")?.value(CellRef::new(0, 0)).is_empty());
    Ok(())
}
#[test]
fn cancel_reset_and_project_close_remove_owned_modal_without_replaying_it() -> R {
    for action in ["cancel", "reset", "close"] {
        let mut a = app("Sub Main()\nRange(\"A1\").Value2=Application.Dialogs(xlDialogColumnWidth).Show\nEnd Sub")?;
        a.run("vba.run", json!({}))?;
        pump(&mut a);
        assert!(a.dialog.is_some());
        match action {
            "cancel" => {
                control(&mut a, "ui.dialog.cancel", json!({}))?;
            }
            "reset" => {
                a.run("vba.reset", json!({}))?;
            }
            _ => {
                a.run("vba.close", json!({"discard":true}))?;
            }
        }
        pump(&mut a);
        assert!(a.dialog.is_none());
        let value = a.session.doc()?.wb.sheet(0).ok_or("sheet")?.value(CellRef::new(0, 0));
        if action == "cancel" {
            complete(&a)?;
            assert_eq!(value, Value::Bool(false));
        } else {
            assert!(value.is_empty());
        }
    }
    Ok(())
}
#[test]
fn filename_pickers_and_status_visibility_use_actual_host_services() -> R {
    let mut a = app(r#"Sub Main()
If Not Application.DisplayStatusBar Then Err.Raise 5
Application.DisplayStatusBar=False
If Application.DisplayStatusBar Then Err.Raise 5
If Application.GetOpenFilename<>False Then Err.Raise 5
If Application.GetSaveAsFilename("report.xlsx")<>"/chosen/report.xlsx" Then Err.Raise 5
End Sub"#)?;
    let calls = Rc::new(RefCell::new(Vec::new()));
    let seen = calls.clone();
    a.services.pick_open = Some(Box::new(move || {
        seen.borrow_mut().push("open".into());
        None
    }));
    let seen = calls.clone();
    a.services.pick_save = Some(Box::new(move |suggested| {
        seen.borrow_mut().push(suggested.to_string());
        Some("/chosen/report.xlsx".into())
    }));
    a.run("vba.run", json!({}))?;
    pump(&mut a);
    complete(&a)?;
    assert_eq!(*calls.borrow(), vec!["open", "report.xlsx"]);
    assert!(!a.ui.status_bar);
    assert!(a.session.doc()?.path.is_none());
    assert_eq!(a.session.documents().len(), 1);
    Ok(())
}
#[test]
fn modal_owner_changes_and_revoked_consent_fail_without_mutating_either_sheet() -> R {
    for revoke in [false, true] {
        let mut a = app("Sub Main()\nx=Application.Dialogs(xlDialogColumnWidth).Show\nRange(\"A1\").Value2=99\nEnd Sub")?;
        a.run("vba.run", json!({}))?;
        pump(&mut a);
        assert!(a.dialog.is_some());
        let old = a.session.doc()?.uid;
        if revoke {
            a.session.doc_mut()?.trust = gridcraft_engine::Trust::FromFile;
            a.vba.as_mut().ok_or("editor")?.bridge.consent(false);
        } else {
            a.session.new_workbook();
        }
        pump(&mut a);
        assert!(a.dialog.is_none());
        assert_eq!(a.vba.as_ref().ok_or("editor")?.ide.mode(), if revoke { Mode::Design } else { Mode::Failed });
        assert!(a.session.document(old).ok_or("owner")?.wb.sheet(0).ok_or("sheet")?.value(CellRef::new(0, 0)).is_empty());
        assert!(a.session.doc()?.wb.sheet(0).ok_or("sheet")?.value(CellRef::new(0, 0)).is_empty());
    }
    Ok(())
}
#[test]
fn existing_dialog_renders_and_keyboard_ok_cancel_complete_the_vba_call() -> R {
    for accepted in [false, true] {
        let mut a = app("Sub Main()\nRange(\"A1\").Value2=Application.Dialogs(xlDialogColumnWidth).Show\nEnd Sub")?;
        a.use_prepared_fonts();
        a.run("vba.run", json!({}))?;
        let screenshot = std::env::var("GRIDCRAFT_OM15_SNAPSHOT").ok();
        let builder = egui_kittest::Harness::builder().with_size(egui::vec2(1100., 760.)).with_context_setup(|ctx| {
            ctx.set_fonts(theme::pinned_font_definitions());
            theme::apply(ctx, false);
        });
        let builder = if screenshot.is_some() { builder.wgpu() } else { builder };
        let mut h = builder.build_ui_state(
            |ui, a: &mut SheetApp| {
                a.logic(ui.ctx());
                a.ui(ui);
            },
            a,
        );
        h.run_steps(15);
        assert!(h.state().dialog.is_some());
        if !accepted && let Some(path) = screenshot {
            h.render()?.save(path)?;
        }
        h.key_press(if accepted { egui::Key::Enter } else { egui::Key::Escape });
        h.run_steps(20);
        complete(h.state())?;
        assert!(h.state().dialog.is_none());
        assert_eq!(h.state().session.doc()?.wb.sheet(0).ok_or("sheet")?.value(CellRef::new(0, 0)), Value::Bool(accepted));
    }
    Ok(())
}
#[test]
fn every_selected_excel_dialog_id_opens_its_existing_native_dialog() -> R {
    for (constant, name) in [
        ("xlDialogColumnWidth", "columnWidth"),
        ("xlDialogRowHeight", "rowHeight"),
        ("xlDialogZoom", "zoom"),
        ("xlDialogInsert", "insertCells"),
        ("xlDialogEditDelete", "deleteCells"),
        ("xlDialogDefineName", "defineName"),
        ("xlDialogInsertHyperlink", "insertLink"),
    ] {
        let mut a = app(&format!("Sub Main()\nIf Application.Dialogs({constant}).Show Then Err.Raise 5\nEnd Sub"))?;
        a.run("vba.run", json!({}))?;
        pump(&mut a);
        assert_eq!(a.dialog.as_ref().ok_or("dialog")?.name(), name);
        control(&mut a, "ui.dialog.cancel", json!({}))?;
        pump(&mut a);
        complete(&a)?;
    }
    Ok(())
}
