#![cfg(feature = "vba")]
use gridcraft_engine::{
    Session,
    core::{CellRef, Value},
};
use gridcraft_ui_egui::{Services, SheetApp};
use gridcraft_vba::vba_ide::{
    form_archive,
    forms::{Bounds, Control, ControlKind, Form},
};
use serde_json::json;
use std::{fs, path::PathBuf};
use vba_ui_egui::{Command, FormCommand};

type R = Result<(), Box<dyn std::error::Error>>;

#[test]
fn imported_tvform_runs_textbox_and_button_callbacks_on_the_workbook_host() -> R {
    let source = "Private Sub TextBox1_Change()\nRange(\"A1\").Value2 = Len(TextBox1.Text)\nEnd Sub\nPrivate Sub CommandButton1_Click()\nRange(\"A2\").Value2 = 42\nMe.Hide\nEnd Sub\n";
    let form = Form {
        name: "M5Form".into(),
        caption: "M5 Form".into(),
        width: 300.0,
        height: 160.0,
        controls: vec![
            Control {
                id: 1,
                name: "TextBox1".into(),
                kind: ControlKind::TextBox,
                bounds: Bounds { left: 12.0, top: 12.0, width: 120.0, height: 24.0 },
                caption: String::new(),
                text: String::new(),
                enabled: true,
                visible: true,
                tab_stop: true,
                group: None,
            },
            Control {
                id: 2,
                name: "CommandButton1".into(),
                kind: ControlKind::CommandButton,
                bounds: Bounds { left: 12.0, top: 48.0, width: 120.0, height: 24.0 },
                caption: "Apply".into(),
                text: String::new(),
                enabled: true,
                visible: true,
                tab_stop: true,
                group: None,
            },
        ],
        tab_order: vec![1, 2],
    };
    let archive = form_archive::encode(&form, source)?;
    let path = PathBuf::from(std::env::temp_dir()).join(format!("gridcraft-m5-form-{}.tvform", std::process::id()));
    fs::write(&path, archive.as_bytes())?;
    let picked = path.clone();
    let services = Services { pick_open: Some(Box::new(move || Some(picked.to_string_lossy().into_owned()))), ..Default::default() };
    let mut session = Session::new();
    session.new_workbook();
    let mut app = SheetApp::new(session, services);
    app.run("vba.open", json!({})).map_err(std::io::Error::other)?;
    let import = app.run("vba.importSource", json!({})).map_err(std::io::Error::other)?;
    assert_eq!(import, json!({"imported":true}));

    let form_id = app
        .vba
        .as_ref()
        .and_then(|editor| editor.ide.modules.iter().find(|module| editor.ide.form(module.id).is_some()))
        .map(|module| module.id)
        .ok_or("imported form missing from embedded IDE")?;
    assert_eq!(app.vba.as_ref().and_then(|editor| editor.ide.form(form_id)).map(|designer| designer.form().name.as_str()), Some("M5Form"));

    dispatch_form(&mut app, FormCommand::Run(form_id, true))?;
    settle(&mut app)?;
    let editor = app.vba.as_ref().ok_or("editor closed")?;
    assert_eq!(editor.ide.mode(), gridcraft_vba::Mode::Waiting, "{:?}", editor.ide.diagnostics);
    assert!(editor.ide.form_visible(form_id), "{:?}", editor.ide.diagnostics);
    assert!(editor.ide.form_is_modal(form_id), "{:?}", editor.ide.diagnostics);
    dispatch_form(&mut app, FormCommand::Event(form_id, Some(1), Some("hello".into())))?;
    settle(&mut app)?;
    assert_eq!(app.session.doc()?.wb.sheet(0).ok_or("sheet missing")?.value(CellRef::new(0, 0)), Value::Number(5.0));
    dispatch_form(&mut app, FormCommand::Event(form_id, Some(2), None))?;
    settle(&mut app)?;
    for _ in 0..20 {
        let editor = app.vba.as_mut().ok_or("editor closed")?;
        if editor.ide.mode() == gridcraft_vba::Mode::Completed {
            break;
        }
        editor.bridge.tick(&mut app.session, &mut editor.ide)?;
    }
    assert_eq!(app.session.doc()?.wb.sheet(0).ok_or("sheet missing")?.value(CellRef::new(1, 0)), Value::Number(42.0));
    let editor = app.vba.as_ref().ok_or("editor closed")?;
    assert!(!editor.ide.form_visible(form_id), "{:?}", editor.ide.diagnostics);
    assert!(!editor.ide.form_is_modal(form_id), "{:?}", editor.ide.diagnostics);
    assert_eq!(app.vba.as_ref().map(|editor| editor.ide.mode()), Some(gridcraft_vba::Mode::Completed));
    let editor = app.vba.as_mut().ok_or("editor closed")?;
    editor.bridge.finish(&mut app.session)?;
    assert!(!app.session.in_group(editor.bridge.document()));

    app.run("vba.reset", json!({})).map_err(std::io::Error::other)?;
    dispatch_form(&mut app, FormCommand::Run(form_id, true))?;
    settle(&mut app)?;
    assert!(app.vba.as_ref().is_some_and(|editor| editor.ide.form_is_modal(form_id)));
    dispatch_form(&mut app, FormCommand::Event(form_id, None, None))?;
    settle(&mut app)?;
    for _ in 0..20 {
        let editor = app.vba.as_mut().ok_or("editor closed")?;
        if editor.ide.mode() == gridcraft_vba::Mode::Completed {
            break;
        }
        editor.bridge.tick(&mut app.session, &mut editor.ide)?;
    }
    let editor = app.vba.as_ref().ok_or("editor closed")?;
    assert_eq!(editor.ide.mode(), gridcraft_vba::Mode::Completed, "{:?}", editor.ide.diagnostics);
    assert!(!editor.ide.form_visible(form_id));
    assert!(!editor.ide.form_is_modal(form_id));

    fs::remove_file(path)?;
    Ok(())
}

fn dispatch_form(app: &mut SheetApp, command: FormCommand) -> R {
    let editor = app.vba.as_mut().ok_or("editor closed")?;
    let mut host = editor.bridge.host(&mut app.session);
    vba_ui_egui::apply(&mut editor.ide, Command::Form(command), &mut host)?;
    Ok(())
}

fn settle(app: &mut SheetApp) -> R {
    for _ in 0..1000 {
        let editor = app.vba.as_mut().ok_or("editor closed")?;
        if editor.ide.mode() != gridcraft_vba::Mode::Running {
            return Ok(());
        }
        editor.bridge.tick(&mut app.session, &mut editor.ide)?;
    }
    Err("form execution exceeded the acceptance tick budget".into())
}
