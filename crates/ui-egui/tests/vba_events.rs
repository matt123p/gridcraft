#![cfg(feature = "vba")]
use gridcraft_engine::{
    Session,
    core::{CellRef, Value},
};
use gridcraft_ui_egui::SheetApp;
use gridcraft_ui_egui::control::{Outcome, handle};
use serde_json::json;

#[test]
fn committed_native_cell_edit_queues_worksheet_change() -> Result<(), Box<dyn std::error::Error>> {
    let mut session = Session::new();
    session.new_workbook();
    {
        let document = session.doc_mut()?;
        std::sync::Arc::make_mut(&mut document.wb).sheet_mut(0).ok_or("sheet missing")?.code_name = Some("SheetCode".into());
    }
    let mut app = SheetApp::new(session, Default::default());
    app.run("vba.open", json!({}))?;
    let editor = app.vba.as_mut().ok_or("editor not opened")?;
    editor.ide.add_document_module(
        "Sheet1.cls",
        "SheetCode",
        "Attribute VB_Name = \"SheetCode\"\nPrivate Sub Worksheet_Change(ByVal Target As Range)\nApplication.EnableEvents = False\nMe.Range(\"A1\").Value2 = Target.Row\nEnd Sub\n",
    )?;
    editor.bridge.configure(&mut editor.ide)?;
    editor.bridge.consent(true);

    app.run("cell.set", json!({"cell":"D5","input":"1"}))?;
    assert_eq!(app.session.doc()?.wb.sheet(0).ok_or("sheet missing")?.value(CellRef::new(4, 3)), Value::Number(1.0));
    assert_eq!(app.session.doc()?.wb.sheet(0).ok_or("sheet missing")?.value(CellRef::default()), Value::Empty);
    for _ in 0..100 {
        let editor = app.vba.as_mut().ok_or("editor closed")?;
        editor.bridge.tick(&mut app.session, &mut editor.ide)?;
        if editor.ide.mode() == gridcraft_vba::Mode::Completed {
            break;
        }
    }
    assert_eq!(app.session.doc()?.wb.sheet(0).ok_or("sheet missing")?.value(CellRef::default()), Value::Number(5.0));
    Ok(())
}

#[test]
fn direct_control_cell_edit_queues_worksheet_change() -> Result<(), Box<dyn std::error::Error>> {
    let mut session = Session::new();
    session.new_workbook();
    {
        let document = session.doc_mut()?;
        std::sync::Arc::make_mut(&mut document.wb).sheet_mut(0).ok_or("sheet missing")?.code_name = Some("SheetCode".into());
    }
    let mut app = SheetApp::new(session, Default::default());
    app.run("vba.open", json!({}))?;
    let editor = app.vba.as_mut().ok_or("editor not opened")?;
    editor.ide.add_document_module(
        "Sheet1.cls",
        "SheetCode",
        "Attribute VB_Name = \"SheetCode\"\nPrivate Sub Worksheet_Change(ByVal Target As Range)\nDebug.Print Target.Address\nEnd Sub\n",
    )?;
    editor.bridge.configure(&mut editor.ide)?;
    editor.bridge.consent(true);

    let response = handle(&mut app, &egui::Context::default(), "engine.execute", &json!({"command":"cell.set","params":{"cell":"B3","input":"4"}}));
    assert!(matches!(response, Outcome::Done(value) if value["ok"] == true));
    for _ in 0..100 {
        let editor = app.vba.as_mut().ok_or("editor closed")?;
        editor.bridge.tick(&mut app.session, &mut editor.ide)?;
        if editor.ide.output().len() == 1 {
            break;
        }
    }
    assert_eq!(app.vba.as_ref().ok_or("editor closed")?.ide.output().last().map(String::as_str), Some("$B$3"));
    Ok(())
}

#[test]
fn committed_cell_editor_edit_queues_worksheet_change() -> Result<(), Box<dyn std::error::Error>> {
    let mut session = Session::new();
    session.new_workbook();
    {
        let document = session.doc_mut()?;
        std::sync::Arc::make_mut(&mut document.wb).sheet_mut(0).ok_or("sheet missing")?.code_name = Some("SheetCode".into());
    }
    let mut app = SheetApp::new(session, Default::default());
    app.run("vba.open", json!({}))?;
    let editor = app.vba.as_mut().ok_or("editor not opened")?;
    editor.ide.add_document_module(
        "Sheet1.cls",
        "SheetCode",
        "Attribute VB_Name = \"SheetCode\"\nPrivate Sub Worksheet_Change(ByVal Target As Range)\nDebug.Print Target.Address\nEnd Sub\n",
    )?;
    editor.bridge.configure(&mut editor.ide)?;
    editor.bridge.consent(true);

    app.begin_edit(Some("7".into()), false);
    assert!(app.commit_edit(0, 0, false, false));
    assert_eq!(app.session.doc()?.wb.sheet(0).ok_or("sheet missing")?.value(CellRef::default()), Value::Number(7.0));
    for _ in 0..100 {
        let editor = app.vba.as_mut().ok_or("editor closed")?;
        editor.bridge.tick(&mut app.session, &mut editor.ide)?;
        if editor.ide.output().len() == 1 {
            break;
        }
    }
    let output = app.vba.as_ref().ok_or("editor closed")?.ide.output();
    assert_eq!(output.last().map(String::as_str), Some("$A$1"));
    Ok(())
}

#[test]
fn native_edit_stays_committed_when_the_event_queue_is_full() -> Result<(), Box<dyn std::error::Error>> {
    let mut session = Session::new();
    session.new_workbook();
    {
        let document = session.doc_mut()?;
        std::sync::Arc::make_mut(&mut document.wb).sheet_mut(0).ok_or("sheet missing")?.code_name = Some("SheetCode".into());
    }
    let mut app = SheetApp::new(session, Default::default());
    app.run("vba.open", json!({}))?;
    let editor = app.vba.as_mut().ok_or("editor not opened")?;
    editor.ide.add_document_module(
        "Sheet1.cls",
        "SheetCode",
        "Attribute VB_Name = \"SheetCode\"\nPrivate Sub Worksheet_Change(ByVal Target As Range)\nDebug.Print Target.Row\nEnd Sub\n",
    )?;
    editor.bridge.configure(&mut editor.ide)?;
    editor.bridge.consent(true);
    for row in 0..256 {
        editor.bridge.notify_sheet_change(&app.session, 0, gridcraft_engine::core::RangeRef::cell(CellRef::new(row, 0)))?;
    }

    app.run("cell.set", json!({"cell":"D5","input":"8"}))?;
    assert_eq!(app.session.doc()?.wb.sheet(0).ok_or("sheet missing")?.value(CellRef::new(4, 3)), Value::Number(8.0));
    let editor = app.vba.as_ref().ok_or("editor closed")?;
    assert!(editor.ide.messages.last().is_some_and(|message| message.contains("worksheet event budget exceeded")));
    Ok(())
}

#[test]
fn native_bulk_content_commands_notify_their_committed_ranges() -> Result<(), Box<dyn std::error::Error>> {
    let mut session = Session::new();
    session.new_workbook();
    {
        let document = session.doc_mut()?;
        std::sync::Arc::make_mut(&mut document.wb).sheet_mut(0).ok_or("sheet missing")?.code_name = Some("SheetCode".into());
    }
    let mut app = SheetApp::new(session, Default::default());
    app.run("vba.open", json!({}))?;
    {
        let editor = app.vba.as_mut().ok_or("editor not opened")?;
        editor.ide.add_document_module(
            "Sheet1.cls",
            "SheetCode",
            "Attribute VB_Name = \"SheetCode\"\nPrivate Sub Worksheet_Change(ByVal Target As Range)\nDebug.Print Target.Address\nEnd Sub\n",
        )?;
        editor.bridge.configure(&mut editor.ide)?;
        editor.bridge.consent(true);
    }
    // Seed source ranges outside SheetApp::run so each assertion observes only its target command.
    app.session.execute("range.setValues", json!({"range":"C1","values":[[3],[4]]}))?;
    app.session.execute("range.setValues", json!({"range":"G1","values":[["x"],["x"]]}))?;
    app.session.execute("range.setValues", json!({"range":"E1","values":[["left,right"]]}))?;
    app.session.execute("range.setValues", json!({"range":"H1","values":[["left"],["right"]]}))?;
    app.session.execute("range.setValues", json!({"range":"K1","values":[[1],[2]]}))?;
    app.session.execute("range.setValues", json!({"range":"L1","values":[[7]]}))?;
    app.session.execute("range.setValues", json!({"range":"M1","values":[[9]]}))?;
    app.session.execute("range.setValues", json!({"range":"O1","values":[[3],[1],[2]]}))?;
    app.session.execute("range.setValues", json!({"range":"Q1","values":[["2024-01-15","15/01/2024"],["2023-12-03",null]]}))?;
    app.session.execute("range.setValues", json!({"range":"U1","values":[[1],[2]]}))?;

    let commands = [
        ("range.setValues", json!({"range":"A1","values":[[1]]}), "$A$1"),
        ("range.fill", json!({"range":"B1:B2","input":"2"}), "$B$1:$B$2"),
        ("edit.clearContents", json!({"range":"A1"}), "$A$1"),
        ("edit.autoFill", json!({"source":"C1:C2","target":"C1:C4"}), "$C$1:$C$4"),
        ("edit.fillSeries", json!({"range":"D1:D3","type":"linear","step":1,"stop":3}), "$D$1:$D$3"),
        ("data.textToColumns", json!({"range":"E1","delimiters":[","]}), "$E$1:$F$1"),
        ("data.removeDuplicates", json!({"range":"G1:G2"}), "$G$1:$G$2"),
        ("edit.cut", json!({"range":"H1:H2"}), "$H$1:$H$2"),
        ("edit.paste", json!({"at":"I1"}), "$I$1:$I$2"),
        ("edit.clearAll", json!({"range":"H1:H2"}), "$H$1:$H$2"),
        ("edit.fillDown", json!({"range":"K1:K3"}), "$K$1:$K$3"),
        ("edit.fillUp", json!({"range":"K1:K3"}), "$K$1:$K$3"),
        ("edit.fillRight", json!({"range":"L1:M1"}), "$L$1:$M$1"),
        ("edit.fillLeft", json!({"range":"L1:M1"}), "$L$1:$M$1"),
        ("data.sortAscending", json!({"range":"O1:O3"}), "$O$1:$O$3"),
        ("data.flashFill", json!({"range":"R2"}), "$R$2"),
    ];
    let mut expected_outputs = 0;
    for (command, params, expected_target) in commands {
        app.run(command, params)?;
        for _ in 0..100 {
            let editor = app.vba.as_mut().ok_or("editor closed")?;
            editor.bridge.tick(&mut app.session, &mut editor.ide)?;
            if editor.ide.output().len() > expected_outputs {
                break;
            }
        }
        let editor = app.vba.as_ref().ok_or("editor closed")?;
        expected_outputs += 1;
        assert_eq!(editor.ide.output().len(), expected_outputs, "{command} did not deliver exactly one event");
        assert_eq!(editor.ide.output().last().map(String::as_str), Some(expected_target), "{command}");
    }
    app.run("data.removeDuplicates", json!({"range":"U1:U2"}))?;
    app.run("home.numberFormat", json!({"range":"A1","format":"0.00"}))?;
    assert_eq!(app.vba.as_ref().ok_or("editor closed")?.ide.output().len(), expected_outputs, "no-op and format-only edits must not notify");
    Ok(())
}
