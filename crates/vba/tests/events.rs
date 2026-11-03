//! Independently authored Excel event callbacks against live GridCraft workbook state.
use gridcraft_engine::{
    Session,
    core::{CellRef, RangeRef, Value},
};
use gridcraft_vba::{Bridge, Ide, Mode};
use serde_json::json;
type R = Result<(), Box<dyn std::error::Error>>;

#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn worksheet_change_uses_document_owner_from_design_mode() -> R {
    let mut session = Session::new();
    session.new_workbook();
    session.execute("home.insertSheet", json!({"name":"Other"}))?;
    {
        let document = session.doc_mut()?;
        let workbook = std::sync::Arc::make_mut(&mut document.wb);
        workbook.sheet_mut(0).ok_or("owner sheet missing")?.code_name = Some("SheetCode".into());
        workbook.set_active_sheet(1);
    }
    let mut ide = Ide::new();
    ide.add_document_module(
        "Sheet1.cls",
        "SheetCode",
        "Attribute VB_Name = \"SheetCode\"\nPrivate Sub Worksheet_Change(ByVal Target As Range)\nApplication.EnableEvents = False\nMe.Range(\"B1\").Value2 = Target.Row\nEnd Sub\n",
    )?;
    let mut bridge = Bridge::new(&session)?;
    bridge.consent(true);
    bridge.configure(&mut ide)?;
    bridge.notify_sheet_change(&session, 0, RangeRef::cell(CellRef::new(4, 3)))?;
    for _ in 0..100 {
        bridge.tick(&mut session, &mut ide)?;
        if matches!(ide.mode(), Mode::Completed) {
            break;
        }
    }
    assert_eq!(session.doc()?.wb.sheet(0).ok_or("owner sheet missing")?.value(CellRef::new(0, 1)), Value::Number(5.0));
    assert_eq!(session.doc()?.wb.sheet(1).ok_or("other sheet missing")?.value(CellRef::new(0, 1)), Value::Empty);
    Ok(())
}

#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn workbook_open_requires_explicit_trusted_queueing() -> R {
    let mut session = Session::new();
    session.new_workbook();
    {
        let document = session.doc_mut()?;
        std::sync::Arc::make_mut(&mut document.wb).code_name = Some("BookCode".into());
    }
    let mut ide = Ide::new();
    ide.add_document_module(
        "ThisWorkbook.cls",
        "BookCode",
        "Attribute VB_Name = \"BookCode\"\nPrivate Sub Workbook_Open()\nDebug.Print \"opened\"\nEnd Sub\n",
    )?;
    let mut bridge = Bridge::new(&session)?;
    bridge.consent(true);
    bridge.configure(&mut ide)?;
    assert!(ide.output().is_empty());
    bridge.queue_workbook_open(&session)?;
    for _ in 0..100 {
        bridge.tick(&mut session, &mut ide)?;
        if matches!(ide.mode(), Mode::Completed) {
            break;
        }
    }
    assert_eq!(ide.output(), &["opened"]);
    Ok(())
}

#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn worksheet_event_queue_is_fifo_and_disabling_discards_only_the_pending_tail() -> R {
    let mut session = Session::new();
    session.new_workbook();
    {
        let document = session.doc_mut()?;
        std::sync::Arc::make_mut(&mut document.wb).sheet_mut(0).ok_or("sheet missing")?.code_name = Some("SheetCode".into());
    }
    let mut ide = Ide::new();
    ide.add_document_module(
        "Sheet1.cls",
        "SheetCode",
        "Attribute VB_Name = \"SheetCode\"\nPrivate Sub Worksheet_Change(ByVal Target As Range)\nDebug.Print Target.Row\nEnd Sub\n",
    )?;
    let mut bridge = Bridge::new(&session)?;
    bridge.consent(true);
    bridge.configure(&mut ide)?;
    bridge.notify_sheet_change(&session, 0, RangeRef::cell(CellRef::new(4, 3)))?;
    bridge.notify_sheet_change(&session, 0, RangeRef::cell(CellRef::new(6, 2)))?;
    for _ in 0..100 {
        bridge.tick(&mut session, &mut ide)?;
        if matches!(ide.mode(), Mode::Completed) && ide.output().len() >= 2 {
            break;
        }
    }
    assert_eq!(ide.output(), &["5", "7"]);

    let mut suppressed = Ide::new();
    suppressed.add_document_module(
        "Sheet1.cls",
        "SheetCode",
        "Attribute VB_Name = \"SheetCode\"\nPrivate Sub Worksheet_Change(ByVal Target As Range)\nApplication.EnableEvents = False\nDebug.Print Target.Row\nApplication.EnableEvents = True\nEnd Sub\n",
    )?;
    bridge.configure(&mut suppressed)?;
    bridge.notify_sheet_change(&session, 0, RangeRef::cell(CellRef::new(4, 3)))?;
    bridge.notify_sheet_change(&session, 0, RangeRef::cell(CellRef::new(6, 2)))?;
    for _ in 0..100 {
        bridge.tick(&mut session, &mut suppressed)?;
        if matches!(suppressed.mode(), Mode::Completed) && !suppressed.output().is_empty() {
            break;
        }
    }
    assert_eq!(suppressed.output(), &["5"]);
    bridge.notify_sheet_change(&session, 0, RangeRef::cell(CellRef::new(8, 0)))?;
    for _ in 0..100 {
        bridge.tick(&mut session, &mut suppressed)?;
        if matches!(suppressed.mode(), Mode::Completed) && suppressed.output().len() >= 2 {
            break;
        }
    }
    assert_eq!(suppressed.output(), &["5", "9"]);
    Ok(())
}

#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn worksheet_event_breakpoint_resumes_once_without_replaying_delivery() -> R {
    let mut session = Session::new();
    session.new_workbook();
    {
        let document = session.doc_mut()?;
        std::sync::Arc::make_mut(&mut document.wb).sheet_mut(0).ok_or("sheet missing")?.code_name = Some("SheetCode".into());
    }
    let mut ide = Ide::new();
    let source = "Attribute VB_Name = \"SheetCode\"\nPrivate Sub Worksheet_Change(ByVal Target As Range)\nDebug.Print Target.Row\nEnd Sub\n";
    ide.add_document_module("Sheet1.cls", "SheetCode", source)?;
    let mut bridge = Bridge::new(&session)?;
    bridge.consent(true);
    bridge.configure(&mut ide)?;
    let breakpoint = u32::try_from(source.find("Debug.Print").ok_or("handler statement missing")?)?;
    ide.modules.first_mut().ok_or("document module missing")?.breakpoints.insert(breakpoint);
    bridge.notify_sheet_change(&session, 0, RangeRef::cell(CellRef::new(4, 3)))?;
    for _ in 0..100 {
        bridge.tick(&mut session, &mut ide)?;
        if ide.mode() == Mode::Break {
            break;
        }
    }
    assert_eq!(ide.mode(), Mode::Break);
    assert!(ide.output().is_empty());
    bridge.dispatch(&mut session, &mut ide, gridcraft_vba::Action::Run)?;
    for _ in 0..100 {
        bridge.tick(&mut session, &mut ide)?;
        if ide.mode() == Mode::Completed {
            break;
        }
    }
    assert_eq!(ide.output(), &["5"]);
    assert_eq!(ide.mode(), Mode::Completed);
    Ok(())
}

#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn reset_close_and_trust_revocation_cancel_queued_events() -> R {
    for lifecycle in ["reset", "close", "revoke"] {
        let mut session = Session::new();
        session.new_workbook();
        {
            let document = session.doc_mut()?;
            std::sync::Arc::make_mut(&mut document.wb).sheet_mut(0).ok_or("sheet missing")?.code_name = Some("SheetCode".into());
        }
        let mut ide = Ide::new();
        let source = "Attribute VB_Name = \"SheetCode\"\nPrivate Sub Worksheet_Change(ByVal Target As Range)\nDebug.Print Target.Row\nEnd Sub\n";
        ide.add_document_module("Sheet1.cls", "SheetCode", source)?;
        let breakpoint = u32::try_from(source.find("Debug.Print").ok_or("handler statement missing")?)?;
        ide.modules.first_mut().ok_or("document module missing")?.breakpoints.insert(breakpoint);
        let mut bridge = Bridge::new(&session)?;
        bridge.consent(true);
        bridge.configure(&mut ide)?;
        bridge.notify_sheet_change(&session, 0, RangeRef::cell(CellRef::new(4, 3)))?;
        bridge.notify_sheet_change(&session, 0, RangeRef::cell(CellRef::new(5, 3)))?;
        for _ in 0..10 {
            bridge.tick(&mut session, &mut ide)?;
            if ide.mode() == Mode::Break {
                break;
            }
        }
        assert_eq!(ide.mode(), Mode::Break);
        match lifecycle {
            "reset" => bridge.dispatch(&mut session, &mut ide, gridcraft_vba::Action::Reset)?,
            "close" => bridge.close(&mut session, &mut ide, true)?,
            _ => {
                session.doc_mut()?.trust = gridcraft_engine::Trust::FromFile;
                bridge.consent(false);
                assert!(bridge.tick(&mut session, &mut ide).is_err());
            }
        }
        for _ in 0..5 {
            if bridge.tick(&mut session, &mut ide).is_err() {
                break;
            }
        }
        assert!(ide.output().is_empty(), "queued event ran after {lifecycle}");
    }
    Ok(())
}

#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn event_chain_budget_rejects_overflow_and_resets_after_drain() -> R {
    let mut session = Session::new();
    session.new_workbook();
    {
        let document = session.doc_mut()?;
        std::sync::Arc::make_mut(&mut document.wb).sheet_mut(0).ok_or("sheet missing")?.code_name = Some("SheetCode".into());
    }
    let mut ide = Ide::new();
    ide.add_document_module(
        "Sheet1.cls",
        "SheetCode",
        "Attribute VB_Name = \"SheetCode\"\nPrivate Sub Worksheet_Change(ByVal Target As Range)\nIf Target.Row = 257 Then Debug.Print Target.Row\nEnd Sub\n",
    )?;
    let mut bridge = Bridge::new(&session)?;
    bridge.consent(true);
    bridge.configure(&mut ide)?;
    for row in 0..256 {
        bridge.notify_sheet_change(&session, 0, RangeRef::cell(CellRef::new(row, 0)))?;
    }
    assert!(bridge.notify_sheet_change(&session, 0, RangeRef::cell(CellRef::new(256, 0))).is_err());
    for _ in 0..5000 {
        bridge.tick(&mut session, &mut ide)?;
        if ide.output().len() == 256 && ide.mode() == Mode::Completed {
            break;
        }
    }
    assert!(ide.output().is_empty());
    assert!(bridge.current(&session));
    assert!(bridge.permitted(&session));
    bridge.notify_sheet_change(&session, 0, RangeRef::cell(CellRef::new(256, 0)))?;
    for _ in 0..100 {
        bridge.tick(&mut session, &mut ide)?;
        if ide.output().last().is_some_and(|line| line == "257") && ide.mode() == Mode::Completed {
            break;
        }
    }
    assert_eq!(ide.output().last().map(String::as_str), Some("257"));
    Ok(())
}

#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn vba_range_write_refuses_before_mutation_when_the_event_queue_is_full() -> R {
    let mut session = Session::new();
    session.new_workbook();
    {
        let document = session.doc_mut()?;
        std::sync::Arc::make_mut(&mut document.wb).sheet_mut(0).ok_or("sheet missing")?.code_name = Some("SheetCode".into());
    }
    let mut ide = Ide::new();
    ide.add_module("Main.bas", "Public Sub Main()\nRange(\"Z1\").Value2 = 40\nEnd Sub\n")?;
    ide.add_document_module(
        "Sheet1.cls",
        "SheetCode",
        "Attribute VB_Name = \"SheetCode\"\nPrivate Sub Worksheet_Change(ByVal Target As Range)\nDebug.Print Target.Row\nEnd Sub\n",
    )?;
    let mut bridge = Bridge::new(&session)?;
    bridge.consent(true);
    bridge.configure(&mut ide)?;
    for row in 0..256 {
        bridge.notify_sheet_change(&session, 0, RangeRef::cell(CellRef::new(row, 0)))?;
    }
    bridge.dispatch(&mut session, &mut ide, gridcraft_vba::Action::Run)?;
    for _ in 0..100 {
        if bridge.tick(&mut session, &mut ide).is_err() || ide.mode() == Mode::Failed {
            break;
        }
    }
    assert_eq!(session.doc()?.wb.sheet(0).ok_or("sheet missing")?.value(CellRef::parse("Z1").ok_or("Z1")?), Value::Empty);
    Ok(())
}

#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn failed_event_cancels_the_tail_and_keeps_prior_writes_undoable() -> R {
    let mut session = Session::new();
    session.new_workbook();
    session.execute("home.insertSheet", serde_json::json!({"name":"Other"}))?;
    {
        let document = session.doc_mut()?;
        std::sync::Arc::make_mut(&mut document.wb).sheet_mut(0).ok_or("sheet missing")?.code_name = Some("SheetCode".into());
    }
    let mut ide = Ide::new();
    ide.add_document_module(
        "Sheet1.cls",
        "SheetCode",
        "Attribute VB_Name = \"SheetCode\"\nPrivate Sub Worksheet_Change(ByVal Target As Range)\nIf Target.Row = 5 Then ThisWorkbook.Worksheets(2).Range(\"B1\").Value2 = 42\nIf Target.Row = 7 Then Err.Raise 5\nEnd Sub\n",
    )?;
    let mut bridge = Bridge::new(&session)?;
    bridge.consent(true);
    bridge.configure(&mut ide)?;
    bridge.notify_sheet_change(&session, 0, RangeRef::cell(CellRef::new(4, 3)))?;
    bridge.notify_sheet_change(&session, 0, RangeRef::cell(CellRef::new(6, 2)))?;
    let mut failed = false;
    for _ in 0..100 {
        if bridge.tick(&mut session, &mut ide).is_err() || ide.mode() == Mode::Failed {
            failed = true;
            break;
        }
    }
    assert!(failed);
    assert_eq!(session.doc()?.wb.sheet(1).ok_or("other sheet missing")?.value(CellRef::new(0, 1)), Value::Number(42.0));
    assert!(!session.in_group(bridge.document()));
    session.execute("edit.undo", serde_json::json!({}))?;
    assert_eq!(session.doc()?.wb.sheet(1).ok_or("other sheet missing")?.value(CellRef::new(0, 1)), Value::Empty);
    Ok(())
}

fn host_call(
    bridge: &mut Bridge,
    session: &mut Session,
    target: gridcraft_vba::ObjectRef,
    kind: vba_foreign::InvokeKind,
    member: &str,
    arguments: Vec<gridcraft_vba::Value>,
) -> vba_foreign::Reply {
    use gridcraft_vba::Host;
    bridge.host(session).invoke(vba_foreign::HostCall {
        target,
        kind,
        member: member.into(),
        arguments: arguments.into_iter().map(|value| vba_foreign::HostArgument { name: None, value, by_ref: false }).collect(),
    })
}

fn event_object(reply: vba_foreign::Reply) -> gridcraft_vba::ObjectRef {
    match reply {
        vba_foreign::Reply::Complete(completion) => match completion.value {
            gridcraft_vba::Value::Object(reference) => reference,
            other => panic!("expected object, got {other:?}"),
        },
        other => panic!("expected completed call, got {other:?}"),
    }
}

fn event_text(text: &str) -> gridcraft_vba::Value {
    gridcraft_vba::Value::String(vba_values::VbaString::from_utf8(text).unwrap())
}

fn event_setup(session: &mut Session) -> Result<(Bridge, Ide, gridcraft_vba::ObjectRef), Box<dyn std::error::Error>> {
    event_setup_with_callback(session, "")
}

fn event_setup_with_callback(session: &mut Session, callback: &str) -> Result<(Bridge, Ide, gridcraft_vba::ObjectRef), Box<dyn std::error::Error>> {
    use gridcraft_vba::Host;
    {
        let document = session.doc_mut()?;
        let workbook = std::sync::Arc::make_mut(&mut document.wb);
        for (index, sheet) in workbook.sheets.iter_mut().enumerate() {
            std::sync::Arc::make_mut(sheet).code_name = Some(format!("SheetCode{index}"));
        }
    }
    let mut ide = Ide::new();
    for index in 0..session.doc()?.wb.sheets.len() {
        ide.add_document_module(
            &format!("Sheet{index}.cls"),
            &format!("SheetCode{index}"),
            &format!("Attribute VB_Name = \"SheetCode{index}\"\nPrivate Sub Worksheet_Change(ByVal Target As Range)\nDebug.Print \"sheet{index}:\" & Target.Address\n{callback}\nEnd Sub\n"),
        )?;
    }
    let mut bridge = Bridge::new(session)?;
    bridge.consent(true);
    bridge.configure(&mut ide)?;
    let application = gridcraft_vba::ObjectRef { host: bridge.host(session).identity(), generation: bridge.document(), object: 0 };
    Ok((bridge, ide, application))
}

fn event_range(bridge: &mut Bridge, session: &mut Session, application: gridcraft_vba::ObjectRef, address: &str) -> gridcraft_vba::ObjectRef {
    event_object(host_call(bridge, session, application, vba_foreign::InvokeKind::Get, "Range", vec![event_text(address)]))
}

fn drain_events(bridge: &mut Bridge, session: &mut Session, ide: &mut Ide) -> R {
    for _ in 0..1000 {
        bridge.tick(session, ide)?;
        assert_ne!(ide.mode(), Mode::Failed, "{:?}", ide.diagnostics);
    }
    Ok(())
}

#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn inactive_workbook_write_delivers_one_event_to_its_owner() -> R {
    use vba_foreign::{InvokeKind, Reply};
    for active_code_name in [None, Some("SheetCode0")] {
        let mut session = Session::new();
        session.new_workbook();
        let (mut bridge, mut ide, application) =
            event_setup_with_callback(&mut session, "Application.EnableEvents = False\nMe.Range(\"C1\").Value2 = Target.Row")?;
        let owner = bridge.document();
        let range = event_range(&mut bridge, &mut session, application, "B3");
        session.new_workbook();
        let active = session.doc()?.uid;
        std::sync::Arc::make_mut(&mut session.doc_mut()?.wb).sheet_mut(0).ok_or("sheet missing")?.code_name = active_code_name.map(str::to_string);
        assert!(matches!(
            host_call(&mut bridge, &mut session, range, InvokeKind::Let, "Value2", vec![gridcraft_vba::Value::Long(42)]),
            Reply::Complete(_)
        ));
        drain_events(&mut bridge, &mut session, &mut ide)?;
        assert_eq!(ide.output(), &["sheet0:$B$3"]);
        assert_eq!(session.doc()?.uid, active);
        assert_eq!(session.doc()?.wb.sheets[0].value(CellRef::new(2, 1)), Value::Empty);
        assert_eq!(session.doc()?.wb.sheets[0].value(CellRef::new(0, 2)), Value::Empty);
        assert_eq!(session.document(owner).ok_or("owner missing")?.wb.sheets[0].value(CellRef::new(0, 2)), Value::Number(3.0));
        assert_eq!(session.document(owner).ok_or("owner missing")?.wb.sheets[0].value(CellRef::new(2, 1)), Value::Number(42.0));
        session.with_target(gridcraft_engine::Target::Doc(owner), |s| s.execute("edit.undo", json!({})))??;
        assert_eq!(session.document(owner).ok_or("owner missing")?.wb.sheets[0].value(CellRef::new(2, 1)), Value::Empty);
    }
    Ok(())
}

#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn foreign_workbook_with_same_code_name_does_not_run_owner_handler() -> R {
    use vba_foreign::{InvokeKind, Reply};
    let mut session = Session::new();
    session.new_workbook();
    let (mut bridge, mut ide, application) = event_setup(&mut session)?;
    session.new_workbook();
    std::sync::Arc::make_mut(&mut session.doc_mut()?.wb).sheet_mut(0).ok_or("sheet missing")?.code_name = Some("SheetCode0".into());
    let range = event_range(&mut bridge, &mut session, application, "B3");
    assert!(matches!(
        host_call(&mut bridge, &mut session, range, InvokeKind::Let, "Value2", vec![gridcraft_vba::Value::Long(42)]),
        Reply::Complete(_)
    ));
    drain_events(&mut bridge, &mut session, &mut ide)?;
    assert!(ide.output().is_empty());
    assert_eq!(session.doc()?.wb.sheets[0].value(CellRef::new(2, 1)), Value::Number(42.0));
    Ok(())
}

#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn assignment_forms_and_clear_contents_each_deliver_exactly_one_change() -> R {
    let mut session = Session::new();
    session.new_workbook();
    let (mut bridge, mut ide, _) = event_setup(&mut session)?;
    ide.add_module(
        "Main.bas",
        "Sub Main()\nRange(\"A1\").Value2 = 1\nRange(\"A2\") = 2\nRange(\"A3:A4\")(2, 1) = 3\nRange(\"A1\").ClearContents\nEnd Sub\n",
    )?;
    bridge.dispatch(&mut session, &mut ide, gridcraft_vba::Action::Run)?;
    drain_events(&mut bridge, &mut session, &mut ide)?;
    assert_eq!(ide.output(), &["sheet0:$A$1", "sheet0:$A$2", "sheet0:$A$4", "sheet0:$A$1"]);
    Ok(())
}

#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn two_sheet_cut_reserves_both_events_before_mutation() -> R {
    use vba_foreign::{InvokeKind, Reply};
    let mut session = Session::new();
    session.new_workbook();
    session.execute("home.insertSheet", json!({"name":"Other"}))?;
    session.execute("cell.set", json!({"cell":"A1","input":"17"}))?;
    let (mut bridge, _ide, application) = event_setup(&mut session)?;
    let source = event_range(&mut bridge, &mut session, application, "A1");
    let worksheet =
        event_object(host_call(&mut bridge, &mut session, application, InvokeKind::Get, "Worksheets", vec![gridcraft_vba::Value::Long(1)]));
    let destination = event_range(&mut bridge, &mut session, worksheet, "B1");
    for row in 0..255 {
        bridge.notify_sheet_change(&session, 0, RangeRef::cell(CellRef::new(row, 4)))?;
    }
    let before = session.doc()?.wb.clone();
    let selection = session.doc()?.selection.clone();
    let clipboard_was_empty = session.clipboard.is_none();
    let reply = host_call(&mut bridge, &mut session, source, InvokeKind::Call, "Cut", vec![gridcraft_vba::Value::Object(destination)]);
    assert!(matches!(&reply, Reply::Error(error) if error.number == 28), "{reply:?}");
    assert_eq!(*session.doc()?.wb, *before);
    assert_eq!(session.doc()?.selection, selection);
    assert_eq!(session.clipboard.is_none(), clipboard_was_empty);
    bridge.finish(&mut session)?;
    Ok(())
}

#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn copy_refuses_before_mutation_when_event_target_handle_cannot_be_allocated() -> R {
    use vba_foreign::{InvokeKind, Reply};
    let mut session = Session::new();
    session.new_workbook();
    session.execute("cell.set", json!({"cell":"A1","input":"17"}))?;
    let (mut bridge, _ide, application) = event_setup(&mut session)?;
    let source = event_range(&mut bridge, &mut session, application, "A1:B1");
    let destination = event_range(&mut bridge, &mut session, application, "D1");
    let mut exhausted = false;
    for row in 2..=gridcraft_vba::MAX_OBJECTS + 2 {
        let reply = host_call(&mut bridge, &mut session, application, InvokeKind::Get, "Range", vec![event_text(&format!("A{row}"))]);
        if let Reply::Error(error) = reply {
            assert_eq!(error.number, 7);
            exhausted = true;
            break;
        }
    }
    assert!(exhausted);
    let before = session.doc()?.wb.clone();
    let selection = session.doc()?.selection.clone();
    let reply = host_call(&mut bridge, &mut session, source, InvokeKind::Call, "Copy", vec![gridcraft_vba::Value::Object(destination)]);
    assert!(matches!(&reply, Reply::Error(error) if error.number == 7), "{reply:?}");
    assert_eq!(*session.doc()?.wb, *before);
    assert_eq!(session.doc()?.selection, selection);
    assert!(session.clipboard.is_none());
    bridge.finish(&mut session)?;
    Ok(())
}

#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn callback_default_assignment_appends_one_followup_event() -> R {
    use vba_foreign::{InvokeKind, Reply};
    let mut session = Session::new();
    session.new_workbook();
    let (mut bridge, mut ide, application) = event_setup_with_callback(&mut session, "If Target.Row = 1 Then Me.Range(\"A2\") = 9")?;
    let range = event_range(&mut bridge, &mut session, application, "A1");
    assert!(matches!(
        host_call(&mut bridge, &mut session, range, InvokeKind::DefaultLet, "", vec![gridcraft_vba::Value::Long(1)]),
        Reply::Complete(_)
    ));
    drain_events(&mut bridge, &mut session, &mut ide)?;
    assert_eq!(ide.output(), &["sheet0:$A$1", "sheet0:$A$2"]);
    assert_eq!(session.doc()?.wb.sheets[0].value(CellRef::new(1, 0)), Value::Number(9.0));
    assert!(!session.in_group(bridge.document()));
    session.execute("edit.undo", json!({}))?;
    assert_eq!(session.doc()?.wb.sheets[0].value(CellRef::new(0, 0)), Value::Empty);
    assert_eq!(session.doc()?.wb.sheets[0].value(CellRef::new(1, 0)), Value::Empty);
    Ok(())
}
