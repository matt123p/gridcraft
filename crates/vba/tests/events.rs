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
