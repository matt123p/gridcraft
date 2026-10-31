//! M6 acceptance: the specifications' unshadowed module-kind binding and live
//! host-global consultation cases, executed against real GridCraft workbook
//! state through document events. These are GridCraft contract tests, not
//! Microsoft Excel observations.
use gridcraft_engine::{
    Session,
    core::{CellRef, RangeRef, Value},
};
use gridcraft_vba::{Bridge, Ide, Mode};

use serde_json::json;
type R = Result<(), Box<dyn std::error::Error>>;

fn tick_until(b: &mut Bridge, s: &mut Session, ide: &mut Ide, mode: Mode) -> R {
    for _ in 0..200 {
        b.tick(s, ide)?;
        if ide.mode() == mode {
            return Ok(());
        }
    }
    Err("callback did not reach the expected mode".into())
}

#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn worksheet_module_bare_range_targets_the_module_sheet_with_another_sheet_active() -> R {
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
        "Attribute VB_Name = \"SheetCode\"\nPrivate Sub Worksheet_Change(ByVal Target As Range)\nRange(\"A1\").Value2 = \"module-sheet\"\nMe.Range(\"A2\").Value2 = ActiveSheet.Name\nEnd Sub\n",
    )?;
    let mut bridge = Bridge::new(&session)?;
    bridge.consent(true);
    bridge.configure_document(&mut ide, &session.doc()?.wb)?;
    bridge.notify_sheet_change(&session, 0, RangeRef::cell(CellRef::new(4, 3)))?;
    tick_until(&mut bridge, &mut session, &mut ide, Mode::Completed)?;
    let sheet = session.doc()?.wb.sheet(0).ok_or("owner sheet missing")?.clone();
    assert_eq!(sheet.value(CellRef::new(0, 0)), Value::Text("module-sheet".into()));
    assert_eq!(sheet.value(CellRef::new(1, 0)), Value::Text("Other".into()));
    let other = session.doc()?.wb.sheet(1).ok_or("active sheet missing")?.clone();
    assert_eq!(other.value(CellRef::new(0, 0)), Value::Empty);
    assert_eq!(other.name, "Other");
    Ok(())
}

#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn workbook_module_bare_range_targets_the_active_sheet_not_thisworkbook() -> R {
    let mut session = Session::new();
    session.new_workbook();
    session.execute("home.insertSheet", json!({"name":"Dashboard"}))?;
    session.execute("sheet.activate", json!({"sheet":1}))?;
    {
        let document = session.doc_mut()?;
        let workbook = std::sync::Arc::make_mut(&mut document.wb);
        workbook.code_name = Some("BookCode".into());
        workbook.sheet_mut(0).ok_or("data sheet missing")?.name = "Data".into();
    }
    let mut ide = Ide::new();
    ide.add_document_module(
        "ThisWorkbook.cls",
        "BookCode",
        "Attribute VB_Name = \"BookCode\"\nPrivate Sub Workbook_Open()\nRange(\"A1\").Value2 = ActiveSheet.Name\nMe.Worksheets(\"Data\").Range(\"A2\").Value2 = Me.Name\nEnd Sub\n",
    )?;
    let mut bridge = Bridge::new(&session)?;
    bridge.consent(true);
    bridge.configure_document(&mut ide, &session.doc()?.wb)?;
    bridge.queue_workbook_open(&session)?;
    tick_until(&mut bridge, &mut session, &mut ide, Mode::Completed)?;
    // Bare Range follows the active sheet even inside the workbook module.
    let active = session.doc()?.wb.sheet(1).ok_or("active sheet missing")?.clone();
    assert_eq!(active.value(CellRef::new(0, 0)), Value::Text("Dashboard".into()));
    // Me stays the project workbook with its own member set.
    let data = session.doc()?.wb.sheet(0).ok_or("data sheet missing")?.clone();
    assert_eq!(data.value(CellRef::new(1, 0)), Value::Text("Book1".into()));
    Ok(())
}

#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn host_globals_and_call_shaped_receivers_complete_without_host_calls() -> R {
    let mut session = Session::new();
    session.new_workbook();
    {
        let document = session.doc_mut()?;
        let workbook = std::sync::Arc::make_mut(&mut document.wb);
        workbook.sheet_mut(0).ok_or("sheet missing")?.code_name = Some("Sheet1".into());
    }
    let mut bridge = Bridge::new(&session)?;
    bridge.consent(true);
    let mut ide = Ide::new();
    ide.add_module("Main.bas", "Sub Main()\nRange(\"A1\").Value2 = 1\nEnd Sub\n")?;
    bridge.configure_document(&mut ide, &session.doc()?.wb)?;
    // Declared host globals complete through the shared editor with no dispatch.
    assert!(ide.completions("ThisWorkbook.W").contains(&"Worksheets".to_string()));
    assert!(ide.completions("ActiveSheet.C").contains(&"Cells".to_string()));
    assert!(ide.completions("Selection.F").contains(&"Font".to_string()));
    // Call-shaped receivers resolve through declared default members.
    assert!(ide.completions("Range(\"A1\").V").contains(&"Value2".to_string()));
    assert!(ide.completions("Worksheets(\"Data\").R").contains(&"Range".to_string()));
    assert!(ide.completions("ThisWorkbook.Worksheets(\"Data\").C").contains(&"Columns".to_string()));
    // A code-name receiver completes the declared worksheet instance type.
    assert!(ide.completions("Sheet1.R").contains(&"Range".to_string()));
    assert_eq!(ide.mode(), Mode::Design);
    assert!(session.doc()?.wb.sheet(0).ok_or("sheet missing")?.value(CellRef::parse("A1").ok_or("cell")?).is_empty());
    Ok(())
}

#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn worksheet_event_globals_consult_live_state_resumed_after_activation() -> R {
    let mut session = Session::new();
    session.new_workbook();
    session.execute("home.insertSheet", json!({"name":"Other"}))?;
    {
        let document = session.doc_mut()?;
        let workbook = std::sync::Arc::make_mut(&mut document.wb);
        workbook.sheet_mut(0).ok_or("owner sheet missing")?.code_name = Some("SheetCode".into());
        workbook.set_active_sheet(1);
    }
    let source = "Attribute VB_Name = \"SheetCode\"\nPrivate Sub Worksheet_Change(ByVal Target As Range)\nRange(\"A1\").Value2 = \"module-sheet\"\nRange(\"A2\").Value2 = ActiveSheet.Name\nRange(\"A3\").Value2 = Selection.Address\nEnd Sub\n";
    let mut ide = Ide::new();
    ide.add_document_module("Sheet1.cls", "SheetCode", source)?;
    let breakpoint = u32::try_from(source.find("Range(\"A2\")").ok_or("resume statement missing")?)?;
    ide.modules.first_mut().ok_or("document module missing")?.breakpoints.insert(breakpoint);
    let mut bridge = Bridge::new(&session)?;
    bridge.consent(true);
    bridge.configure_document(&mut ide, &session.doc()?.wb)?;
    bridge.notify_sheet_change(&session, 0, RangeRef::cell(CellRef::new(4, 3)))?;
    tick_until(&mut bridge, &mut session, &mut ide, Mode::Break)?;
    // The user activates another sheet and moves the selection while paused.
    session.execute("sheet.activate", json!({"sheet":1}))?;
    session.execute("selection.set", json!({"range":"D4:E5"}))?;
    // Clear the breakpoint so the resumed delivery (and any re-queued change
    // event from the handler's own reviewed edits) settles to completion.
    ide.modules.first_mut().ok_or("document module missing")?.breakpoints.clear();
    bridge.dispatch(&mut session, &mut ide, gridcraft_vba::Action::Run)?;
    for _ in 0..400 {
        bridge.tick(&mut session, &mut ide)?;
        if matches!(ide.mode(), Mode::Completed | Mode::Design) {
            break;
        }
    }
    let sheet = session.doc()?.wb.sheet(0).ok_or("owner sheet missing")?.clone();
    assert_eq!(sheet.value(CellRef::new(0, 0)), Value::Text("module-sheet".into()));
    assert_eq!(sheet.value(CellRef::new(1, 0)), Value::Text("Other".into()));
    assert_eq!(sheet.value(CellRef::new(2, 0)), Value::Text("$D$4:$E$5".into()));
    let other = session.doc()?.wb.sheet(1).ok_or("active sheet missing")?.clone();
    assert_eq!(other.value(CellRef::new(0, 0)), Value::Empty);
    Ok(())
}
