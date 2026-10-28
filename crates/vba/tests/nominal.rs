//! Original unchanged Excel source. Excel definitions and assertions stay in this fork.
use gridcraft_engine::{
    Session,
    core::{CellRef, Value},
};
use gridcraft_vba::{Action, Bridge, Ide, Mode};
type R = Result<(), Box<dyn std::error::Error>>;
fn prepare(source: &str) -> Result<(Session, Bridge, Ide), Box<dyn std::error::Error>> {
    let mut session = Session::new();
    session.new_workbook();
    let mut bridge = Bridge::new(&session)?;
    bridge.consent(true);
    let mut ide = Ide::new();
    ide.add_module("Original.bas", source)?;
    bridge.configure(&mut ide)?;
    Ok((session, bridge, ide))
}
fn drain(session: &mut Session, bridge: &mut Bridge, ide: &mut Ide) -> R {
    for _ in 0..1000 {
        bridge.tick(session, ide)?;
        if !matches!(ide.mode(), Mode::Running | Mode::Waiting) {
            return Ok(());
        }
    }
    Err("macro timeout".into())
}
#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn unchanged_excel_types_parameters_results_and_real_cell_writes() -> R {
    const SOURCE: &str = "Option Explicit\nSub Main()\nDim app As Excel.Application\nDim book As Workbook\nDim ws As Worksheet\nDim cell As Excel.Range\nSet app=Application\nSet book=ThisWorkbook\nSet ws=book.Worksheets(1)\nSet cell=Pick(ws)\nWriteCell cell\nSet cell=Nothing\nEnd Sub\nPrivate Function Pick(ByVal ws As Excel.Worksheet) As Excel.Range\nSet Pick=ws.Range(\"A1\")\nEnd Function\nPrivate Sub WriteCell(ByRef cell As Excel.Range)\ncell.Value2=42\nEnd Sub";
    let (mut session, mut bridge, mut ide) = prepare(SOURCE)?;
    bridge.dispatch(&mut session, &mut ide, Action::Run)?;
    drain(&mut session, &mut bridge, &mut ide)?;
    assert_eq!(ide.mode(), Mode::Completed, "{:?}", ide.messages);
    assert_eq!(session.doc()?.wb.sheet(0).ok_or("sheet")?.value(CellRef::parse("A1").ok_or("cell")?), Value::Number(42.0));
    assert_eq!(ide.active_module().ok_or("source")?.text, SOURCE);
    assert!(!session.in_group(bridge.document()));
    Ok(())
}
#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn wrong_excel_class_fails_without_retargeting_or_rolling_back_prior_effects() -> R {
    let source = "Sub Main()\nRange(\"B1\").Value2=25\nDim ws As Worksheet\nSet ws=Range(\"A1\")\nws.Range(\"A1\").Value2=99\nEnd Sub";
    let (mut session, mut bridge, mut ide) = prepare(source)?;
    bridge.dispatch(&mut session, &mut ide, Action::Run)?;
    drain(&mut session, &mut bridge, &mut ide)?;
    assert_eq!(ide.mode(), Mode::Failed);
    assert!(ide.diagnostics.iter().any(|d| d.message.contains("Error 13:")), "{:?}", ide.diagnostics);
    let sheet = session.doc()?.wb.sheet(0).ok_or("sheet")?;
    assert!(sheet.value(CellRef::parse("A1").ok_or("A1")?).is_empty());
    assert_eq!(sheet.value(CellRef::parse("B1").ok_or("B1")?), Value::Number(25.0));
    assert!(!session.in_group(bridge.document()));
    Ok(())
}
#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn paused_excel_reference_edits_keep_the_declared_type() -> R {
    let source = "Sub Main()\nDim ws As Excel.Worksheet\nDim cell As Range\nSet ws=ActiveSheet\nSet cell=ws.Range(\"A1\")\nDebug.Print 1\ncell.Value2=7\nEnd Sub";
    let (mut session, mut bridge, mut ide) = prepare(source)?;
    ide.toggle_breakpoint_at(source.find("Debug.Print").ok_or("breakpoint")?)?;
    bridge.dispatch(&mut session, &mut ide, Action::Run)?;
    drain(&mut session, &mut bridge, &mut ide)?;
    assert_eq!(ide.mode(), Mode::Break);
    assert!(ide.completions("ws.R").contains(&"Range".to_string()));
    assert!(ide.completions("cell.V").contains(&"Value2".to_string()));
    let original = ide.evaluate("cell")?;
    assert!(ide.edit_local("cell", "ws").is_err());
    assert_eq!(ide.evaluate("cell")?, original);
    bridge.dispatch(&mut session, &mut ide, Action::Run)?;
    drain(&mut session, &mut bridge, &mut ide)?;
    assert_eq!(ide.mode(), Mode::Completed);
    assert_eq!(session.doc()?.wb.sheet(0).ok_or("sheet")?.value(CellRef::parse("A1").ok_or("cell")?), Value::Number(7.0));
    Ok(())
}
