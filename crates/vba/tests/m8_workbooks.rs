//! M8 independently authored workbook workflows; no fresh Office observations.
mod support;
use gridcraft_engine::{
    DocState, Session, Trust,
    core::{CellRef, Value},
    model::Workbook,
};
use gridcraft_vba::{Action, Bridge, Mode};
use gridcraft_xlsx::{XlsxKind, read_xlsx, write_xlsx_as};
use serde_json::json;
type R = Result<(), Box<dyn std::error::Error>>;
fn cell(wb: &Workbook, address: &str) -> Value {
    wb.sheet(0).unwrap().value(CellRef::parse(address).unwrap())
}
fn workflow(source: &str, expected: Value, expected_type: &str) -> R {
    let mut initial = Session::new();
    initial.new_workbook();
    for (address, input) in [("A1", "12"), ("A2", "10"), ("A3", "8"), ("A4", " alice "), ("A5", "A😀")] {
        initial.execute("cell.set", json!({"cell":address,"input":input}))?;
    }
    let mut wb = initial.doc()?.wb.as_ref().clone();
    wb.sheet_mut(0).unwrap().name = "Data".into();
    let project = support::project(&[("Workflow", "Module", source)]);
    wb.vba_project = Some(project.clone().into());
    let bytes = write_xlsx_as(&wb, XlsxKind::MacroWorkbook)?;
    let (loaded, _) = read_xlsx(&bytes)?;
    let mut session = Session::new();
    session.add_document(DocState::from_saved(loaded, None, "Workflow.xlsm".into(), Trust::FromFile));
    let mut bridge = Bridge::new(&session)?;
    let mut ide = bridge.import_project(&session)?;
    assert_eq!(ide.mode(), Mode::Design);
    assert!(ide.output().is_empty());
    assert_eq!(ide.modules.first().unwrap().text, source);
    assert_eq!(cell(&session.doc()?.wb, "B1"), Value::Empty);
    ide.entry = "Main".into();
    bridge.dispatch(&mut session, &mut ide, Action::Compile)?;
    assert!(bridge.dispatch(&mut session, &mut ide, Action::Run).is_err());
    assert_eq!(cell(&session.doc()?.wb, "B1"), Value::Empty);
    bridge.consent(true);
    bridge.dispatch(&mut session, &mut ide, Action::Run)?;
    for _ in 0..1000 {
        bridge.tick(&mut session, &mut ide)?;
        if !matches!(ide.mode(), Mode::Running | Mode::Waiting) {
            break;
        }
    }
    assert_eq!(ide.mode(), Mode::Completed, "{:?}", ide.diagnostics);
    assert_eq!(cell(&session.doc()?.wb, "B1"), expected);
    assert_eq!(ide.output(), vec![expected_type.to_string()]);
    for address in ["A1", "A2", "A3", "A4", "A5"] {
        assert_eq!(cell(&session.doc()?.wb, address), cell(&wb, address));
    }
    let saved = write_xlsx_as(&session.doc()?.wb, XlsxKind::MacroWorkbook)?;
    let (back, _) = read_xlsx(&saved)?;
    assert_eq!(cell(&back, "B1"), expected);
    assert_eq!(back.vba_project.as_deref(), Some(&project));
    let mut reopened = Session::new();
    reopened.add_document(DocState::from_saved(back, None, "Reopened.xlsm".into(), Trust::FromFile));
    let reopened_bridge = Bridge::new(&reopened)?;
    let reopened_ide = reopened_bridge.import_project(&reopened)?;
    assert_eq!(reopened_ide.modules.first().unwrap().text, source);
    assert!(reopened_ide.output().is_empty());
    assert_eq!(session.doc()?.undo.len(), 1);
    session.execute("edit.undo", json!({}))?;
    assert_eq!(cell(&session.doc()?.wb, "B1"), Value::Empty);
    bridge.dispatch(&mut session, &mut ide, Action::Reset)?;
    assert_eq!(session.doc()?.wb.vba_project.as_deref(), Some(&project));
    Ok(())
}

#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn invoice() -> R {
    workflow(include_str!("../../../tests/fixtures/m8/invoice.bas"), Value::Number(120.0), "Double")
}

#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn sales_total() -> R {
    workflow(include_str!("../../../tests/fixtures/m8/sales-total.bas"), Value::Number(30.0), "Double")
}

#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn discount() -> R {
    workflow(include_str!("../../../tests/fixtures/m8/discount.bas"), Value::Number(10.8), "Double")
}

#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn positive_balance() -> R {
    workflow(include_str!("../../../tests/fixtures/m8/positive-balance.bas"), Value::Number(30.0), "Double")
}

#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn normalize_name() -> R {
    workflow(include_str!("../../../tests/fixtures/m8/normalize-name.bas"), Value::Text("ALICE".into()), "String")
}

#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn leap_date() -> R {
    workflow(include_str!("../../../tests/fixtures/m8/leap-date.bas"), Value::Number(45351.0), "Double")
}

#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn currency_invoice() -> R {
    workflow(include_str!("../../../tests/fixtures/m8/currency-invoice.bas"), Value::Number(19.95), "Double")
}

#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn bankers_rounding() -> R {
    workflow(include_str!("../../../tests/fixtures/m8/bankers-rounding.bas"), Value::Number(2.0), "Double")
}

#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn overflow_recovery() -> R {
    workflow(include_str!("../../../tests/fixtures/m8/overflow-recovery.bas"), Value::Number(6.0), "Double")
}

#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn bounded_array() -> R {
    workflow(include_str!("../../../tests/fixtures/m8/bounded-array.bas"), Value::Number(6.0), "Double")
}

#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn byref_update() -> R {
    workflow(include_str!("../../../tests/fixtures/m8/byref-update.bas"), Value::Number(14.0), "Double")
}

#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn static_counter() -> R {
    workflow(include_str!("../../../tests/fixtures/m8/static-counter.bas"), Value::Number(2.0), "Double")
}

#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn bucket_label() -> R {
    workflow(include_str!("../../../tests/fixtures/m8/bucket-label.bas"), Value::Text("medium".into()), "String")
}

#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn null_predicate() -> R {
    workflow(include_str!("../../../tests/fixtures/m8/null-predicate.bas"), Value::Bool(true), "Boolean")
}

#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn formula_recalc() -> R {
    workflow(include_str!("../../../tests/fixtures/m8/formula-recalc.bas"), Value::Number(30.0), "Double")
}

#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn named_rate() -> R {
    workflow(include_str!("../../../tests/fixtures/m8/named-rate.bas"), Value::Number(2.0), "Double")
}

#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn worksheet_context() -> R {
    workflow(include_str!("../../../tests/fixtures/m8/worksheet-context.bas"), Value::Number(12.0), "Double")
}

#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn foreach_total() -> R {
    workflow(include_str!("../../../tests/fixtures/m8/foreach-total.bas"), Value::Number(10.0), "Double")
}

#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn early_exit() -> R {
    workflow(include_str!("../../../tests/fixtures/m8/early-exit.bas"), Value::Number(3.0), "Double")
}

#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn utf16_length() -> R {
    workflow(include_str!("../../../tests/fixtures/m8/utf16-length.bas"), Value::Number(3.0), "Double")
}
