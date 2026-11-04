//! M7 authored unchanged-document acceptance. Expectations are host contracts, not Office observations.
mod support;
use gridcraft_engine::{
    DocState, Session, Trust,
    core::{CellRef, Value},
    model::Workbook,
};
use gridcraft_vba::{Action, Bridge, Ide, Mode};
use gridcraft_xlsx::{XlsxKind, read_xlsx, write_xlsx_as};
use std::sync::Arc;
type R = Result<(), Box<dyn std::error::Error>>;
const STANDARD: &str = include_str!("../../../tests/fixtures/m7/Package.bas");
const SHEET: &str = include_str!("../../../tests/fixtures/m7/SheetCode.cls");
const SINK: &str = include_str!("../../../tests/fixtures/m7/AppSink.cls");
const COUNTER: &str = include_str!("../../../tests/fixtures/m7/Counter.cls");
const BOOK: &str = include_str!("../../../tests/fixtures/m7/BookCode.cls");
fn package() -> Vec<u8> {
    let mut wb = Workbook::new();
    wb.code_name = Some("BookCode".into());
    wb.sheet_mut(0).unwrap().name = "Data".into();
    wb.sheet_mut(0).unwrap().code_name = Some("SheetCode".into());
    wb.vba_project = Some(
        support::project(&[
            ("Package", "Module", STANDARD),
            ("SheetCode", "Document", SHEET),
            ("BookCode", "Document", BOOK),
            ("Counter", "Class", COUNTER),
            ("AppSink", "Class", SINK),
        ])
        .into(),
    );
    write_xlsx_as(&wb, XlsxKind::MacroWorkbook).unwrap()
}
fn load(bytes: &[u8]) -> (Session, Bridge, Ide) {
    let (wb, _) = read_xlsx(bytes).unwrap();
    let mut s = Session::new();
    s.add_document(DocState::from_saved(wb, None, "A.xlsm".into(), Trust::FromFile));
    let b = Bridge::new(&s).unwrap();
    let ide = b.import_project(&s).unwrap();
    (s, b, ide)
}
fn settle(s: &mut Session, b: &mut Bridge, ide: &mut Ide) {
    for _ in 0..200 {
        b.tick(s, ide).unwrap();
        if matches!(ide.mode(), Mode::Completed | Mode::Failed) {
            assert_eq!(ide.mode(), Mode::Completed, "{:?}", ide.diagnostics);
            return;
        }
    }
    panic!("M7 acceptance timeout");
}
fn value(wb: &Workbook, sheet: usize, address: &str) -> Value {
    wb.sheet(sheet).unwrap().value(CellRef::parse(address).unwrap())
}
#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn unchanged_xlsm_load_compile_trust_run_save_reopen() -> R {
    let bytes = package();
    let (mut s, mut b, mut ide) = load(&bytes);
    let original = s.doc()?.wb.vba_project.clone().unwrap();
    assert_eq!(ide.mode(), Mode::Design);
    assert_eq!(ide.modules.iter().map(|m| m.text.as_str()).collect::<Vec<_>>(), [STANDARD, SHEET, BOOK, COUNTER, SINK]);
    assert_eq!(value(&s.doc()?.wb, 0, "Z1"), Value::Empty);
    ide.entry = "Main".into();
    b.dispatch(&mut s, &mut ide, Action::Compile)?;
    assert!(b.dispatch(&mut s, &mut ide, Action::Run).is_err(), "load/compile must not grant trust");
    b.consent(true);
    // The code belongs to A; every global call must consult live B.
    let owner = b.document();
    s.new_workbook();
    b.dispatch(&mut s, &mut ide, Action::Run)?;
    settle(&mut s, &mut b, &mut ide);
    assert_eq!(value(&s.doc()?.wb, 0, "A1"), Value::Number(11.));
    assert_eq!(value(&s.doc()?.wb, 0, "A3"), Value::Number(33.));
    let a = &s.document(owner).unwrap().wb;
    for (address, n) in [("A2", 22.), ("A4", 44.), ("B4", 44.), ("A5", 6.), ("A6", 11.), ("A9", 301.)] {
        assert_eq!(value(a, 0, address), Value::Number(n), "{address}");
    }
    assert_eq!(value(a, 0, "A1"), Value::Empty);
    ide.entry = "Shadow".into();
    b.dispatch(&mut s, &mut ide, Action::Run)?;
    settle(&mut s, &mut b, &mut ide);
    assert_eq!(value(&s.document(owner).unwrap().wb, 0, "A8"), Value::Number(88.));
    ide.entry = "Mixed".into();
    b.dispatch(&mut s, &mut ide, Action::Run)?;
    settle(&mut s, &mut b, &mut ide);
    assert_eq!(value(&s.document(owner).unwrap().wb, 0, "A10"), Value::Number(1004.));
    assert_eq!(value(&s.doc()?.wb, 0, "A10"), Value::Empty);
    let saved = write_xlsx_as(&s.document(owner).unwrap().wb, XlsxKind::MacroWorkbook)?;
    let (back, _) = read_xlsx(&saved)?;
    assert_eq!(back.vba_project.as_deref(), Some(original.as_ref()));
    assert_eq!(back.code_name.as_deref(), Some("BookCode"));
    assert_eq!(back.sheet(0).unwrap().code_name.as_deref(), Some("SheetCode"));
    let (_, _, reopened) = load(&saved);
    assert_eq!(reopened.modules.iter().map(|m| m.text.as_str()).collect::<Vec<_>>(), [STANDARD, SHEET, BOOK, COUNTER, SINK]);
    Ok(())
}
#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn packaged_module_contexts_survive_tab_rename_with_another_workbook_active() -> R {
    let (mut s, mut b, mut ide) = load(&package());
    s.execute("home.insertSheet", serde_json::json!({"name":"Other"}))?;
    Arc::make_mut(&mut s.doc_mut()?.wb).sheet_mut(0).unwrap().name = "Renamed".into();
    b.consent(true);
    b.configure_document(&mut ide, &s.doc()?.wb)?;
    b.notify_sheet_change(&s, 0, gridcraft_engine::core::RangeRef::cell(CellRef::new(9, 9)))?;
    let owner = b.document();
    s.new_workbook();
    settle(&mut s, &mut b, &mut ide);
    let a = &s.document(owner).unwrap().wb;
    for (address, n) in [("B1", 101.), ("B2", 102.), ("B3", 103.), ("B5", 105.), ("B6", 6.), ("B7", 3.)] {
        assert_eq!(value(a, 0, address), Value::Number(n), "{address}");
    }
    assert_eq!(value(a, 1, "B4"), Value::Number(104.));
    assert_eq!(value(&s.doc()?.wb, 0, "A7"), Value::Number(77.), "standard helper keeps its own module context");
    assert_eq!(a.sheet(0).unwrap().code_name.as_deref(), Some("SheetCode"));
    let renamed = write_xlsx_as(a, XlsxKind::MacroWorkbook)?;
    let (renamed_back, _) = read_xlsx(&renamed)?;
    assert_eq!(renamed_back.sheet(0).unwrap().name, "Renamed");
    assert_eq!(renamed_back.sheet(0).unwrap().code_name.as_deref(), Some("SheetCode"));
    // Workbook event: bare Range is global, explicit Me remains its project workbook.
    s.set_active(s.document_index(owner).unwrap());
    Arc::make_mut(&mut s.doc_mut()?.wb).sheet_mut(0).unwrap().name = "Data".into();
    b.queue_workbook_open(&s)?;
    s.set_active(1);
    settle(&mut s, &mut b, &mut ide);
    assert_eq!(value(&s.doc()?.wb, 0, "C1"), Value::Number(201.));
    assert_eq!(value(&s.document(owner).unwrap().wb, 0, "C2"), Value::Number(202.));
    let saved = write_xlsx_as(&s.document(owner).unwrap().wb, XlsxKind::MacroWorkbook)?;
    assert_eq!(read_xlsx(&saved)?.0.vba_project, s.document(owner).unwrap().wb.vba_project);
    Ok(())
}

#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn packaged_code_edits_export_hidden_records_without_claiming_binary_writeback() -> R {
    let (mut s, mut b, mut ide) = load(&package());
    let original = s.doc()?.wb.vba_project.clone();
    ide.active = 0;
    let display = gridcraft_vba::vba_ide::editor_text(STANDARD);
    assert!(!display.contains("Attribute"));
    let edited = ide.active_module().unwrap().edited_source(&display.replace("= 11", "= 12"));
    ide.set_text(edited.clone())?;
    assert!(edited.contains("Attribute Main.VB_ProcData.VB_Invoke_Func"));
    assert!(edited.starts_with("Attribute VB_Name"));
    b.dispatch(&mut s, &mut ide, Action::Save)?;
    let requests = ide.take_requests();
    assert!(requests.iter().any(|r| matches!(r,gridcraft_vba::vba_ide::HostRequest::Save {text,..} if text == &edited)));
    let m = ide.active_module().unwrap();
    let (id, revision) = (m.id, m.revision);
    ide.set_text(format!("{edited}\n' intervening edit\n"))?;
    assert!(ide.mark_saved(id, revision).is_err());
    let bytes = write_xlsx_as(&s.doc()?.wb, XlsxKind::MacroWorkbook)?;
    let (_, _, reopened) = load(&bytes);
    assert_eq!(reopened.modules[0].text, STANDARD, "document save preserves the original binary project");
    assert_eq!(s.doc()?.wb.vba_project, original);
    Ok(())
}

#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn packaged_live_globals_follow_activation_between_calls_without_replaying() -> R {
    let (mut s, mut b, mut ide) = load(&package());
    let owner = b.document();
    b.consent(true);
    s.new_workbook();
    ide.entry = "Live".into();
    b.dispatch(&mut s, &mut ide, Action::Run)?;
    for _ in 0..200 {
        b.tick(&mut s, &mut ide)?;
        if ide.mode() == Mode::Break {
            break;
        }
    }
    assert_eq!(ide.mode(), Mode::Break);
    assert_eq!(value(&s.doc()?.wb, 0, "D1"), Value::Number(401.));
    s.set_active(s.document_index(owner).unwrap());
    b.dispatch(&mut s, &mut ide, Action::Run)?;
    settle(&mut s, &mut b, &mut ide);
    assert_eq!(value(&s.doc()?.wb, 0, "D1"), Value::Empty, "first call was not replayed");
    assert_eq!(value(&s.doc()?.wb, 0, "D2"), Value::Number(402.));
    assert_eq!(value(&s.documents()[1].wb, 0, "D2"), Value::Empty);
    Ok(())
}

#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn packaged_excel_withevents_is_live_only_until_reset() -> R {
    use gridcraft_vba::LifecycleEvent;
    let (mut s, mut b, mut ide) = load(&package());
    let owner = b.document();
    b.consent(true);
    ide.entry = "Connect".into();
    b.dispatch(&mut s, &mut ide, Action::Run)?;
    settle(&mut s, &mut b, &mut ide);
    let mut batch =
        b.prepare_lifecycle_events(&mut s, &mut ide, owner, LifecycleEvent::BeforeSave { save_as_ui: false }).map_err(|e| e.description)?;
    assert!(!batch.is_empty());
    let mut cancelled = None;
    for _ in 0..200 {
        if let Some(value) = batch.poll(&mut b, &mut s, &mut ide).map_err(|e| e.description)? {
            cancelled = Some(value);
            break;
        }
        b.tick_lifecycle(&mut s, &mut ide)?;
    }
    assert_eq!(cancelled, Some(true));
    assert_eq!(value(&s.doc()?.wb, 0, "E1"), Value::Number(501.));
    b.finish(&mut s)?;
    b.dispatch(&mut s, &mut ide, Action::Reset)?;
    let batch = b.prepare_lifecycle_events(&mut s, &mut ide, owner, LifecycleEvent::BeforeSave { save_as_ui: false }).map_err(|e| e.description)?;
    assert!(batch.is_empty(), "reset discards publisher subscriptions");
    b.finish(&mut s)?;
    Ok(())
}
