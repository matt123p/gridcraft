#![cfg(feature = "vba")]
use gridcraft_engine::Session;
use gridcraft_ui_egui::{Services, SheetApp, printing::Printer};
use gridcraft_vba::Mode;
use serde_json::json;
use std::{cell::RefCell, rc::Rc};
fn app(source: &str) -> SheetApp {
    let mut s = Session::new();
    s.new_workbook();
    let mut a = SheetApp::new(s, Services::default());
    a.run("vba.open", json!({"source":source})).unwrap();
    a
}
#[test]
fn vba_printout_uses_actual_worker_and_native_printer_settings() {
    let mut a = app(
        "Sub Main()\nApplication.ActivePrinter=\"Chosen queue\"\nOn Error Resume Next\nActiveSheet.PrintOut Copies:=3\nRange(\"A1\").Value2=Err.Number\nEnd Sub",
    );
    let jobs = Rc::new(RefCell::new(Vec::new()));
    let seen = jobs.clone();
    a.services.print_native = Some(Box::new(move |pdf, settings| {
        seen.borrow_mut().push((pdf, settings));
        let (tx, rx) = std::sync::mpsc::channel();
        tx.send(Err("queue offline".into())).unwrap();
        rx
    }));
    a.printing.printers.push(Printer { name: "Chosen queue".into(), label: String::new(), default: false });
    a.run("vba.run", json!({})).unwrap();
    let ctx = egui::Context::default();
    for _ in 0..100 {
        gridcraft_ui_egui::vba::tick(&mut a, &ctx);
        if a.vba.as_ref().unwrap().ide.mode() == Mode::Completed {
            break;
        }
    }
    assert_eq!(a.vba.as_ref().unwrap().ide.mode(), Mode::Completed, "{:?}", a.vba.as_ref().unwrap().ide.diagnostics);
    assert_eq!(jobs.borrow().len(), 1);
    assert_eq!(jobs.borrow()[0].1.copies, 3);
    assert_eq!(jobs.borrow()[0].1.printer, "Chosen queue");
    assert!(jobs.borrow()[0].0.starts_with(b"%PDF"));
    assert_eq!(a.session.doc().unwrap().wb.active().unwrap().value(gridcraft_engine::core::CellRef::new(0, 0)).display(), "1004");
}
#[test]
fn vba_preview_opens_native_backstage_and_resumes_only_when_closed() {
    let mut a = app("Sub Main()\nActiveSheet.PrintPreview\nRange(\"A1\").Value2=7\nEnd Sub");
    a.run("vba.run", json!({})).unwrap();
    let ctx = egui::Context::default();
    for _ in 0..100 {
        gridcraft_ui_egui::vba::tick(&mut a, &ctx);
        if a.backstage.is_some() {
            break;
        }
    }
    assert!(a.backstage.is_some());
    assert!(!a.vba.as_ref().unwrap().visible);
    assert_eq!(a.vba.as_ref().unwrap().ide.mode(), Mode::Waiting);
    for _ in 0..3 {
        gridcraft_ui_egui::vba::tick(&mut a, &ctx);
        assert_eq!(a.vba.as_ref().unwrap().ide.mode(), Mode::Waiting);
    }
    a.backstage = None;
    for _ in 0..100 {
        gridcraft_ui_egui::vba::tick(&mut a, &ctx);
        if a.vba.as_ref().unwrap().ide.mode() == Mode::Completed {
            break;
        }
    }
    assert_eq!(a.vba.as_ref().unwrap().ide.mode(), Mode::Completed, "{:?}", a.vba.as_ref().unwrap().ide.diagnostics);
    assert_eq!(a.session.doc().unwrap().wb.active().unwrap().value(gridcraft_engine::core::CellRef::new(0, 0)).display(), "7");
    assert!(a.vba.as_ref().unwrap().visible);
    if let Ok(path) = std::env::var("GRIDCRAFT_OM10_PREVIEW_CAPTURE") {
        a.run("vba.run", json!({})).unwrap();
        let mut h = egui_kittest::Harness::builder().with_size(egui::vec2(1440.0, 900.0)).with_max_steps(30).wgpu().build_ui_state(
            |ui, app: &mut SheetApp| {
                let ctx = ui.ctx().clone();
                app.logic(&ctx);
                app.ui(ui);
            },
            a,
        );
        h.run_steps(16);
        h.render().unwrap().save(path).unwrap();
    }
}
#[test]
fn vba_export_picker_cancellation_returns_to_the_macro_without_writing() {
    let mut a = app("Sub Main()\nOn Error Resume Next\nActiveSheet.ExportAsFixedFormat xlTypePDF\nRange(\"A1\").Value2=Err.Number\nEnd Sub");
    let picks = Rc::new(RefCell::new(Vec::new()));
    let seen = picks.clone();
    a.services.pick_save = Some(Box::new(move |name| {
        seen.borrow_mut().push(name.to_string());
        None
    }));
    a.run("vba.run", json!({})).unwrap();
    let ctx = egui::Context::default();
    for _ in 0..100 {
        gridcraft_ui_egui::vba::tick(&mut a, &ctx);
        if a.vba.as_ref().unwrap().ide.mode() == Mode::Completed {
            break;
        }
    }
    assert_eq!(a.vba.as_ref().unwrap().ide.mode(), Mode::Completed, "{:?}", a.vba.as_ref().unwrap().ide.diagnostics);
    assert_eq!(picks.borrow().len(), 1);
    assert!(picks.borrow()[0].ends_with(".pdf"));
    assert_eq!(a.session.doc().unwrap().wb.active().unwrap().value(gridcraft_engine::core::CellRef::new(0, 0)).display(), "1004");
    assert!(a.session.doc().unwrap().path.is_none());
}
#[test]
fn vba_export_picker_delegates_to_native_pdf_export_and_keeps_workbook_path() {
    let mut a = app("Sub Main()\nRange(\"A1\").Value2=\"Picker report\"\nActiveSheet.ExportAsFixedFormat xlTypePDF\nEnd Sub");
    let path = std::env::temp_dir().join(format!("gridcraft-om10-picker-{}.pdf", std::process::id()));
    let selected = path.to_string_lossy().into_owned();
    a.services.pick_save = Some(Box::new(move |_| Some(selected.clone())));
    a.run("vba.run", json!({})).unwrap();
    let ctx = egui::Context::default();
    for _ in 0..100 {
        gridcraft_ui_egui::vba::tick(&mut a, &ctx);
        if a.vba.as_ref().unwrap().ide.mode() == Mode::Completed {
            break;
        }
    }
    assert_eq!(a.vba.as_ref().unwrap().ide.mode(), Mode::Completed, "{:?}", a.vba.as_ref().unwrap().ide.diagnostics);
    let pdf = std::fs::read(&path).unwrap();
    assert!(pdf.starts_with(b"%PDF"));
    assert!(String::from_utf8_lossy(&pdf).contains("(Picker report) Tj"));
    assert!(a.session.doc().unwrap().path.is_none());
    std::fs::remove_file(path).unwrap();
}
#[test]
fn vba_unavailable_printer_never_submits_a_job() {
    let mut a = app("Sub Main()\nOn Error Resume Next\nActiveSheet.PrintOut\nRange(\"A1\").Value2=Err.Number\nEnd Sub");
    let count = Rc::new(RefCell::new(0));
    let seen = count.clone();
    a.services.print_native = Some(Box::new(move |_, _| {
        *seen.borrow_mut() += 1;
        let (tx, rx) = std::sync::mpsc::channel();
        tx.send(Ok("job".into())).unwrap();
        rx
    }));
    a.run("vba.run", json!({})).unwrap();
    let ctx = egui::Context::default();
    for _ in 0..100 {
        gridcraft_ui_egui::vba::tick(&mut a, &ctx);
        if a.vba.as_ref().unwrap().ide.mode() == Mode::Completed {
            break;
        }
    }
    assert_eq!(a.vba.as_ref().unwrap().ide.mode(), Mode::Completed, "{:?}", a.vba.as_ref().unwrap().ide.diagnostics);
    assert_eq!(*count.borrow(), 0);
    assert_eq!(a.session.doc().unwrap().wb.active().unwrap().value(gridcraft_engine::core::CellRef::new(0, 0)).display(), "1004");
}
