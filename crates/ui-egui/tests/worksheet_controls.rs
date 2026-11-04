#![cfg(feature = "vba")]
#[allow(dead_code)]
mod common;
use egui_kittest::kittest::Queryable as _;
use gridcraft_engine::{
    DocState, Session, Trust,
    core::{CellRef, Value},
};
use gridcraft_ui_egui::SheetApp;
use gridcraft_vba::Mode;
use serde_json::{Value as Json, json};
const SOURCE: &str = "Public Sub Hit()\nRange(\"D1\").Value2=Range(\"D1\").Value2+1\nRange(\"D2\").Value2=Range(\"A1\").Value2\nRange(\"D3\").Value2=Application.Caller\nEnd Sub\nPrivate Sub Hidden()\nEnd Sub\n";
fn app() -> (SheetApp, Json) {
    let mut s = Session::new();
    s.new_workbook();
    let mut a = common::app(s);
    a.run("vba.open", json!({"source":SOURCE})).unwrap();
    a.run("vba.apply", json!({})).unwrap();
    a.run("vba.return", json!({})).unwrap();
    let id = a
        .run("control.add", json!({"type":1,"name":"Agree","caption":"Agree to continue","left":200,"top":70,"width":180,"height":30}))
        .unwrap()["shape"]
        .clone();
    a.run("control.set", json!({"id":id,"linkedCell":"A1"})).unwrap();
    a.run("shape.setMacro", json!({"id":id,"entry":"[0]!Hit"})).unwrap();
    (a, id)
}
fn value(a: &SheetApp, cell: &str) -> Value {
    a.session.doc().unwrap().wb.sheet(0).unwrap().value(CellRef::parse(cell).unwrap())
}
fn finish(a: &mut SheetApp) {
    for _ in 0..80 {
        gridcraft_ui_egui::vba::tick(a, &egui::Context::default());
        if a.vba.as_ref().unwrap().ide.mode() == Mode::Completed {
            break;
        }
    }
    assert_eq!(a.vba.as_ref().unwrap().ide.mode(), Mode::Completed, "{:?}", a.vba.as_ref().unwrap().ide.diagnostics);
}
#[test]
fn pointer_and_keyboard_activate_once_and_macro_sees_committed_link_and_caller() {
    let (a, _) = app();
    let mut h = common::harness_with(common::builder().with_size(egui::vec2(1200., 800.)), a, 4);
    h.get_by_label("Agree to continue").click();
    h.run_steps(12);
    assert_eq!(value(h.state(), "D1"), Value::Number(1.));
    assert_eq!(value(h.state(), "D2"), Value::Bool(true));
    assert_eq!(value(h.state(), "D3"), Value::Text("Agree".into()));
    // Native keyboard focus activates once without a grid-cell edit.
    h.get_by_label("Agree to continue").focus();
    h.run_steps(2);
    for pressed in [true, false] {
        h.input_mut().events.push(egui::Event::Key {
            key: egui::Key::Space,
            physical_key: None,
            pressed,
            repeat: false,
            modifiers: egui::Modifiers::NONE,
        });
        h.run_steps(3);
    }
    h.run_steps(12);
    assert_eq!(value(h.state(), "D1"), Value::Number(2.));
    assert_eq!(value(h.state(), "A1"), Value::Bool(false));
}
#[test]
fn file_controls_are_inert_before_consent_and_qualified_targets_are_checked_before_mutation() {
    let (a, _) = app();
    let bytes = gridcraft_xlsx::write_xlsx(&a.session.doc().unwrap().wb).unwrap();
    let (wb, _) = gridcraft_xlsx::read_xlsx(&bytes).unwrap();
    let mut s = Session::new();
    s.add_document(DocState::new(wb, None, "Authored.xlsm".into(), Trust::FromFile));
    let mut a = common::app(s);
    let id = json!(a.session.doc().unwrap().wb.sheet(0).unwrap().shapes[0].id);
    let before = a.session.doc().unwrap().wb.clone();
    assert!(a.run("vba.control", json!({"id":id})).is_err());
    assert_eq!(a.session.doc().unwrap().wb, before);
    a.run("vba.trust", json!({"allow":true})).unwrap();
    for entry in ["[2]!Hit", "'Other.xlsm'!Hit", "Missing.Hit", "Hidden"] {
        a.run("shape.setMacro", json!({"id":id,"entry":entry})).unwrap();
        let before = a.session.doc().unwrap().wb.clone();
        assert!(a.run("vba.control", json!({"id":id})).is_err(), "{entry}");
        assert_eq!(a.session.doc().unwrap().wb, before);
    }
    a.run("shape.setMacro", json!({"id":id,"entry":"'Authored.xlsm'!Hit"})).unwrap();
    a.run("vba.control", json!({"id":id})).unwrap();
    finish(&mut a);
    assert_eq!(value(&a, "D1"), Value::Number(1.));
}
#[test]
fn break_reset_design_disabled_and_missing_owners_never_queue_a_second_click() {
    let (mut a, id) = app();
    let at = SOURCE.find("Range(\"D1\")").unwrap();
    a.vba.as_mut().unwrap().ide.entry = "Hit".into();
    a.vba.as_mut().unwrap().ide.toggle_breakpoint_at(at).unwrap();
    a.run("vba.control", json!({"id":id})).unwrap();
    for _ in 0..15 {
        gridcraft_ui_egui::vba::tick(&mut a, &egui::Context::default());
    }
    assert_eq!(a.vba.as_ref().unwrap().ide.mode(), Mode::Break);
    let before = a.session.doc().unwrap().wb.clone();
    assert!(a.run("vba.control", json!({"id":id})).is_err());
    assert_eq!(a.session.doc().unwrap().wb, before);
    a.run("vba.run", json!({})).unwrap();
    finish(&mut a);
    assert_eq!(value(&a, "D1"), Value::Number(1.));
    a.run("vba.reset", json!({})).unwrap();
    a.run("vba.designMode", json!({"on":true})).unwrap();
    assert!(a.run("vba.control", json!({"id":id})).is_err());
    a.run("vba.designMode", json!({"on":false})).unwrap();
    a.run("control.set", json!({"id":id,"enabled":false})).unwrap();
    assert!(a.run("vba.control", json!({"id":id})).is_err());
    a.run("object.delete", json!({"kind":"shape","id":id})).unwrap();
    assert!(a.run("vba.control", json!({"id":id})).is_err());
}
#[test]
fn every_family_has_a_visible_accessible_widget_and_list_input_updates_link() {
    let mut s = Session::new();
    s.new_workbook();
    let mut a = common::app(s);
    let mut ids = vec![];
    for (i, kind) in [0, 1, 2, 4, 5, 6, 7, 8, 9].into_iter().enumerate() {
        let id=a.run("control.add",json!({"type":kind,"caption":format!("Native {kind}"),"left":80+(i%3)*240,"top":40+(i/3)*150,"width":190,"height":if [3,4,6].contains(&kind){100}else{30}})).unwrap()["shape"].clone();
        if [2, 6].contains(&kind) {
            a.run(
                "control.set",
                json!({"id":id,"items":if kind==2{vec!["Drop A","Drop B"]}else{vec!["Alpha","Beta"]},"linkedCell":if kind==2{"G1"}else{"G2"}}),
            )
            .unwrap();
        }
        ids.push(id);
    }
    let mut h = common::harness_with(common::builder().with_size(egui::vec2(1440., 1000.)), a, 4);
    h.get_by_label("Beta").click();
    h.run_steps(4);
    assert_eq!(value(h.state(), "G2"), Value::Number(2.));
    h.get_by_label("Native 1").click();
    h.run_steps(4);
    let sh = h.state().session.doc().unwrap().wb.sheet(0).unwrap();
    assert_eq!(sh.shapes[1].format.form_control.as_ref().unwrap().value, 1);
    h.get_by_label("Native 2").click();
    h.run_steps(2);
    h.get_by_label("Drop B").click();
    h.run_steps(4);
    assert_eq!(value(h.state(), "G1"), Value::Number(2.));
    h.get_by_label("Native 7").click();
    h.run_steps(4);
    assert_eq!(h.state().session.doc().unwrap().wb.sheet(0).unwrap().shapes[6].format.form_control.as_ref().unwrap().value, 1);
    h.get_by_label("▲").click();
    h.run_steps(4);
    assert_eq!(h.state().session.doc().unwrap().wb.sheet(0).unwrap().shapes[8].format.form_control.as_ref().unwrap().value, 1);
    h.get_by_label("Native 8").focus();
    h.run_steps(2);
    for pressed in [true, false] {
        h.input_mut().events.push(egui::Event::Key {
            key: egui::Key::PageDown,
            physical_key: None,
            pressed,
            repeat: false,
            modifiers: egui::Modifiers::NONE,
        });
        h.run_steps(2);
    }
    assert_eq!(h.state().session.doc().unwrap().wb.sheet(0).unwrap().shapes[7].format.form_control.as_ref().unwrap().value, 10);
    assert_eq!(ids.len(), 9);
}

#[test]
fn native_button_package_click_after_consent_and_paused_link_edits_use_current_state() {
    let (mut a, _) = app();
    let button = a
        .run("control.add", json!({"type":0,"name":"ReportButton","caption":"Run native report","left":420,"top":70,"width":170,"height":30}))
        .unwrap()["shape"]
        .clone();
    a.run("shape.setMacro", json!({"id":button,"entry":"[0]!Hit"})).unwrap();
    let (wb, _) = gridcraft_xlsx::read_xlsx(&gridcraft_xlsx::write_xlsx(&a.session.doc().unwrap().wb).unwrap()).unwrap();
    let mut s = Session::new();
    s.add_document(DocState::new(wb, None, "Report.xlsm".into(), Trust::FromFile));
    let a = common::app(s);
    let mut h = common::harness_with(common::builder().with_size(egui::vec2(1200., 800.)), a, 4);
    assert_eq!(h.state().session.doc().unwrap().wb.sheet(0).unwrap().shapes.len(), 2);
    h.get_by_label("Run native report").click();
    h.run_steps(4);
    assert_eq!(value(h.state(), "D1"), Value::Empty);
    h.get_by_label("  OK  ").click();
    h.run_steps(2);
    h.state_mut().run("vba.trust", json!({"allow":true})).unwrap();
    h.get_by_label("Run native report").click();
    h.run_steps(12);
    assert_eq!(value(h.state(), "D1"), Value::Number(1.));
    assert_eq!(value(h.state(), "D3"), Value::Text("ReportButton".into()));
    let editor = h.state_mut().vba.as_mut().unwrap();
    editor.ide.entry = "Hit".into();
    editor.ide.toggle_breakpoint_at(SOURCE.find("Range(\"D1\")").unwrap()).unwrap();
    let check = json!(
        h.state().session.doc().unwrap().wb.sheet(0).unwrap().shapes.iter().find(|s| s.format.office_name.as_deref() == Some("Agree")).unwrap().id
    );
    h.state_mut().run("vba.control", json!({"id":check})).unwrap();
    h.run_steps(6);
    assert_eq!(h.state().vba.as_ref().unwrap().ide.mode(), Mode::Break);
    h.state_mut().run("cell.set", json!({"cell":"A1","input":"FALSE"})).unwrap();
    h.state_mut().run("vba.run", json!({})).unwrap();
    h.run_steps(12);
    assert_eq!(value(h.state(), "D2"), Value::Bool(false));
}

#[test]
fn activation_keeps_its_captured_sheet_when_another_sheet_is_active() {
    let (mut a, id) = app();
    a.run("home.insertSheet", json!({"name":"Other"})).unwrap();
    a.run("vba.control", json!({"sheet":0,"id":id})).unwrap();
    finish(&mut a);
    assert_eq!(value(&a, "D1"), Value::Number(1.));
    let other = a.session.doc().unwrap().wb.sheet(1).unwrap();
    assert_eq!(other.value(CellRef::parse("D1").unwrap()), Value::Empty);
}
