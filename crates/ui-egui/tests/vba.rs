#![cfg(feature = "vba")]
#[allow(dead_code)]
mod common;
use egui_kittest::kittest::Queryable as _;
use gridcraft_engine::{DocState, Session, Trust, model::Workbook};
use gridcraft_ui_egui::SheetApp;
use gridcraft_vba::{Action, Mode};
use serde_json::json;

fn app() -> SheetApp {
    let mut session = Session::new();
    session.new_workbook();
    common::app(session)
}
const WORKSHEET_SOURCE: &str = "Public Sub Fill()\nRange(\"A1\").Value2 = 17\nEnd Sub\nPublic Function Triple(ByVal x As Double) As Double\nTriple = x * 3\nEnd Function\nPrivate Sub Hidden()\nEnd Sub\n";
fn worksheet_app() -> SheetApp {
    let mut a = app();
    a.run("vba.open", json!({"source":WORKSHEET_SOURCE})).unwrap();
    a
}
fn cell(app: &SheetApp, at: &str) -> gridcraft_engine::core::Value {
    app.session.doc().unwrap().wb.active().unwrap().value(gridcraft_engine::core::CellRef::parse(at).unwrap())
}
#[test]
fn publish_source_enables_formulas_and_revoke_or_close_disables_them() {
    use gridcraft_engine::core::{CellError, Value};
    let mut a = worksheet_app();
    a.run("vba.apply", json!({})).unwrap();
    a.run("cell.set", json!({"cell":"B1","input":"=Triple(4)"})).unwrap();
    assert_eq!(cell(&a, "B1"), Value::Number(12.));
    assert!(!a.vba.as_ref().unwrap().ide.modules[0].dirty());
    a.run("vba.close", json!({})).unwrap();
    assert_eq!(cell(&a, "B1"), Value::Error(CellError::Name));
    a.run("vba.open", json!({})).unwrap();
    assert_eq!(a.vba.as_ref().unwrap().ide.modules[0].text, WORKSHEET_SOURCE);
    a.run("vba.apply", json!({})).unwrap();
    assert_eq!(cell(&a, "B1"), Value::Number(12.));
}
#[test]
fn button_and_toolbar_assignment_are_persistent_undoable_and_public_only() {
    use gridcraft_vba::worksheet::Project;
    let mut a = worksheet_app();
    assert!(a.run("vba.assign", json!({"entry":"Hidden"})).is_err());
    assert!(a.run("vba.assign", json!({"entry":"Triple"})).is_err());
    assert!(a.run("vba.assign", json!({"entry":"Fill","id":u64::MAX})).is_err());
    let result = a.run("vba.assign", json!({"entry":"Fill","at":"D4","caption":"Fill values"})).unwrap();
    let id = result["shape"].as_u64().unwrap() as u32;
    assert_eq!(gridcraft_ui_egui::vba::assigned(&a, id).as_deref(), Some("Fill"));
    assert_eq!(a.session.doc().unwrap().undo.len(), 1);
    a.run("edit.undo", json!({})).unwrap();
    assert!(a.session.doc().unwrap().wb.active().unwrap().shapes.is_empty());
    assert!(Project::read(&a.session.doc().unwrap().wb).unwrap().is_none());
    a.run("edit.redo", json!({})).unwrap();
    a.run("vba.pin", json!({"entry":"Fill"})).unwrap();
    assert_eq!(Project::read(&a.session.doc().unwrap().wb).unwrap().unwrap().toolbar, ["Fill"]);
    a.run("cell.set", json!({"cell":"B1","input":"=Triple(A1)"})).unwrap();
    a.run("vba.invoke", json!({"entry":"Fill"})).unwrap();
    for _ in 0..20 {
        gridcraft_ui_egui::vba::tick(&mut a, &egui::Context::default());
    }
    assert_eq!(cell(&a, "A1"), gridcraft_engine::core::Value::Number(17.));
    assert_eq!(cell(&a, "B1"), gridcraft_engine::core::Value::Number(51.));
    a.run("vba.pin", json!({"entry":"Fill","remove":true})).unwrap();
    assert!(Project::read(&a.session.doc().unwrap().wb).unwrap().unwrap().toolbar.is_empty());
}
#[test]
fn file_source_is_inert_until_explicit_consent_and_revoke_recalculates() {
    use gridcraft_engine::core::{CellError, Value};
    let mut a = worksheet_app();
    a.run("vba.apply", json!({})).unwrap();
    let wb = (*a.session.doc().unwrap().wb).clone();
    let mut session = Session::new();
    session.add_document(DocState::new(wb, None, "File".into(), Trust::FromFile));
    let mut b = common::app(session);
    b.run("cell.set", json!({"cell":"B1","input":"=Triple(8)"})).unwrap();
    b.run("vba.open", json!({})).unwrap();
    assert_eq!(cell(&b, "B1"), Value::Error(CellError::Name));
    assert!(b.run("vba.invoke", json!({"entry":"Fill"})).is_err());
    b.run("vba.trust", json!({"allow":true})).unwrap();
    assert_eq!(cell(&b, "B1"), Value::Number(24.));
    b.run("vba.trust", json!({"allow":false})).unwrap();
    assert_eq!(cell(&b, "B1"), Value::Error(CellError::Name));
}
#[test]
fn toolbar_macro_picker_and_developer_controls_are_reachable() {
    let mut a = worksheet_app();
    a.run("vba.return", json!({})).unwrap();
    a.ui.ribbon_tab = "Developer".into();
    a.ui.developer = true;
    let mut h = common::harness_with(common::builder().with_size(egui::vec2(1440., 900.)), a, 4);
    assert!(h.query_by_label("Visual Basic").is_some());
    h.get_by_label("Macros (Alt+F8)").click();
    h.run_steps(2);
    h.get_by_label("GridCraft customizations").click();
    h.run_steps(3);
    h.get_by_label("Add to Toolbar").click();
    h.run_steps(3);
    h.get_by_label("Assign Button").click();
    h.run_steps(3);
    assert_eq!(h.state().session.doc().unwrap().wb.active().unwrap().shapes.len(), 1);
    h.state_mut().open_dialog("macros", json!({}));
    h.run_steps(3);
    h.get_by_label("Run").click();
    h.run_steps(8);
    assert!(h.state().dialog.is_none());
    assert_eq!(cell(h.state(), "A1"), gridcraft_engine::core::Value::Number(17.));
    assert!(h.query_by_label("VB").is_none());
    h.get_by_label("Visual Basic").click();
    h.run_steps(3);
    assert!(h.state().vba.as_ref().unwrap().visible);
}
#[test]
fn worksheet_button_click_runs_and_design_mode_selects() {
    use egui::{Event, Modifiers, PointerButton, vec2};
    use gridcraft_engine::core::{CellRef, Value};
    let mut a = worksheet_app();
    let id = a.run("vba.assign", json!({"entry":"Fill","at":"D4"})).unwrap()["shape"].as_u64().unwrap() as u32;
    a.run("vba.return", json!({})).unwrap();
    let mut h = common::harness_with(common::builder().with_size(vec2(1440., 900.)), a, 4);
    let app = h.state();
    let sh = app.session.doc().unwrap().wb.active().unwrap();
    let geo = gridcraft_ui_egui::grid::Geo::new(sh, app.grid.rect.unwrap(), app.view().scroll);
    let pos = geo.cell_rect(sh, CellRef::parse("D4").unwrap()).min + vec2(30., 12.);
    for pressed in [true, false] {
        h.input_mut()
            .events
            .extend([Event::PointerMoved(pos), Event::PointerButton { pos, button: PointerButton::Primary, pressed, modifiers: Modifiers::NONE }]);
        h.run_steps(3);
    }
    h.run_steps(8);
    assert_eq!(cell(h.state(), "A1"), Value::Number(17.));
    h.state_mut().run("cell.set", json!({"cell":"A1","input":"0"})).unwrap();
    h.state_mut().run("vba.designMode", json!({"on":true})).unwrap();
    for pressed in [true, false] {
        h.input_mut()
            .events
            .extend([Event::PointerMoved(pos), Event::PointerButton { pos, button: PointerButton::Primary, pressed, modifiers: Modifiers::NONE }]);
        h.run_steps(3);
    }
    assert_eq!(cell(h.state(), "A1"), Value::Number(0.));
    assert_eq!(h.state().selected_chart, Some(id));
    for pressed in [true, false] {
        h.input_mut()
            .events
            .extend([Event::PointerMoved(pos), Event::PointerButton { pos, button: PointerButton::Secondary, pressed, modifiers: Modifiers::NONE }]);
        h.run_steps(3);
    }
    h.get_by_label("Assign Macro…").click();
    h.run_steps(3);
    assert_eq!(h.state().dialog.as_ref().unwrap().values["shape"], id);
    h.get_by_label("Assign").click();
    h.run_steps(3);
    assert_eq!(h.state().session.doc().unwrap().wb.active().unwrap().shapes.len(), 1);
}
#[test]
fn stale_editor_cannot_publish_or_assign_into_another_workbook() {
    let mut a = worksheet_app();
    a.session.new_workbook();
    let wb = a.session.doc().unwrap().wb.clone();
    assert!(a.run("vba.apply", json!({})).is_err());
    assert!(a.run("vba.assign", json!({"entry":"Fill"})).is_err());
    assert_eq!(a.session.doc().unwrap().wb, wb);
    assert!(a.session.doc().unwrap().undo.is_empty());
}
#[test]
fn failed_publish_keeps_previous_workbook_source_and_functions() {
    let mut a = worksheet_app();
    a.run("vba.apply", json!({})).unwrap();
    a.run("cell.set", json!({"cell":"B1","input":"=Triple(3)"})).unwrap();
    let wb = a.session.doc().unwrap().wb.clone();
    a.vba.as_mut().unwrap().ide.set_text("Function Triple(\n".into()).unwrap();
    assert!(a.run("vba.apply", json!({})).is_err());
    assert_eq!(a.session.doc().unwrap().wb, wb);
    assert_eq!(cell(&a, "B1"), gridcraft_engine::core::Value::Number(9.));
}
#[test]
fn deleting_an_assigned_shape_does_not_assign_its_reused_id() {
    let mut a = worksheet_app();
    let id = a.run("vba.assign", json!({"entry":"Fill"})).unwrap()["shape"].as_u64().unwrap() as u32;
    a.run("object.delete", json!({"kind":"shape","id":id})).unwrap();
    let new_id = a.run("insert.shape", json!({"text":"Unassigned"})).unwrap()["shape"].as_u64().unwrap() as u32;
    assert_eq!(new_id, id);
    assert!(gridcraft_ui_egui::vba::assigned(&a, new_id).is_none());
    a.run("edit.undo", json!({})).unwrap();
    a.run("edit.undo", json!({})).unwrap();
    assert_eq!(gridcraft_ui_egui::vba::assigned(&a, id).as_deref(), Some("Fill"));
}
#[test]
fn assigned_macro_without_source_or_with_corrupt_binary_never_runs_the_demo() {
    let mut a = app();
    assert!(a.run("vba.invoke", json!({"entry":"Main"})).is_err());
    assert!(a.vba.is_none());
    let mut wb = Workbook::new();
    wb.vba_project = Some(std::sync::Arc::new(vec![1, 2, 3]));
    let mut session = Session::new();
    session.add_document(DocState::new(wb, None, "Invalid project".into(), Trust::Local));
    let mut a = common::app(session);
    assert!(a.run("vba.invoke", json!({"entry":"Main"})).is_err());
    assert!(a.vba.is_none());
    assert_eq!(cell(&a, "A1"), gridcraft_engine::core::Value::Empty);
}
#[test]
fn keyboard_opens_toggles_editor_and_macro_picker_and_developer_preference_persists() {
    let a = app();
    assert!(!a.ui.developer);
    let mut h = common::harness_with(common::builder().with_size(egui::vec2(1440., 900.)), a, 4);
    let key = |h: &mut common::Harness, key| {
        for pressed in [true, false] {
            h.input_mut().events.push(egui::Event::Key { key, physical_key: None, pressed, repeat: false, modifiers: egui::Modifiers::ALT });
            h.run_steps(2);
        }
    };
    key(&mut h, egui::Key::F11);
    assert!(h.state().vba.as_ref().unwrap().visible);
    key(&mut h, egui::Key::F11);
    assert!(!h.state().vba.as_ref().unwrap().visible);
    key(&mut h, egui::Key::F8);
    assert_eq!(h.state().dialog.as_ref().unwrap().name, "macros");
    h.state_mut().run("vba.showDeveloper", json!({"on":true})).unwrap();
    let encoded = serde_json::to_value(&h.state().ui).unwrap();
    assert_eq!(encoded["developer"], true);
    let restored: gridcraft_ui_egui::UiState = serde_json::from_value(encoded).unwrap();
    assert!(restored.developer);
}
#[test]
fn normal_launch_has_no_demo_and_window_requests_require_native_capability() {
    let mut a = app();
    assert!(a.run("vba.open", json!({"separateWindow":true})).is_err());
    assert!(a.vba.is_none());
    a.run("vba.open", json!({"separateWindow":false})).unwrap();
    assert!(!a.vba.as_ref().unwrap().separate_window);
    assert!(a.run("vba.entries", json!({})).unwrap().as_array().unwrap().is_empty());
    assert!(!a.vba.as_ref().unwrap().ide.modules[0].text.contains("Sub Main"));
}
#[test]
fn view_menu_and_options_customize_ribbon_are_the_launch_routes() {
    let a = app();
    let mut h = common::harness_with(common::builder().with_size(egui::vec2(1440., 900.)), a, 4);
    assert!(h.query_by_label("VB").is_none());
    h.state_mut().ui.ribbon_tab = "View".into();
    h.run_steps(2);
    h.get_by_label("Macros menu").click();
    h.run_steps(2);
    h.get_by_label("View Macros").click();
    h.run_steps(3);
    assert_eq!(h.state().dialog.as_ref().unwrap().name, "macros");
    h.state_mut().dialog = None;
    h.state_mut().run("ui.backstage", json!({"page":"options"})).unwrap();
    h.run_steps(3);
    h.get_by_label("Developer").click();
    h.run_steps(3);
    assert!(h.state().ui.developer);
}
#[test]
fn macro_edit_and_step_into_navigate_to_the_selected_module_without_listing_functions() {
    let mut a = worksheet_app();
    a.vba.as_mut().unwrap().ide.add_module("Other.bas", "Public Sub SecondOne()\nDebug.Print 8\nEnd Sub\n").unwrap();
    a.run("vba.return", json!({})).unwrap();
    a.open_dialog("macros", json!({}));
    let mut h = common::harness_with(common::builder().with_size(egui::vec2(1440., 900.)), a, 4);
    assert!(h.query_by_label("Triple").is_none());
    h.get_by_label("SecondOne").click();
    h.run_steps(2);
    h.get_by_label("Edit").click();
    h.run_steps(3);
    let editor = h.state().vba.as_ref().unwrap();
    assert!(editor.visible);
    assert_eq!(editor.ide.entry, "SecondOne");
    assert_eq!(editor.ide.modules[editor.ide.active].name, "Other.bas");
    h.state_mut().run("vba.return", json!({})).unwrap();
    h.state_mut().open_dialog("macros", json!({"macro":"SecondOne"}));
    h.run_steps(3);
    h.get_by_label("Step Into").click();
    h.run_steps(3);
    assert_eq!(h.state().vba.as_ref().unwrap().ide.mode(), Mode::Break);
    assert_eq!(h.state().vba.as_ref().unwrap().ide.entry, "SecondOne");
}
#[test]
fn control_commands_open_without_execution_run_and_close() {
    let mut app = app();
    app.run("vba.open", json!({"source":gridcraft_vba::DEMO})).unwrap();
    assert_eq!(app.vba.as_ref().unwrap().ide.mode(), Mode::Design);
    assert_eq!(app.session.doc().unwrap().undo.len(), 0);
    app.run("vba.run", json!({})).unwrap();
    let ctx = egui::Context::default();
    for _ in 0..30 {
        gridcraft_ui_egui::vba::tick(&mut app, &ctx);
    }
    assert_eq!(app.run("vba.status", json!({})).unwrap()["mode"], "Completed");
    assert_eq!(app.session.doc().unwrap().undo.len(), 1);
    app.run("vba.close", json!({})).unwrap();
    assert!(app.vba.is_none());
}
#[test]
fn untrusted_pure_code_requires_explicit_consent_and_revoke_stops_it() {
    let mut session = Session::new();
    session.add_document(DocState::new(Workbook::new(), None, "File".into(), Trust::FromFile));
    let mut app = common::app(session);
    app.run("vba.open", json!({"source":"Sub Main()\nDim i As Long\nFor i = 1 To 1000000\nNext i\nEnd Sub\n"})).unwrap();
    assert!(app.run("vba.run", json!({})).is_err());
    assert!(app.run("vba.trust", json!({"allow":"yes"})).is_err());
    app.run("vba.trust", json!({"allow":true})).unwrap();
    app.run("vba.run", json!({})).unwrap();
    gridcraft_ui_egui::vba::tick(&mut app, &egui::Context::default());
    app.run("vba.trust", json!({"allow":false})).unwrap();
    gridcraft_ui_egui::vba::tick(&mut app, &egui::Context::default());
    assert_eq!(app.vba.as_ref().unwrap().ide.mode(), Mode::Design);
    assert!(!app.session.in_group(app.session.doc().unwrap().uid));
}
#[test]
fn edited_source_requires_explicit_discard_and_failed_compile_releases_group() {
    let mut app = app();
    app.run("vba.open", json!({})).unwrap();
    app.vba.as_mut().unwrap().ide.set_text("Sub Main()\ninvalid syntax @\nEnd Sub\n".into()).unwrap();
    assert!(app.run("vba.close", json!({})).is_err());
    assert!(app.run("file.close", json!({"force":true})).is_err());
    assert!(app.run("vba.run", json!({})).is_err());
    assert!(!app.session.in_group(app.session.doc().unwrap().uid));
    app.run("vba.close", json!({"discard":true})).unwrap();
    assert!(app.vba.is_none());
}
#[test]
fn shared_editor_accessible_in_light_dark_narrow_and_scaled_hosts() {
    for (dark, width, scale) in [(false, 1200., 1.), (true, 1200., 1.), (false, 640., 1.5)] {
        let mut app = app();
        app.ui.dark = dark;
        app.ui.system_theme = false;
        app.run("vba.open", json!({"source":gridcraft_vba::DEMO})).unwrap();
        app.run("vba.stepInto", json!({})).unwrap();
        let mut h = common::harness_with(common::builder().with_size(egui::vec2(width, 850.)).with_pixels_per_point(scale), app, 6);
        assert_eq!(h.state().vba.as_ref().unwrap().ide.mode(), Mode::Break);
        assert!(h.query_by_label("Return to workbook").is_some());
        let report = h.state().vba.as_ref().unwrap().view.report.as_ref().unwrap();
        assert!(report.editor.width() > 100. && report.editor.height() > 80.);
        h.get_by_label("Return to workbook").click();
        h.run_steps(2);
        assert!(!h.state().vba.as_ref().unwrap().visible);
        h.state_mut().run("vba.reset", json!({})).unwrap();
        assert_eq!(h.state().vba.as_ref().unwrap().ide.mode(), Mode::Design);
        assert!(h.state().vba.as_ref().unwrap().ide.enabled(Action::Run));
    }
}

#[test]
fn close_prompt_preserves_source_until_explicit_discard() {
    let mut app = app();
    app.run("vba.open", json!({})).unwrap();
    app.vba.as_mut().unwrap().ide.set_text("Sub Main()\nDebug.Print 7\nEnd Sub\n".into()).unwrap();
    let mut h = common::harness_with(common::builder().with_size(egui::vec2(1200., 850.)), app, 4);
    h.get_by_label("Close project").click();
    h.run_steps(2);
    assert!(h.query_by_label("Discard source changes").is_some());
    assert!(h.state().vba.is_some());
    h.get_by_label("Keep editing").click();
    h.run_steps(2);
    assert!(h.state().vba.as_ref().unwrap().ide.modules[0].dirty());
    h.get_by_label("Close project").click();
    h.run_steps(2);
    h.get_by_label("Discard source changes").click();
    h.run_steps(2);
    assert!(h.state().vba.is_none());
}

#[test]
fn shared_save_exports_exact_source_through_host_service() {
    let mut app = app();
    let saved = std::sync::Arc::new(std::sync::Mutex::new(Vec::<u8>::new()));
    let output = saved.clone();
    app.services.download = Some(Box::new(move |_name, bytes| {
        output.lock().unwrap().extend(bytes);
    }));
    app.run("vba.open", json!({})).unwrap();
    let text = "Sub Main()\r\nDebug.Print 8\r\nEnd Sub\r\n";
    let editor = app.vba.as_mut().unwrap();
    editor.ide.set_text(text.into()).unwrap();
    editor.bridge.dispatch(&mut app.session, &mut editor.ide, Action::Save).unwrap();
    let h = common::harness_with(common::builder().with_size(egui::vec2(1200., 850.)), app, 4);
    assert_eq!(*saved.lock().unwrap(), text.as_bytes());
    assert!(!h.state().vba.as_ref().unwrap().ide.modules[0].dirty());
}

#[test]
fn vba_commands_are_discoverable_and_use_the_authenticated_control_route() {
    use gridcraft_ui_egui::control::{self, Outcome};
    let mut app = app();
    let ctx = egui::Context::default();
    let Outcome::Done(reply) = control::handle(&mut app, &ctx, "engine.commands", &json!({"search":"vba."})) else { panic!("command result") };
    let commands = reply["result"].as_array().unwrap();
    assert!(commands.iter().any(|c| c["id"] == "vba.open" && c["enabled"] == true));
    let Outcome::Done(reply) = control::handle(&mut app, &ctx, "engine.execute", &json!({"command":"vba.open","params":{}})) else {
        panic!("open result")
    };
    assert_eq!(reply["ok"], true);
    assert_eq!(app.vba.as_ref().unwrap().ide.mode(), Mode::Design);
}

#[test]
fn clean_idle_editor_does_not_block_closing_its_workbook() {
    let mut app = app();
    app.run("vba.open", json!({})).unwrap();
    assert_eq!(app.vba.as_ref().unwrap().ide.mode(), Mode::Design);
    app.run("file.close", json!({"force":true,"document":true})).unwrap();
    assert!(app.session.documents().is_empty());
}
