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
#[test]
fn control_commands_open_without_execution_run_and_close() {
    let mut app = app();
    app.run("vba.open", json!({})).unwrap();
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
        app.run("vba.open", json!({})).unwrap();
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
