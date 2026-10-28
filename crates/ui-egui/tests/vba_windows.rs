#![cfg(all(feature = "vba-window", not(target_arch = "wasm32")))]
use egui::{Context, RawInput, ViewportCommand, ViewportEvent, ViewportId, ViewportInfo};
use gridcraft_engine::Session;
use gridcraft_ui_egui::{SheetApp, doc_windows::DocWindows, vba};
use serde_json::json;
use std::cell::RefCell;
thread_local! {
    static INPUT:RefCell<egui::ViewportIdMap<ViewportInfo>>=RefCell::default();
    static DRAWS:RefCell<Vec<ViewportId>>=const{RefCell::new(Vec::new())};
    static COMMANDS:RefCell<Vec<(ViewportId,ViewportCommand)>>=const{RefCell::new(Vec::new())};
}
fn record(mut out: egui::FullOutput) {
    out.textures_delta.clear();
    COMMANDS.with(|s| s.borrow_mut().extend(out.viewport_output.into_iter().flat_map(|(id, v)| v.commands.into_iter().map(move |c| (id, c)))));
}
struct Desk {
    ctx: Context,
    app: SheetApp,
    windows: DocWindows,
}
impl Desk {
    fn new() -> Self {
        INPUT.with(|i| i.borrow_mut().clear());
        let ctx = Context::default();
        ctx.set_embed_viewports(false);
        ctx.set_fonts(gridcraft_ui_egui::theme::pinned_font_definitions());
        Context::set_immediate_viewport_renderer(|ctx, viewport| {
            let id = viewport.ids.this;
            DRAWS.with(|s| s.borrow_mut().push(id));
            let mut input = RawInput { viewport_id: id, ..Default::default() };
            input.viewports = INPUT.with(|s| s.borrow().clone());
            input.viewports.entry(id).or_default().parent = Some(viewport.ids.parent);
            let mut callback = viewport.viewport_ui_cb;
            record(ctx.run_ui(input, |ui| callback(ui)));
        });
        let mut session = Session::new();
        session.new_workbook();
        let mut app = SheetApp::new(session, Default::default());
        app.use_prepared_fonts();
        let windows = DocWindows::new(&mut app, Default::default(), true);
        Self { ctx, app, windows }
    }
    fn frame(&mut self) -> Vec<ViewportId> {
        DRAWS.with(|s| s.borrow_mut().clear());
        COMMANDS.with(|s| s.borrow_mut().clear());
        let mut input = RawInput::default();
        input.viewports = INPUT.with(|s| s.borrow().clone());
        input.viewports.entry(ViewportId::ROOT).or_default();
        let out = self.ctx.clone().run_ui(input, |ui| {
            let ctx = ui.ctx().clone();
            self.app.logic(&ctx);
            self.windows.logic(&mut self.app, &ctx);
            self.windows.show(&mut self.app, ui);
            let before = self.app.session.journal.len();
            vba::show_desktop(&mut self.app, &ctx, self.windows.focused_viewport());
            assert_eq!(before, self.app.session.journal.len(), "rendering the child must not advance execution");
        });
        record(out);
        INPUT.with(|s| s.borrow_mut().values_mut().for_each(|v| v.events.clear()));
        DRAWS.with(|s| s.borrow().clone())
    }
}
#[test]
fn native_editor_has_one_separate_viewport_and_keeps_the_grid_live() {
    let mut d = Desk::new();
    d.app.run("vba.open", json!({"source":gridcraft_vba::DEMO})).unwrap();
    assert!(d.app.vba.as_ref().unwrap().separate_window);
    let ids = d.frame();
    assert_eq!(ids.len(), 1);
    let id = ids[0];
    assert_ne!(id, ViewportId::ROOT);
    assert!(d.app.grid.rect.is_some());
    assert!(d.app.vba.as_ref().unwrap().view.report.is_some());
    d.app.run("vba.open", json!({})).unwrap();
    assert_eq!(d.frame(), [id]);
    let byte = gridcraft_vba::DEMO.find("Range(\"B1\")").unwrap();
    d.app.vba.as_mut().unwrap().ide.toggle_breakpoint_at(byte).unwrap();
    d.app.run("vba.invoke", json!({"entry":"Main"})).unwrap();
    assert!(!d.app.vba.as_ref().unwrap().visible);
    assert_eq!(d.frame(), [id]);
    assert_eq!(d.app.vba.as_ref().unwrap().ide.mode(), gridcraft_vba::Mode::Break);
    assert!(d.app.vba.as_ref().unwrap().visible);
    assert!(!d.app.session.doc().unwrap().wb.active().unwrap().cells.is_empty());
}
#[test]
fn window_close_hides_editor_preserves_dirty_source_and_focuses_workbook() {
    let mut d = Desk::new();
    d.app.run("vba.open", json!({})).unwrap();
    let id = d.frame()[0];
    let text = "Sub KeepMe()\nDebug.Print 7\nEnd Sub\n";
    d.app.vba.as_mut().unwrap().ide.set_text(text.into()).unwrap();
    INPUT.with(|s| s.borrow_mut().entry(id).or_default().events.push(ViewportEvent::Close));
    d.frame();
    assert!(!d.app.vba.as_ref().unwrap().visible);
    assert!(d.app.vba.as_ref().unwrap().ide.modules[0].dirty());
    assert_eq!(d.app.session.documents().len(), 1);
    COMMANDS.with(|s| {
        let c = s.borrow();
        assert!(c.iter().any(|(vp, c)| *vp == id && matches!(c, ViewportCommand::CancelClose)));
        assert!(c.iter().any(|(vp, c)| *vp == ViewportId::ROOT && matches!(c, ViewportCommand::Focus)));
    });
    assert!(d.frame().is_empty());
    assert!(d.app.run("file.close", json!({"force":true})).is_err());
    d.app.run("vba.open", json!({})).unwrap();
    assert_eq!(d.frame(), [id]);
    assert_eq!(d.app.vba.as_ref().unwrap().ide.modules[0].text, text);
}
#[test]
fn embedded_launch_override_and_switching_share_one_project() {
    let mut d = Desk::new();
    d.app.run("vba.open", json!({"separateWindow":false})).unwrap();
    assert!(!d.app.vba.as_ref().unwrap().separate_window);
    assert!(d.frame().is_empty());
    let document = d.app.vba.as_ref().unwrap().bridge.document();
    d.app.run("vba.open", json!({"separateWindow":true})).unwrap();
    assert_eq!(d.frame().len(), 1);
    assert_eq!(d.app.vba.as_ref().unwrap().bridge.document(), document);
    d.app.run("vba.open", json!({"separateWindow":false})).unwrap();
    assert!(d.frame().is_empty());
    assert_eq!(d.app.vba.as_ref().unwrap().bridge.document(), document);
}
#[test]
fn rendering_a_running_editor_does_not_execute_an_additional_slice() {
    let mut d = Desk::new();
    d.app.run("vba.open", json!({"source":"Sub Main()\nDim i As Long\nFor i = 1 To 100000\nCells(1,1).Value2 = i\nNext i\nEnd Sub\n"})).unwrap();
    d.app.run("vba.run", json!({})).unwrap();
    // frame asserts that all workbook effects occur before child viewport rendering.
    assert_eq!(d.frame().len(), 1);
    assert_eq!(d.app.vba.as_ref().unwrap().ide.mode(), gridcraft_vba::Mode::Running);
    assert!(!d.app.session.journal.is_empty());
    d.app.run("vba.reset", json!({})).unwrap();
}
